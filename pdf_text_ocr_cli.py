import os
import sys
import io
import platform
import re
import fitz                   # PyMuPDF
import pytesseract
from PIL import Image, ImageOps

# =========================
# 0. tessdata_best 경로 설정
# =========================
# tessdata_best 안에 kor.traineddata, (eng.traineddata, osd.traineddata 있어도 상관 없음)
os.environ["TESSDATA_PREFIX"] = "/Users/kim_jinwoong/Desktop/project/프로젝트_개인/Extract_Text_from_PDF(img)/tessdata_best"

# =========================
# 헬퍼함수 추가
# =========================
def force_heads_to_newline(text: str) -> str:
    """
    본문 안에 섞여 있는 머리표들(1. / (1) / (가) / 가. / *) 앞에
    강제로 줄바꿈(\n)을 넣어, 항상 줄 맨 앞에 오도록 만든다.
    """
    patterns = [
        r'\s+(\(\d+\))',      # ... (1)
        r'\s+(\([가-힣]\))',  # ... (가)
        r'\s+([가-힣]\.)',    # ... 가.
        r'\s+(\d+\.)',        # ... 1.
        r'\s+(\*)',           # ... *
    ]
    for pat in patterns:
        text = re.sub(pat, r'\n\1', text)
    return text

# =========================
# 노이즈 지우기
# =========================
def clean_noise(text: str) -> str:
    """
    OCR 및 PDF 텍스트에서 공통으로 노이즈 제거:
    - '|' 제거
    - 영문 알파벳(A~Z, a~z) 제거
    """
    # 1) '|' 완전 제거
    text = text.replace("|", "")

    # 2) 영문 알파벳 제거
    text = re.sub(r'[A-Za-z]+', ' ', text)

    return text

# =========================
# 1. Tesseract 실행 파일 경로 설정
# =========================
def init_tesseract_path():
    system = platform.system()

    # PyInstaller나 exe 기준 base_dir
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(sys.argv[0])))

    # 1) exe 옆 tesseract 폴더 우선
    bundled_tesseract = os.path.join(base_dir, "tesseract", "tesseract.exe")
    if os.path.exists(bundled_tesseract):
        pytesseract.pytesseract.tesseract_cmd = bundled_tesseract
        print(f"[INFO] 번들된 Tesseract 사용: {bundled_tesseract}")
        return

    # 2) OS별 후보 경로
    if system == "Windows":
        candidates = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        ]
    elif system == "Darwin":  # macOS
        candidates = [
            "/opt/homebrew/bin/tesseract",
            "/usr/local/bin/tesseract",
        ]
    else:  # Linux 등
        candidates = ["tesseract"]

    for path in candidates:
        if path == "tesseract" or os.path.exists(path):
            pytesseract.pytesseract.tesseract_cmd = path
            print(f"[INFO] Tesseract 경로 설정: {path}")
            return

    print("[WARN] Tesseract 실행 파일을 찾지 못했습니다. PATH에 있는 tesseract를 사용합니다.")


init_tesseract_path()


# =========================
# 2. 문단 머리표 패턴 정의
# =========================
HEAD_PATTERNS = [
    re.compile(r'^\s*\d+\.'),        # 1. 2. 3. ...
    re.compile(r'^\s*\(\d+\)'),      # (1) (2) ...
    re.compile(r'^\s*\([가-힣]\)'),   # (가) (나) ...
    re.compile(r'^\s*[가-힣]\.'),     # 가. 나. ...
    re.compile(r'^\s*\*'),           # * 항목
]


EXCEPTION_HEADS = [
    re.compile(r'^\s*\(연령\)'),  
    # 필요하면 더 추가 가능
]

def is_paragraph_head(line: str) -> bool:
    for ex in EXCEPTION_HEADS:
        if ex.match(line):
            return False
    for pat in HEAD_PATTERNS:
        if pat.match(line):
            return True
    return False


# =========================
# 3. PDF 내 텍스트 블록 추출
# =========================
def extract_text_blocks(page, min_chars: int = 30) -> str:
    """
    PDF 페이지에서 selectable text가 존재할 때 사용하는 경로.
    '|' 문자는 테이블 윤곽/셀 구분에 쓰이는 경우가 많으므로
    그냥 삭제해버린다.
    """
    blocks = page.get_text("blocks")
    if not blocks:
        return ""

    # (x0,y0,x1,y1, text, block_no, block_type, block_flags, ...)
    blocks = sorted(blocks, key=lambda b: (b[1], b[0]))

    paragraphs = []
    for block in blocks:
        text = block[4]
        if not text:
            continue

        # 공통 노이즈 제거 (|, 영어알파벳 삭제)
        text = clean_noise(text.strip())

        if text.strip():
            paragraphs.append(text.strip())

    if not paragraphs:
        return ""

    full_text = "\n\n".join(paragraphs).strip()

    # 너무 짧으면(노이즈 수준) 무시하고 OCR로 넘기기
    if len(full_text) < min_chars:
        return ""

    return full_text


# =========================
# 4. OCR 결과 문단 정규화
# =========================
def normalize_paragraphs(raw_text: str) -> str:
    """
    OCR 결과처럼 줄 단위로 끊긴 텍스트를 문단 단위로 재구성하는 함수.

    - '|' 문자는 테이블 구분용이므로 그냥 삭제
    - 빈 줄(공백만 있는 줄)은 문단 구분으로 사용
    - 그 외 줄들은 기본적으로 이전 줄과 이어붙임
    """
    # '|' + 영어 알파벳 등 노이즈 제거
    cleaned = clean_noise(raw_text)
    lines = [line.rstrip() for line in cleaned.splitlines()]

    paragraphs = []
    buffer = ""

    for line in lines:
        stripped = line.strip()

        # 1) 완전 빈 줄이면 → 문단 끝
        if not stripped:
            if buffer:
                paragraphs.append(buffer.strip())
                buffer = ""
            continue

        # 2) 내용 있는 줄
        if not buffer:
            # 새 문단 시작
            buffer = stripped
        else:
            # 이전 버퍼가 문장 부호로 끝나는지 확인
            # (.,?!… 등 + 한글 문장 끝에 자주 오는 것들)
            if buffer[-1] in ".?!…）)":
                # 문장이 확실히 끝났으면 줄바꿈 후 이어붙이기
                buffer = buffer + "\n" + stripped
            else:
                # 같은 문단 안에서 줄만 바뀐 거라고 보고 공백으로 이어붙임
                buffer = buffer + " " + stripped

    # 마지막 문단 처리
    if buffer:
        paragraphs.append(buffer.strip())

    # 문단 사이를 빈 줄로 구분
    return "\n\n".join(paragraphs)


# =========================
# 5. OCR 경로
# =========================
def ocr_page(page, lang: str = "kor") -> str:
    """
    PDF 페이지를 이미지로 렌더링한 뒤 Tesseract로 OCR 수행.
    - 400dpi
    - Grayscale + autocontrast
    - kor only / psm 4 / oem 1 / preserve_interword_spaces=1
    """
    pix = page.get_pixmap(dpi=400)
    img_data = pix.tobytes("png")
    img = Image.open(io.BytesIO(img_data))

    # 그레이스케일 + 자동 대비
    gray = img.convert("L")
    gray = ImageOps.autocontrast(gray)

    # 언어는 오직 kor만 사용
    lang_for_tess = "kor"

    # Tesseract 설정: 한 컬럼 위주 공문서에 맞게 psm 4
    config = "--psm 4 --oem 1 -c preserve_interword_spaces=1"

    raw_text = pytesseract.image_to_string(gray, lang=lang_for_tess, config=config)

    # 줄 단위 결과를 문단 단위로 재구성
    normalized = normalize_paragraphs(raw_text)
    return normalized


# =========================
# 6. 머리표 기준 문단 분해
# =========================
def split_paragraphs_by_heads(full_text: str) -> str:
    """
    전체 텍스트를 줄 단위로 보면서
    1. / (1) / (가) / 가. / * 로 시작하는 줄을
    '새 문단의 머리표'로 보고 문단을 나누는 함수.
    """
    # 0) 줄 중간에 끼어 있는 머리표도 강제로 줄바꿈
    full_text = force_heads_to_newline(full_text)

    lines = full_text.splitlines()

    paragraphs = []
    current = []

    for line in lines:
        # 혹시 남아 있을지 모르는 '|' + 영어 노이즈 제거
        line = clean_noise(line)
        stripped = line.rstrip()

        # 빈 줄이면 문단 종료
        if not stripped:
            if current:
                paragraphs.append("\n".join(current).strip())
                current = []
            continue

        # 머리표인지 확인 (앞쪽 공백 제거 후 판별)
        head_candidate = stripped.lstrip()

        if is_paragraph_head(head_candidate):
            # 이전 문단 마감
            if current:
                paragraphs.append("\n".join(current).strip())
                current = []
            # 새 문단 시작: 머리표가 있는 줄을 첫 줄로
            current.append(head_candidate)
        else:
            # 기존 문단에 이어붙임
            current.append(stripped)

    # 마지막 문단 처리
    if current:
        paragraphs.append("\n".join(current).strip())

    # 문단 사이에 빈 줄 하나씩
    return "\n\n".join(paragraphs)


# =========================
# 7. PDF 전체 처리
# =========================
def extract_pdf_to_text(pdf_path: str, lang: str = "kor") -> str:
    doc = fitz.open(pdf_path)
    all_pages_text = []

    for page_index in range(len(doc)):
        page = doc[page_index]
        page_num = page_index + 1

        # 1차: PDF 텍스트 추출
        text = extract_text_blocks(page)

        used_ocr = False
        if not text:
            # 텍스트가 거의 없으면 OCR 사용
            text = ocr_page(page, lang=lang)
            used_ocr = True

        header = f"-------- {page_num}페이지 --------"
        page_content = header + "\n\n" + text + "\n"
        all_pages_text.append(page_content)

        mode = "OCR" if used_ocr else "텍스트"
        print(f"[INFO] {page_num}/{len(doc)}페이지 처리 ({mode})")

    # 1) 페이지별 텍스트를 하나로 합치기
    full_text = "\n\n".join(all_pages_text).strip()

    # 2) 문단 머리표(1. / (1) / (가) / 가. / *) 기준으로 다시 문단 분해
    full_text = split_paragraphs_by_heads(full_text)

    return full_text


# =========================
# 8. CLI 진입점
# =========================
def main():
    if len(sys.argv) < 2:
        print("사용법: python pdf_text_ocr_cli.py <PDF파일경로>")
        sys.exit(1)

    pdf_path = sys.argv[1]

    if not os.path.exists(pdf_path):
        print(f"[ERROR] 파일을 찾을 수 없습니다: {pdf_path}")
        sys.exit(1)

    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    output_path = os.path.join(desktop, base_name + ".txt")

    print(f"[INFO] PDF 처리 시작: {pdf_path}")
    text = extract_pdf_to_text(pdf_path, lang="kor")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(text)

    print(f"[완료] 결과 저장: {output_path}")


if __name__ == "__main__":
    main()
