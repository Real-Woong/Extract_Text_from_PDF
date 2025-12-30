import os
import sys
import io
import platform
import fitz
import pytesseract
from PIL import Image, ImageOps

# ======= Tesseract 경로 자동 설정 =========
def init_tesseract_path():
    system = platform.system()

    # PyInstaller나 exe 기준 base_dir
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(sys.argv[0])))

    # 1) exe 옆 tesseract 폴더 우선
    bundled_tesseract = os.path.join(base_dir, "tesseract", "tesseract.exe")
    if os.path.exists(bundled_tesseract):
        pytesseract.pytesseract.tesseract_cmd = bundled_tesseract
        return

    # 2) OS별 후보 경로
    if system == "Windows":
        candidates = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        ]
    elif system == "Darwin":
        candidates = [
            "/opt/homebrew/bin/tesseract",
            "/usr/local/bin/tesseract",
        ]
    else:
        candidates = ["tesseract"]

    for path in candidates:
        if os.path.exists(path):
            pytesseract.pytesseract.tesseract_cmd = path
            return
    # 끝까지 못 찾으면 PATH에 맡김

init_tesseract_path()
# ====================================

def extract_text_blocks(page, min_chars=30):
    blocks = page.get_text("blocks")
    if not blocks:
        return ""
    
    blocks = sorted(blocks, key=lambda b: (b[1], b[0]))

    paragraphs = []
    for block in blocks:
        text = block[4].strip()
        if text:
            paragraphs.append(text)
    
    if not paragraphs:
        return ""
    
    full_text = "\n\n".join(paragraphs).strip()

    if len(full_text) < min_chars:
        return ""
    
    return full_text
    
def normalize_paragraphs(raw_text: str) -> str:
    """
    OCR 결과처럼 줄 단위로 끊긴 텍스트를
    문단 단위로 재구성하는 함수.

    - 빈 줄(공백만 있는 줄)은 문단 구분으로 사용
    - 그 외 줄들은 기본적으로 이전 줄과 이어붙임
    """
    lines = [line.rstrip() for line in raw_text.splitlines()]

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
            # (.,?! 등 + 한글 문장 끝에 자주 오는 것들)
            if buffer[-1] in ".?!…）)" :
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


def ocr_page(page, lang="kor"):
    # 1) 해상도 조금 더 높게 렌더링 (300 → 400)
    pix = page.get_pixmap(dpi=400)
    img_data = pix.tobytes("png")
    img = Image.open(io.BytesIO(img_data))

    # 2) 그레이스케일 변환
    gray = img.convert("L")

    # 3) 자동 대비 조정 (어두운/밝은 영역 강조)
    gray = ImageOps.autocontrast(gray)

    # 4) 단순 이진화 (threshold 값은 상황 따라 조정 가능, 160~190 사이 시도)
    bw = gray.point(lambda x: 0 if x < 180 else 255, "1")

    # 5) Tesseract 설정: 한국어 + 일반 문단(`psm 6`), LSTM 엔진(`oem 1`)
    config = "--psm 6 --oem 1"

    raw_text = pytesseract.image_to_string(bw, lang=lang, config=config)

    # 줄 단위 결과를 문단 단위로 재구성
    normalized = normalize_paragraphs(raw_text)

    return normalized



def extract_pdf_to_text(pdf_path, lang="kor"):
    doc = fitz.open(pdf_path)
    all_pages_text = []
    
    for page_index in range(len(doc)):
        page = doc[page_index]
        page_num = page_index + 1

        text = extract_text_blocks(page)

        used_ocr = False
        if not text:
            text = ocr_page(page, lang=lang)
            used_ocr = True

        header = f"-------- {page_num}페이지 --------"
        page_content = header + "\n\n" + text + "\n"
        all_pages_text.append(page_content)

        mode = "OCR" if used_ocr else "텍스트"
        print(f"[INFO] {page_num}/{len(doc)}페이지 처리 ({mode})")

    full_text = "\n\n".join(all_pages_text).strip()
    return full_text


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
