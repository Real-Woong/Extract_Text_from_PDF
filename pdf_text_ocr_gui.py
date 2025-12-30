import os
import sys
import io
import platform
import fitz
import pytesseract
from PIL import Image, ImageOps
import tkinter as tk
from tkinter import filedialog, messagebox

# ======= Tesseract 경로 자동 설정 =========
def init_tesseract_path():
    system = platform.system()

    # PyInstaller로 빌드된 exe 안에서 실행될 때는
    # sys._MEIPASS 를 통해 base_dir가 잡히고,
    # 그냥 .py로 실행할 때는 현재 파일 위치 기준으로 잡히게 함.
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(sys.argv[0])))

    # 1) 우선 exe 옆의 tesseract 폴더 우선 사용
    bundled_tesseract = os.path.join(base_dir, "tesseract", "tesseract.exe")
    if os.path.exists(bundled_tesseract):
        pytesseract.pytesseract.tesseract_cmd = bundled_tesseract
        return

    # 2) 못 찾으면 OS별 기본 설치 경로 탐색 (개발용/테스트용)
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
        candidates = ["tesseract"]  # 리눅스 등 기본 PATH

    for path in candidates:
        if os.path.exists(path):
            pytesseract.pytesseract.tesseract_cmd = path
            return
    # 여기까지 못 찾으면, 환경변수 PATH에 있는 tesseract에 맡김

init_tesseract_path()
# ====================================


def extract_text_blocks(page, min_chars=30):
    """
    텍스트 기반 PDF 추출 함수.
    - 페이지에서 text blocks를 가져와 위치 기준으로 정렬
    - 블록들을 문단처럼 이어 붙임
    - 너무 짧으면 (페이지 번호 수준) '' 반환해서 OCR로 넘김
    """
    blocks = page.get_text("blocks")
    if not blocks:
        return ""
    
    # (x0, y0, x1, y1, text, ...) -> y(위), x(왼쪽) 순 정렬
    blocks = sorted(blocks, key=lambda b: (b[1], b[0]))

    paragraphs = []
    for block in blocks:
        text = block[4].strip()
        if text:
            paragraphs.append(text)
    
    if not paragraphs:
        return ""
    
    full_text = "\n\n".join(paragraphs).strip()

    # 너무 짧으면 "텍스트 PDF"라고 보기 어렵다고 판단
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


def ocr_page(page, lang="kor"):
    """
    스캔(이미지) 기반 PDF용 OCR 함수.
    - 페이지를 400dpi 이미지로 렌더링
    - 그레이스케일 + autocontrast + 이진화 후
    - Tesseract로 한글 인식
    - normalize_paragraphs로 문단 재구성
    """
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


def extract_pdf_to_text(pdf_path, lang="kor", callback=None):
    """
    텍스트 PDF + 스캔 PDF 모두 처리.
    각 페이지마다:
      1) 텍스트 추출 시도
      2) 실패 시 OCR
    """
    doc = fitz.open(pdf_path)
    all_pages_text = []
    
    for page_index in range(len(doc)):
        page = doc[page_index]
        page_num = page_index + 1

        # 1) 텍스트 기반 먼저 시도
        text = extract_text_blocks(page)

        used_ocr = False
        if not text:
            # 2) 스캔본으로 보고 OCR
            text = ocr_page(page, lang=lang)
            used_ocr = True

        header = f"-------- {page_num}페이지 --------"
        page_content = header + "\n\n" + text + "\n"
        all_pages_text.append(page_content)

        if callback:
            callback(page_num, len(doc), used_ocr)
        
    full_text = "\n\n".join(all_pages_text).strip()
    return full_text


def get_desktop_path():
    return os.path.join(os.path.expanduser("~"), "Desktop")


class PDFTextOCRApp:
    def __init__(self, master):
        self.master = master
        master.title("PDF 텍스트 풀기 (텍스트 + 스캔 OCR)")
        master.geometry("500x220")

        self.pdf_path = None

        # 설명 라벨
        self.label = tk.Label(master, text="PDF 파일을 선택한 뒤 '변환 시작'을 눌러주세요.")
        self.label.pack(pady=10)

        # 선택된 파일 경로 표시
        self.path_label = tk.Label(master, text="선택된 파일: 없음", wraplength=480, justify="left")
        self.path_label.pack(pady=5)

        # PDF 선택 버튼
        self.select_button = tk.Button(master, text="PDF 선택하기", command=self.select_pdf)
        self.select_button.pack(pady=5)

        # 변환 버튼
        self.convert_button = tk.Button(master, text="변환 시작", command=self.convert_pdf)
        self.convert_button.pack(pady=10)

        # 상태 표시 라벨
        self.status_label = tk.Label(master, text="대기 중", fg="gray")
        self.status_label.pack(pady=5)

        # 안내 라벨
        self.info_label = tk.Label(master, text="결과는 바탕화면에 [원본파일명].txt 로 저장됩니다.")
        self.info_label.pack(pady=5)

    def select_pdf(self):
        filetypes = [("PDF 파일", "*.pdf"), ("모든 파일", "*.*")]
        path = filedialog.askopenfilename(title="PDF 선택", filetypes=filetypes)

        if path:
            self.pdf_path = path
            self.path_label.config(text=f"선택된 파일: {path}")
            self.status_label.config(text="대기 중", fg="gray")

        else:
            self.pdf_path = None
            self.path_label.config(text="선택된 파일: 없음")

    def progress_callback(self, current, total, used_ocr):
        mode = "OCR" if used_ocr else "텍스트"
        self.status_label.config(
            text=f"{current}/{total}페이지 처리중... ({mode})",
            fg="blue"
        )
        self.master.update_idletasks()

    def convert_pdf(self):
        if not self.pdf_path:
            messagebox.showwarning("알림", "먼저 PDF 파일을 선택해주세요.")
            return
        try:
            self.master.config(cursor="wait")
            self.status_label.config(text="변환 중...", fg="blue")
            self.master.update()

            text = extract_pdf_to_text(
                self.pdf_path,
                lang="kor",  # 한글 위주 문서라고 보고 기본값 kor 사용
                callback=self.progress_callback
            )
            desktop = get_desktop_path()
            base_name = os.path.splitext(os.path.basename(self.pdf_path))[0]
            output_path = os.path.join(desktop, base_name + ".txt")

            with open(output_path, "w", encoding="utf-8") as f:
                f.write(text)

            self.status_label.config(text="완료", fg="green")
            messagebox.showinfo(
                "완료",
                f"변환이 완료되었습니다.\n\n저장 위치: \n{output_path}"
            )

        except Exception as e: 
            self.status_label.config(text="에러 발생", fg="red")
            messagebox.showerror("에러", f"변환 중 오류가 발생했습니다.\n\n{e}")
        finally:
            self.master.config(cursor="")
            self.master.update_idletasks()


def main():
    root = tk.Tk()
    app = PDFTextOCRApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
