## 📄 Local PDF Text Extraction Tool (OCR-based)

This project is a **local-only PDF text extraction tool** designed to reduce inefficient manual retyping of documents in public-sector and administrative workflows.

The tool automatically handles both:
- **Text-based PDFs** (direct text extraction)
- **Scanned image PDFs** (OCR-based extraction)

All processing is performed **entirely on the local machine** — no server, no cloud API, and no external data transmission — making it suitable for environments with strict security and privacy constraints.

### Key Notes
- Output format: **TXT only** (no DOCX / Word export)
- Designed and tested on **macOS**
- This repository contains the **Python source code only**
- **No PyInstaller / executable build is included** (macOS development environment)

The primary goal of this project is to demonstrate how existing OCR technologies can be combined into a practical, secure, and lightweight workflow automation tool for real-world administrative use.

### Technical Scope
- Language: Python
- PDF Processing: PyMuPDF
- OCR Engine: Tesseract OCR (local)
- Image Preprocessing: Pillow
- Platform: macOS (development environment)
- Output: Plain text (.txt)

This project intentionally avoids cloud-based OCR services and focuses on local execution to align with security-sensitive environments.
