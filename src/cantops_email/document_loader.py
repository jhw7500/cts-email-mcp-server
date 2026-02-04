import os
import sys

def extract_text_from_file(file_path: str) -> str:
    """파일 타입을 감지하고 내용을 추출합니다. (지연 로딩 적용)"""
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext == '.pptx': return _read_pptx(file_path)
        elif ext == '.docx': return _read_docx(file_path)
        elif ext == '.xlsx': return _read_xlsx(file_path)
        elif ext == '.pdf': return _read_pdf(file_path)
        return _read_text(file_path)
    except Exception as e:
        return f"Error reading file {file_path}: {str(e)}"

def _read_pptx(path):
    try: from pptx import Presentation
    except ImportError: return "Error: python-pptx not installed"
    
    prs = Presentation(path)
    text = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"): text.append(shape.text)
    return "\n".join(text)

def _read_docx(path):
    try: from docx import Document
    except ImportError: return "Error: python-docx not installed"
    
    doc = Document(path)
    return "\n".join([para.text for para in doc.paragraphs])

def _read_xlsx(path):
    try: import openpyxl
    except ImportError: return "Error: openpyxl not installed"
    
    wb = openpyxl.load_workbook(path, data_only=True)
    text = []
    for sheet in wb.sheetnames:
        text.append(f"--- Sheet: {sheet} ---")
        ws = wb[sheet]
        for row in ws.iter_rows(values_only=True):
            row_text = [str(cell) for cell in row if cell is not None]
            if row_text: text.append("\t".join(row_text))
    return "\n".join(text)

def _read_pdf(path):
    try: from pypdf import PdfReader
    except ImportError: return "Error: pypdf not installed"
    
    reader = PdfReader(path)
    return "\n".join([page.extract_text() or "" for page in reader.pages])

def _read_text(path):
    with open(path, 'r', encoding='utf-8', errors='ignore') as f:
        return f.read()