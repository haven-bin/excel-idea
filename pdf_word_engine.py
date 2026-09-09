import os
import io
import zipfile
import uuid
import tempfile
import traceback
from typing import List, Tuple

# Try imports
try:
    from pdf2docx import Converter as PdfConverter
except ImportError:
    PdfConverter = None

try:
    from docx import Document
except ImportError:
    Document = None

TEMP_DOWNLOADS = {}

def process_pdf_word_conversion(files_data: List[dict], mode: str) -> dict:
    """
    mode: 'pdf_to_word' or 'word_to_pdf'
    files_data: [{'filename': str, 'content': bytes}]
    """
    if not files_data:
        return {'error': '未提供需要转换的文件'}

    converted_files: List[Tuple[str, bytes]] = []

    for file_info in files_data:
        orig_name = file_info['filename']
        content = file_info['content']
        base_name, ext = os.path.splitext(orig_name)

        if mode == 'pdf_to_word':
            target_name = f"{base_name}.docx"
            out_bytes = _pdf_to_docx(content, orig_name)
            converted_files.append((target_name, out_bytes))

        elif mode == 'word_to_pdf':
            target_name = f"{base_name}.pdf"
            out_bytes = _docx_to_pdf(content, orig_name)
            converted_files.append((target_name, out_bytes))

        else:
            return {'error': f'不支持的转换模式: {mode}'}

    if not converted_files:
        return {'error': '文件转换处理失败'}

    # If single file converted, export directly
    if len(converted_files) == 1:
        out_filename, out_data = converted_files[0]
        download_id = str(uuid.uuid4())
        TEMP_DOWNLOADS[download_id] = {
            'filename': out_filename,
            'data': out_data
        }
        return {
            'success': True,
            'count': 1,
            'download_id': download_id,
            'filename': out_filename
        }

    # If multiple files, package into a zip archive
    zip_buffer = io.BytesIO()
    zip_filename = f"转换结果_共{len(converted_files)}个文件.zip"
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for fname, fbytes in converted_files:
            zip_file.writestr(fname, fbytes)

    zip_bytes = zip_buffer.getvalue()
    download_id = str(uuid.uuid4())
    TEMP_DOWNLOADS[download_id] = {
        'filename': zip_filename,
        'data': zip_bytes
    }

    return {
        'success': True,
        'count': len(converted_files),
        'download_id': download_id,
        'filename': zip_filename
    }


def _pdf_to_docx(pdf_bytes: bytes, filename: str) -> bytes:
    """PDF -> DOCX using pdf2docx"""
    if PdfConverter is None:
        raise RuntimeError("未安装 pdf2docx 模块")

    with tempfile.TemporaryDirectory() as tmp_dir:
        input_pdf = os.path.join(tmp_dir, "input.pdf")
        output_docx = os.path.join(tmp_dir, "output.docx")

        with open(input_pdf, "wb") as f:
            f.write(pdf_bytes)

        cv = PdfConverter(input_pdf)
        cv.convert(output_docx, start=0, end=None)
        cv.close()

        with open(output_docx, "rb") as f:
            return f.read()


def _docx_to_pdf(docx_bytes: bytes, filename: str) -> bytes:
    """DOCX -> PDF using docx2pdf (with COM CoInitialize) or PyMuPDF/python-docx fallback"""
    with tempfile.TemporaryDirectory() as tmp_dir:
        input_docx = os.path.join(tmp_dir, "input.docx")
        output_pdf = os.path.join(tmp_dir, "output.pdf")

        with open(input_docx, "wb") as f:
            f.write(docx_bytes)

        # 1. Try docx2pdf with COM CoInitialize (Windows MS Word / Office)
        try:
            import pythoncom
            pythoncom.CoInitialize()
            from docx2pdf import convert as docx2pdf_convert
            docx2pdf_convert(input_docx, output_pdf)
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

            if os.path.exists(output_pdf) and os.path.getsize(output_pdf) > 0:
                with open(output_pdf, "rb") as f:
                    return f.read()
        except Exception as e:
            print(f"docx2pdf convert warning: {e}, attempting python PyMuPDF fallback...")

        # 2. Pure Python fallback: parse docx with python-docx and render PDF with PyMuPDF CJK font
        try:
            return _docx_to_pdf_fallback_pymupdf(docx_bytes, filename)
        except Exception as e:
            traceback.print_exc()
            raise RuntimeError(f"Word 转 PDF 转换失败: {e}")


def _docx_to_pdf_fallback_pymupdf(docx_bytes: bytes, filename: str) -> bytes:
    """Fallback text & table to PDF generator using python-docx & PyMuPDF with CJK font support"""
    import docx
    import pymupdf

    doc = docx.Document(io.BytesIO(docx_bytes))
    full_text = []

    # Extract headings & paragraphs
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt:
            full_text.append((txt, 'p'))
        else:
            full_text.append(('', 'blank'))

    # Extract tables
    for table in doc.tables:
        for row in table.rows:
            row_str = " | ".join(cell.text.strip() for cell in row.cells)
            if row_str.strip():
                full_text.append((row_str, 'table'))

    pdf_doc = pymupdf.open()
    page = pdf_doc.new_page()
    
    y = 50
    margin = 40
    line_height = 20

    for item, item_type in full_text:
        if item_type == 'blank':
            y += 10
            continue

        if y > page.rect.height - margin:
            page = pdf_doc.new_page()
            y = margin

        font_size = 13 if item_type == 'p' and len(item) < 30 else 10.5
        line_chunks = [item[i:i+60] for i in range(0, len(item), 60)]
        for chunk in line_chunks:
            page.insert_text(
                pymupdf.Point(margin, y),
                chunk,
                fontname="china-s",
                fontsize=font_size
            )
            y += line_height
            if y > page.rect.height - margin:
                page = pdf_doc.new_page()
                y = margin

    pdf_bytes_out = pdf_doc.tobytes()
    pdf_doc.close()
    return pdf_bytes_out
