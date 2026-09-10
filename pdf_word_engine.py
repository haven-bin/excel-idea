import os
import io
import re
import zipfile
import uuid
import shutil
import tempfile
import traceback
import subprocess
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
    """PDF -> DOCX using LibreOffice CLI or 1-to-1 exact page layout renderer"""
    with tempfile.TemporaryDirectory() as tmp_dir:
        input_pdf = os.path.join(tmp_dir, "input.pdf")
        output_docx = os.path.join(tmp_dir, "output.docx")

        with open(input_pdf, "wb") as f:
            f.write(pdf_bytes)

        # 1. Try LibreOffice CLI for PDF -> DOCX (Highest quality across all CAD & text pages)
        if _try_libreoffice_pdf_to_docx(input_pdf, output_docx, tmp_dir):
            with open(output_docx, "rb") as f:
                return f.read()

        # 2. Use 1-to-1 Exact Page Layout Renderer (Guarantees 1-to-1 positions for CAD diagrams & ZERO blank pages)
        try:
            return _pdf_to_docx_exact_1to1(pdf_bytes)
        except Exception as e:
            print(f"1-to-1 PDF to DOCX convert error: {e}, falling back to PyMuPDF image extraction...")
            traceback.print_exc()

        # 3. Ultimate Fallback: extract every page as a high-definition image
        return _scanned_pdf_to_docx(pdf_bytes)


def _pdf_to_docx_exact_1to1(pdf_bytes: bytes) -> bytes:
    """
    1-to-1 Exact Page Layout Renderer:
    - Matches PDF page width, height, and orientation (Portrait/Landscape) per section.
    - Zero margins and zero paragraph spacing prevent ANY extra blank pages.
    - 100% 1-to-1 exact positioning for CAD drawings, vector diagrams, title blocks & stamp seals.
    """
    import pymupdf
    import docx
    from docx.shared import Pt
    from docx.enum.section import WD_ORIENT, WD_SECTION

    pdf_doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    doc = docx.Document()
    
    # Remove initial default paragraph created by Document()
    if len(doc.paragraphs) > 0:
        p_elem = doc.paragraphs[0]._element
        p_elem.getparent().remove(p_elem)

    for page_idx in range(len(pdf_doc)):
        page = pdf_doc[page_idx]
        rect = page.rect
        w_pt, h_pt = rect.width, rect.height
        is_landscape = w_pt > h_pt

        # Add section for page >= 1
        if page_idx == 0:
            section = doc.sections[0]
        else:
            section = doc.add_section(WD_SECTION.NEW_PAGE)

        # Set section orientation and dimensions to match PDF page exactly
        if is_landscape:
            section.orientation = WD_ORIENT.LANDSCAPE
            section.page_width = Pt(w_pt)
            section.page_height = Pt(h_pt)
        else:
            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = Pt(w_pt)
            section.page_height = Pt(h_pt)

        # Zero margins for 1-to-1 exact fit without blank pages
        section.top_margin = Pt(0)
        section.bottom_margin = Pt(0)
        section.left_margin = Pt(0)
        section.right_margin = Pt(0)
        section.header_distance = Pt(0)
        section.footer_distance = Pt(0)

        # Render page as high-res 200 DPI PNG
        pix = page.get_pixmap(dpi=200)
        img_bytes = pix.tobytes("png")
        img_stream = io.BytesIO(img_bytes)

        # Create clean paragraph with 0 spacing
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0

        run = p.add_run()
        run.add_picture(img_stream, width=Pt(w_pt), height=Pt(h_pt))

    pdf_doc.close()
    out_buf = io.BytesIO()
    doc.save(out_buf)
    return out_buf.getvalue()


def _try_libreoffice_pdf_to_docx(input_pdf: str, output_docx: str, tmp_dir: str) -> bool:
    candidates = [
        'libreoffice', 'soffice',
        '/usr/bin/libreoffice', '/usr/bin/soffice',
        r'C:\Program Files\LibreOffice\program\soffice.exe',
        r'C:\Program Files (x86)\LibreOffice\program\soffice.exe'
    ]
    soffice_bin = None
    for cand in candidates:
        if shutil.which(cand) or os.path.exists(cand):
            soffice_bin = cand
            break

    if not soffice_bin:
        return False

    try:
        res = subprocess.run(
            [soffice_bin, '--headless', '--convert-to', 'docx', '--outdir', tmp_dir, input_pdf],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=120
        )
        base_name = os.path.splitext(os.path.basename(input_pdf))[0]
        generated_docx = os.path.join(tmp_dir, f"{base_name}.docx")

        if os.path.exists(generated_docx) and os.path.getsize(generated_docx) > 500:
            if generated_docx != output_docx:
                shutil.move(generated_docx, output_docx)
            return True
    except Exception as e:
        print(f"LibreOffice PDF to DOCX execution error: {e}")
    return False


def _scanned_pdf_to_docx(pdf_bytes: bytes) -> bytes:
    import pymupdf
    import docx
    from docx.shared import Inches

    pdf_doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    doc = docx.Document()

    for page_idx in range(len(pdf_doc)):
        page = pdf_doc[page_idx]
        pix = page.get_pixmap(dpi=150)
        img_bytes = pix.tobytes("png")
        img_stream = io.BytesIO(img_bytes)

        if page_idx > 0:
            doc.add_page_break()

        doc.add_picture(img_stream, width=Inches(6.5))

    pdf_doc.close()
    out_buf = io.BytesIO()
    doc.save(out_buf)
    return out_buf.getvalue()


def _docx_to_pdf(docx_bytes: bytes, filename: str) -> bytes:
    """
    DOCX -> PDF converter with multi-tier engine:
    1. Try LibreOffice CLI (`soffice` / `libreoffice`) - High fidelity (Linux/Windows)
    2. Try docx2pdf (MS Word COM) - High fidelity (Windows MS Word)
    3. Enhanced PyMuPDF Layout Renderer - Handles text, styles, tables & embedded images
    """
    with tempfile.TemporaryDirectory() as tmp_dir:
        input_docx = os.path.join(tmp_dir, "input.docx")
        output_pdf = os.path.join(tmp_dir, "output.pdf")

        with open(input_docx, "wb") as f:
            f.write(docx_bytes)

        # 1. Try LibreOffice CLI
        if _try_libreoffice_convert(input_docx, output_pdf, tmp_dir):
            with open(output_pdf, "rb") as f:
                return f.read()

        # 2. Try docx2pdf with COM (Windows with MS Word installed)
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
            print(f"docx2pdf convert warning: {e}, attempting enhanced PyMuPDF layout fallback...")

        # 3. Enhanced Python layout fallback
        try:
            return _docx_to_pdf_fallback_pymupdf(docx_bytes, filename)
        except Exception as e:
            traceback.print_exc()
            raise RuntimeError(f"Word 转 PDF 转换失败: {e}")


def _try_libreoffice_convert(input_docx: str, output_pdf: str, tmp_dir: str) -> bool:
    """Uses LibreOffice / soffice in headless mode if installed on system"""
    candidates = [
        'libreoffice', 'soffice',
        '/usr/bin/libreoffice', '/usr/bin/soffice',
        r'C:\Program Files\LibreOffice\program\soffice.exe',
        r'C:\Program Files (x86)\LibreOffice\program\soffice.exe'
    ]
    soffice_bin = None
    for cand in candidates:
        if shutil.which(cand) or os.path.exists(cand):
            soffice_bin = cand
            break

    if not soffice_bin:
        return False

    try:
        res = subprocess.run(
            [soffice_bin, '--headless', '--convert-to', 'pdf', '--outdir', tmp_dir, input_docx],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=60
        )
        base_name = os.path.splitext(os.path.basename(input_docx))[0]
        generated_pdf = os.path.join(tmp_dir, f"{base_name}.pdf")

        if os.path.exists(generated_pdf) and os.path.getsize(generated_pdf) > 0:
            if generated_pdf != output_pdf:
                shutil.move(generated_pdf, output_pdf)
            return True
    except Exception as e:
        print(f"LibreOffice execution error: {e}")
    return False


def _docx_to_pdf_fallback_pymupdf(docx_bytes: bytes, filename: str) -> bytes:
    """
    Enhanced Python fallback using python-docx & PyMuPDF with:
    - Headings & paragraph styles
    - Text alignment (left, center, right)
    - Full table drawing (borders, background fill, multi-column wrap)
    - Embedded image extraction & positioning
    - CJK Chinese font rendering
    """
    import docx
    import pymupdf

    doc = docx.Document(io.BytesIO(docx_bytes))
    pdf_doc = pymupdf.open()
    page = pdf_doc.new_page()

    margin = 40
    page_width = page.rect.width
    page_height = page.rect.height
    content_width = page_width - 2 * margin
    y = margin

    def check_page_break(needed_height=20):
        nonlocal page, y
        if y + needed_height > page_height - margin:
            page = pdf_doc.new_page()
            y = margin

    # Process elements in order of document body
    for element in doc.element.body:
        tag_name = element.tag.split('}')[-1]

        # 1. Paragraph element
        if tag_name == 'p':
            p = docx.text.paragraph.Paragraph(element, doc)
            text = p.text.strip()
            
            # Check for images in paragraph
            images = []
            for r in p.runs:
                for drawing in r._element.xpath('.//a:blip'):
                    embed_id = drawing.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                    if embed_id and embed_id in p.part.rels:
                        img_part = p.part.rels[embed_id].target_part
                        images.append(img_part.blob)

            if images:
                for img_blob in images:
                    try:
                        img_stream = io.BytesIO(img_blob)
                        img_pix = pymupdf.Pixmap(img_stream)
                        
                        # Scale image to fit page width
                        img_w = min(img_pix.width, content_width)
                        img_h = (img_w / img_pix.width) * img_pix.height
                        
                        check_page_break(img_h + 10)
                        rect = pymupdf.Rect(margin, y, margin + img_w, y + img_h)
                        page.insert_image(rect, stream=img_blob)
                        y += img_h + 10
                    except Exception as e:
                        print(f"Error embedding image: {e}")

            if not text:
                y += 8
                continue

            # Determine font size and style from paragraph style / headings
            style_name = p.style.name.lower() if p.style else ''
            
            if 'title' in style_name:
                font_size = 20
                line_height = 28
            elif 'heading 1' in style_name or 'heading1' in style_name:
                font_size = 16
                line_height = 24
            elif 'heading 2' in style_name or 'heading2' in style_name:
                font_size = 14
                line_height = 22
            else:
                font_size = 11
                line_height = 18

            # Determine alignment
            alignment = p.alignment
            align_str = 'left'
            if alignment == docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER:
                align_str = 'center'
            elif alignment == docx.enum.text.WD_ALIGN_PARAGRAPH.RIGHT:
                align_str = 'right'

            chars_per_line = int(content_width / (font_size * 0.55))
            chars_per_line = max(15, chars_per_line)
            
            lines = [text[i:i+chars_per_line] for i in range(0, len(text), chars_per_line)]

            for line in lines:
                check_page_break(line_height)
                
                # Calculate X position based on alignment
                line_w = len(line) * (font_size * 0.55)
                if align_str == 'center':
                    start_x = margin + max(0, (content_width - line_w) / 2)
                elif align_str == 'right':
                    start_x = margin + max(0, content_width - line_w)
                else:
                    start_x = margin

                page.insert_text(
                    pymupdf.Point(start_x, y + font_size),
                    line,
                    fontname="china-s",
                    fontsize=font_size
                )
                y += line_height

            y += 4

        # 2. Table element
        elif tag_name == 'tbl':
            table = docx.table.Table(element, doc)
            num_cols = len(table.columns)
            if num_cols == 0:
                continue

            col_width = content_width / num_cols
            cell_padding = 6
            row_min_height = 24

            for row in table.rows:
                # Calculate row height based on cell text length
                max_lines = 1
                for cell in row.cells:
                    cell_txt = cell.text.strip()
                    chars_per_cell = max(5, int((col_width - 2 * cell_padding) / 6.5))
                    cell_lines = (len(cell_txt) // chars_per_cell) + 1
                    max_lines = max(max_lines, cell_lines)

                cell_h = max(row_min_height, max_lines * 16 + cell_padding * 2)
                check_page_break(cell_h + 4)

                # Draw row cells
                for col_idx, cell in enumerate(row.cells):
                    cell_x0 = margin + col_idx * col_width
                    cell_y0 = y
                    cell_x1 = cell_x0 + col_width
                    cell_y1 = y + cell_h

                    cell_rect = pymupdf.Rect(cell_x0, cell_y0, cell_x1, cell_y1)
                    
                    # Draw cell border line
                    page.draw_rect(cell_rect, color=(0.7, 0.7, 0.7), fill=(0.97, 0.98, 0.99) if row == table.rows[0] else None, width=0.5)

                    # Insert cell text
                    cell_txt = cell.text.strip()
                    if cell_txt:
                        chars_per_cell = max(5, int((col_width - 2 * cell_padding) / 6.5))
                        c_lines = [cell_txt[i:i+chars_per_cell] for i in range(0, len(cell_txt), chars_per_cell)]
                        
                        txt_y = cell_y0 + cell_padding + 10
                        for c_line in c_lines:
                            if txt_y < cell_y1 - 2:
                                page.insert_text(
                                    pymupdf.Point(cell_x0 + cell_padding, txt_y),
                                    c_line,
                                    fontname="china-s",
                                    fontsize=9.5
                                )
                                txt_y += 14

                y += cell_h

            y += 10

    pdf_bytes_out = pdf_doc.tobytes()
    pdf_doc.close()
    return pdf_bytes_out
