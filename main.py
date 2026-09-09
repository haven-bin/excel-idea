import os
import json
from typing import List, Optional
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, Response, FileResponse
from fastapi.staticfiles import StaticFiles
import uvicorn

from excel_engine import analyze_excel_files, process_merge_excel_files, process_clean_excel_file, generate_blank_template, TEMP_DOWNLOADS
from pdf_word_engine import process_pdf_word_conversion, TEMP_DOWNLOADS as PDF_TEMP_DOWNLOADS
from image_compressor_engine import process_image_compression, TEMP_DOWNLOADS as IMAGE_TEMP_DOWNLOADS

import re

app = FastAPI(title="Smart Excel Folder Merger", version="2.6")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
os.makedirs(STATIC_DIR, exist_ok=True)

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

@app.get("/", response_class=HTMLResponse)
async def serve_index():
    index_path = os.path.join(STATIC_DIR, "index.html")
    if os.path.exists(index_path):
        with open(index_path, "r", encoding="utf-8") as f:
            content = f.read()

        def resolve_includes(text):
            def replacer(match):
                rel_path = match.group(1).strip()
                target_path = os.path.join(STATIC_DIR, rel_path)
                if os.path.exists(target_path):
                    with open(target_path, "r", encoding="utf-8") as inc_f:
                        return resolve_includes(inc_f.read())
                return f"<!-- Missing component: {rel_path} -->"
            return re.sub(r'<!--\s*INCLUDE:\s*(.*?)\s*-->', replacer, text)

        final_html = resolve_includes(content)
        return HTMLResponse(content=final_html)
    return HTMLResponse(content="<h1>Smart Excel Merger</h1><p>Index file not found.</p>")

@app.post("/api/analyze")
async def analyze_headers_api(
    files: List[UploadFile] = File(...),
    custom_rules_json: Optional[str] = Form(None),
    similarity_threshold: float = Form(0.75)
):
    try:
        custom_rules = json.loads(custom_rules_json) if custom_rules_json else None
    except Exception:
        custom_rules = None

    files_data = []
    for upload in files:
        content = await upload.read()
        files_data.append({
            'filename': upload.filename,
            'content': content
        })

    if not files_data:
        raise HTTPException(status_code=400, detail="未接收到有效文件")

    result = analyze_excel_files(files_data, custom_rules, similarity_threshold)
    return result

@app.post("/api/export-template")
async def export_blank_template_api(headers_json: str = Form(...)):
    """根据传递的表头列表生成统一空白 Excel 格式模板"""
    try:
        headers = json.loads(headers_json)
    except Exception:
        headers = []

    template_download_id = generate_blank_template(headers)
    return {"template_download_id": template_download_id}

@app.post("/api/merge")
async def merge_files_api(
    files: List[UploadFile] = File(...),
    custom_rules_json: Optional[str] = Form(None),
    selected_columns_json: Optional[str] = Form(None),
    similarity_threshold: float = Form(0.75),
    dedup_columns_json: Optional[str] = Form(None),
    dedup_keep: str = Form("first"),
    clean_rules_json: Optional[str] = Form(None)
):
    try:
        custom_rules = json.loads(custom_rules_json) if custom_rules_json else None
    except Exception:
        custom_rules = None

    try:
        selected_columns = json.loads(selected_columns_json) if selected_columns_json else None
    except Exception:
        selected_columns = None

    try:
        dedup_columns = json.loads(dedup_columns_json) if dedup_columns_json else None
    except Exception:
        dedup_columns = None

    try:
        clean_rules = json.loads(clean_rules_json) if clean_rules_json else None
    except Exception:
        clean_rules = None

    files_data = []
    for upload in files:
        content = await upload.read()
        files_data.append({
            'filename': upload.filename,
            'content': content
        })

    if not files_data:
        raise HTTPException(status_code=400, detail="未接收到有效文件")

    try:
        result = process_merge_excel_files(
            files_data,
            custom_rules,
            selected_columns,
            similarity_threshold,
            dedup_columns,
            dedup_keep,
            clean_rules
        )
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/clean")
async def clean_files_api(
    files: List[UploadFile] = File(...),
    clean_rules_json: Optional[str] = Form(None)
):
    try:
        clean_rules = json.loads(clean_rules_json) if clean_rules_json else None
    except Exception:
        clean_rules = None

    files_data = []
    for upload in files:
        content = await upload.read()
        files_data.append({
            'filename': upload.filename,
            'content': content
        })

    if not files_data:
        raise HTTPException(status_code=400, detail="未接收到有效文件")

    try:
        result = process_clean_excel_file(files_data, clean_rules)
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/convert-pdf-word")
async def convert_pdf_word_api(
    files: List[UploadFile] = File(...),
    mode: str = Form("pdf_to_word")
):
    files_data = []
    for upload in files:
        content = await upload.read()
        files_data.append({
            'filename': upload.filename,
            'content': content
        })

    if not files_data:
        raise HTTPException(status_code=400, detail="未接收到有效文件")

    try:
        result = process_pdf_word_conversion(files_data, mode)
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/compress-images")
async def compress_images_api(
    files: List[UploadFile] = File(...),
    mode: str = Form("high_compress")
):
    files_data = []
    for upload in files:
        content = await upload.read()
        files_data.append({
            'filename': upload.filename,
            'content': content
        })

    if not files_data:
        raise HTTPException(status_code=400, detail="未接收到有效图片文件")

    try:
        result = process_image_compression(files_data, mode)
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/download/{download_id}")
async def download_merged_file(download_id: str):
    file_info = TEMP_DOWNLOADS.get(download_id) or PDF_TEMP_DOWNLOADS.get(download_id) or IMAGE_TEMP_DOWNLOADS.get(download_id)
    if not file_info:
        raise HTTPException(status_code=404, detail="文件不存在或已过期")

    filename = file_info['filename']
    content = file_info.get('content') or file_info.get('data')

    return Response(
        content=content,
        media_type="application/octet-stream",
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{filename}"
        }
    )

if __name__ == "__main__":
    print("🚀 启动 Smart Excel Merger 服务在 http://localhost:8000")
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
