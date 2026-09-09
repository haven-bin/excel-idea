import os
import io
import zipfile
import uuid
import tempfile
from typing import List, Tuple
from PIL import Image

TEMP_DOWNLOADS = {}

def process_image_compression(files_data: List[dict], mode: str) -> dict:
    """
    mode: 'high_compress' (高压缩: 几十~几百KB) or 'high_quality' (高质量压缩: 3-4MB -> ~1MB)
    files_data: [{'filename': str, 'content': bytes}]
    """
    if not files_data:
        return {'error': '未提供需要压缩的图片文件'}

    total_orig_size = sum(len(f['content']) for f in files_data)
    compressed_files: List[Tuple[str, bytes]] = []
    total_compressed_size = 0

    for file_info in files_data:
        orig_name = file_info['filename']
        content = file_info['content']
        base_name, ext = os.path.splitext(orig_name)
        ext_lower = ext.lower()

        # Perform Image Compression
        out_bytes, out_ext = _compress_single_image(content, ext_lower, mode)
        target_name = f"{base_name}_compressed{out_ext}"
        
        compressed_files.append((target_name, out_bytes))
        total_compressed_size += len(out_bytes)

    if not compressed_files:
        return {'error': '图片压缩处理失败'}

    saved_bytes = max(0, total_orig_size - total_compressed_size)
    saved_percent = (saved_bytes / total_orig_size * 100) if total_orig_size > 0 else 0

    # Single File return
    if len(compressed_files) == 1:
        out_filename, out_data = compressed_files[0]
        download_id = str(uuid.uuid4())
        TEMP_DOWNLOADS[download_id] = {
            'filename': out_filename,
            'data': out_data
        }
        return {
            'success': True,
            'count': 1,
            'download_id': download_id,
            'filename': out_filename,
            'orig_size_str': _format_file_size(total_orig_size),
            'compressed_size_str': _format_file_size(total_compressed_size),
            'saved_percent_str': f"{saved_percent:.1f}%"
        }

    # Multiple Files ZIP packaging
    zip_buffer = io.BytesIO()
    zip_filename = f"压缩图片包_共{len(compressed_files)}张.zip"
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for fname, fbytes in compressed_files:
            zip_file.writestr(fname, fbytes)

    zip_bytes = zip_buffer.getvalue()
    download_id = str(uuid.uuid4())
    TEMP_DOWNLOADS[download_id] = {
        'filename': zip_filename,
        'data': zip_bytes
    }

    return {
        'success': True,
        'count': len(compressed_files),
        'download_id': download_id,
        'filename': zip_filename,
        'orig_size_str': _format_file_size(total_orig_size),
        'compressed_size_str': _format_file_size(total_compressed_size),
        'saved_percent_str': f"{saved_percent:.1f}%"
    }


def _compress_single_image(img_bytes: bytes, ext: str, mode: str) -> Tuple[bytes, str]:
    """Compress single image buffer based on mode"""
    img = Image.open(io.BytesIO(img_bytes))

    # Determine max dimensions & quality based on mode
    if mode == 'high_compress':
        max_dim = 1600
        quality = 40
    else:  # 'high_quality'
        max_dim = 2560
        quality = 78

    # Resize if exceeds max_dim
    w, h = img.size
    if w > max_dim or h > max_dim:
        if w >= h:
            new_w = max_dim
            new_h = int(h * (max_dim / w))
        else:
            new_h = max_dim
            new_w = int(w * (max_dim / h))
        img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)

    output = io.BytesIO()

    # Format handling
    if ext in ['.png', '.webp']:
        # If PNG has transparency, convert to RGBA or preserve
        if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
            if mode == 'high_compress':
                # Convert RGBA to RGB with white background for high compression JPEG
                background = Image.new('RGB', img.size, (255, 255, 255))
                background.paste(img, mask=img.split()[3] if img.mode == 'RGBA' else None)
                background.save(output, format='JPEG', quality=quality, optimize=True)
                return output.getvalue(), '.jpg'
            else:
                img.save(output, format='PNG', optimize=True)
                return output.getvalue(), '.png'
        else:
            img = img.convert('RGB')
            img.save(output, format='JPEG', quality=quality, optimize=True)
            return output.getvalue(), '.jpg'
    else:
        # Standard JPG / JPEG / BMP / WEBP
        if img.mode != 'RGB':
            img = img.convert('RGB')
        img.save(output, format='JPEG', quality=quality, optimize=True)
        return output.getvalue(), '.jpg'


def _format_file_size(size_bytes: int) -> str:
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.2f} MB"
