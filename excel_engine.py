import os
import re
import io
import uuid
import glob
from difflib import SequenceMatcher
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter

# 字典全局存储临时生成的下载文件
TEMP_DOWNLOADS = {}

def normalize_header(header):
    """文本规范化：转小写、去除空格与常见标点符号/修饰词"""
    if header is None:
        return ""
    s = str(header).strip()
    s = re.sub(r'[\s\*\（\）\(\)\【\】\_\-\:\：]', '', s)
    return s.lower()

def calculate_similarity(t1, t2):
    """计算两个标题之间的文本相似度 (0.0 - 1.0)"""
    n1 = normalize_header(t1)
    n2 = normalize_header(t2)
    if not n1 or not n2:
        return 0.0
    if n1 == n2:
        return 1.0
    if n1 in n2 or n2 in n1:
        len_ratio = min(len(n1), len(n2)) / max(len(n1), len(n2))
        return 0.85 + 0.15 * len_ratio
    return SequenceMatcher(None, n1, n2).ratio()

def auto_fit_excel_columns(writer, df, sheet_name="合并总表"):
    """
    自动计算并加宽 Excel 所有列宽，
    针对用户指定的 id, 手机号, phone, email, 邮箱, time, 时间, password, 密码 等常见长字段与敏感字段
    赋予超大保障列宽 (28 ~ 35 字符宽) 与 8 字节边距，
    确保 MS Excel 与 WPS Office 打开时文字极其舒展宽敞、绝对不会被裁切或显示 ###。
    """
    worksheet = writer.sheets[sheet_name]
    
    for i, col in enumerate(df.columns):
        col_str = str(col)
        # 计算中文字符实际物理显示宽度 (中文字符按 1.5 倍物理宽度计)
        max_display_len = sum(1.5 if ord(char) > 127 else 1.0 for char in col_str)
        
        if len(df) > 0:
            sample_series = df[col].dropna().astype(str).head(300)
            for val in sample_series:
                val_len = sum(1.5 if ord(c) > 127 else 1.0 for c in str(val))
                if val_len > max_display_len:
                    max_display_len = val_len

        col_lower = col_str.lower()
        min_width = 20
        
        # 1. 手机号 / phone / mobile / 电话
        if any(k in col_lower for k in ['手机', '电话', 'phone', 'tel', 'mobile', '联系方式', '呼叫']):
            min_width = 28
        # 2. 邮箱 / email / mail
        elif any(k in col_lower for k in ['邮箱', 'email', 'mail', '邮局']):
            min_width = 35
        # 3. 时间 / time / 日期 / date / 生日 / birthday
        elif any(k in col_lower for k in ['时间', 'time', '日期', 'date', '生日', 'birthday', '时刻', '创建', '更新', '年份', '月份']):
            min_width = 30
        # 4. 密码 / password / pwd / pass
        elif any(k in col_lower for k in ['密码', 'password', 'pwd', 'pass', '口令', '秘钥', 'token']):
            min_width = 32
        # 5. id / 单号 / 编号 / code / uuid / 文件 / 来源
        elif any(k in col_lower for k in ['id', '单号', '编号', 'code', 'uuid', '文件', '来源', '地址', '说明', '备注', '编码']):
            min_width = 30

        # 增加 8 个字符大保护边距，限制最大列宽为 90
        final_width = min(max(max_display_len + 8, min_width), 90)
        
        col_letter = get_column_letter(i + 1)
        worksheet.column_dimensions[col_letter].width = float(final_width)

def build_cluster_mapping(all_headers, custom_rules=None, similarity_threshold=0.75):
    mapping = {}
    mapped_raws = set()
    
    if custom_rules:
        for rule in custom_rules:
            target = rule.get('target', '').strip()
            aliases = rule.get('aliases', [])
            if not target:
                continue
            norm_aliases = [normalize_header(a) for a in aliases] + [normalize_header(target)]
            
            for raw_h in all_headers:
                norm_raw = normalize_header(raw_h)
                if norm_raw in norm_aliases or any(norm_raw == a for a in norm_aliases):
                    mapping[raw_h] = target
                    mapped_raws.add(raw_h)

    unmapped = [h for h in all_headers if h not in mapped_raws]
    clusters = []
    
    for h in unmapped:
        assigned = False
        for cluster in clusters:
            rep = cluster[0]
            if calculate_similarity(h, rep) >= similarity_threshold:
                cluster.append(h)
                assigned = True
                break
        if not assigned:
            clusters.append([h])

    for cluster in clusters:
        std_col = sorted(cluster, key=lambda x: (len(normalize_header(x)), x))[0]
        for h in cluster:
            mapping[h] = std_col
            
    return mapping

def read_excel_file_content(file_bytes, filename):
    """从字节流中读取 Excel 文件内容为 DataFrame"""
    try:
        if filename.endswith('.xls'):
            df = pd.read_excel(io.BytesIO(file_bytes), engine='xlrd')
        else:
            df = pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl')
        return df
    except Exception as e:
        return pd.read_excel(io.BytesIO(file_bytes))

def inspect_cluster_quality(file_dfs, raw_headers):
    samples_by_file = []
    total_rows = 0
    non_null_rows = 0
    detected_types = set()

    for fname, df in file_dfs:
        total_rows += len(df)
        file_samples = []
        for rh in raw_headers:
            if rh in df.columns:
                series = df[rh].dropna()
                non_null_rows += len(series)
                
                dtype_name = str(df[rh].dtype)
                if 'int' in dtype_name or 'float' in dtype_name:
                    detected_types.add('数值/金额')
                elif 'datetime' in dtype_name:
                    detected_types.add('日期时间')
                else:
                    detected_types.add('文本')
                    
                for val in series:
                    val_str = str(val).strip()
                    if val_str and val_str not in file_samples and len(file_samples) < 2:
                        file_samples.append(val_str)
                        
        if file_samples:
            samples_by_file.append({
                'filename': fname,
                'samples': file_samples
            })

    non_null_rate = round((non_null_rows / total_rows * 100), 1) if total_rows > 0 else 0
    type_str = '/'.join(detected_types) if detected_types else '全空数据'
    is_all_empty = (non_null_rows == 0)

    return {
        'samples_by_file': samples_by_file[:3],
        'non_null_rate': non_null_rate,
        'data_type': type_str,
        'is_all_empty': is_all_empty,
        'non_null_count': non_null_rows,
        'total_rows': total_rows
    }

def mask_value(val, mask_type):
    """单项值数据脱敏"""
    if val is None or pd.isna(val):
        return val
    s = str(val).strip()
    if not s:
        return s
        
    if mask_type == 'phone_3_4':
        digits = re.sub(r'\D', '', s)
        if len(digits) == 11:
            return digits[:3] + '****' + digits[7:]
        elif len(s) >= 7:
            return s[:3] + '****' + s[-4:]
        return s
    elif mask_type == 'email':
        if '@' in s:
            parts = s.split('@', 1)
            uname, domain = parts[0], parts[1]
            if len(uname) <= 2:
                masked_u = uname[0] + '*'
            else:
                masked_u = uname[0] + '***' + uname[-1]
            return masked_u + '@' + domain
        return s
    elif mask_type == 'idcard':
        if len(s) == 18:
            return s[:6] + '******' + s[14:]
        elif len(s) == 15:
            return s[:6] + '******' + s[12:]
        return s
    elif mask_type == 'name':
        if len(s) == 2:
            return s[0] + '*'
        elif len(s) > 2:
            return s[0] + '*' * (len(s) - 2) + s[-1]
        return s
    elif mask_type == 'full_star':
        return '****'
    return s

def standardize_date_val(val, date_format='YYYY-MM-DD'):
    """日期格式统一规范化"""
    if val is None or pd.isna(val):
        return val
    s = str(val).strip()
    if not s or s.lower() in ['none', 'null', 'nan', 'n/a']:
        return s
        
    try:
        dt = pd.to_datetime(s, errors='coerce')
        if pd.isna(dt):
            return s
        if date_format == 'YYYY-MM-DD':
            return dt.strftime('%Y-%m-%d')
        elif date_format == 'YYYY/MM/DD':
            return dt.strftime('%Y/%m/%d')
        elif date_format == 'YYYY年MM月DD日':
            return dt.strftime('%Y年%m月%d日')
        elif date_format == 'YYYY-MM-DD HH:mm:ss':
            return dt.strftime('%Y-%m-%d %H:%M:%S')
        elif date_format == 'YYYYMMDD':
            return dt.strftime('%Y%m%d')
        return dt.strftime('%Y-%m-%d')
    except Exception:
        return s

def clean_text_val(val, trim_space=True, remove_linebreaks=False, case_transform='none', fill_empty=None):
    """文本过滤与转换"""
    if val is None or pd.isna(val) or str(val).strip() == "":
        if fill_empty is not None and str(fill_empty).strip() != "":
            return fill_empty
        return ""
    s = str(val)
    if trim_space:
        s = s.strip()
    if remove_linebreaks:
        s = re.sub(r'[\r\n]+', ' ', s)
    if case_transform == 'upper':
        s = s.upper()
    elif case_transform == 'lower':
        s = s.lower()
    return s

def recommend_cleaning_rules(cluster_target, quality_info):
    """根据列名和数据特征智能推荐清洗与治理规则"""
    recs = []
    col_lower = str(cluster_target).lower()
    
    # 检查手机号
    if any(k in col_lower for k in ['手机', '电话', 'phone', 'mobile', 'tel']):
        recs.append({
            'type': 'mask',
            'suggested_rule': 'phone_3_4',
            'label': '💡 建议手机号脱敏 (前3后4)',
            'desc': '涉及敏感联系方式，推荐脱敏保护隐私'
        })
    # 检查邮箱
    elif any(k in col_lower for k in ['邮箱', 'email', 'mail']):
        recs.append({
            'type': 'mask',
            'suggested_rule': 'email',
            'label': '💡 建议邮箱掩码',
            'desc': '推荐对邮箱用户名进行掩码保护'
        })
    # 检查身份证
    elif any(k in col_lower for k in ['身份证', '证件号', 'idcard']):
        recs.append({
            'type': 'mask',
            'suggested_rule': 'idcard',
            'label': '💡 建议身份证隐藏生日',
            'desc': '隐藏身份证中间 6 位出生日期'
        })
    
    # 检查日期
    if any(k in col_lower for k in ['日期', 'date', '时间', 'time', '生日', 'birthday']):
        recs.append({
            'type': 'date_format',
            'suggested_rule': 'YYYY-MM-DD',
            'label': '💡 建议统一日期格式为 YYYY-MM-DD',
            'desc': '自动将多种混乱格式转换为标准 YYYY-MM-DD'
        })
        
    # 检查是否有首尾空格
    if quality_info:
        samples = []
        for sf in quality_info.get('samples_by_file', []):
            samples.extend(sf.get('samples', []))
        has_whitespace = any(isinstance(s, str) and s != s.strip() for s in samples)
        if has_whitespace:
            recs.append({
                'type': 'trim',
                'suggested_rule': 'trim',
                'label': '💡 建议去除首尾多余空格',
                'desc': '检测到部分样本含有不可见首尾空格'
            })
            
    return recs

def analyze_excel_files(files_data, custom_rules=None, similarity_threshold=0.75):
    all_headers = set()
    file_summaries = []
    file_dfs = []

    for file_info in files_data:
        fname = file_info['filename']
        content = file_info['content']
        try:
            df = read_excel_file_content(content, fname)
            headers = [str(c) for c in df.columns]
            all_headers.update(headers)
            file_dfs.append((fname, df))
            
            file_summaries.append({
                'filename': fname,
                'rows': len(df),
                'columns_count': len(headers),
                'columns': headers
            })
        except Exception as e:
            file_summaries.append({
                'filename': fname,
                'error': f"文件解析失败: {str(e)}",
                'rows': 0,
                'columns_count': 0,
                'columns': []
            })

    all_headers_list = list(all_headers)
    mapping = build_cluster_mapping(all_headers_list, custom_rules, similarity_threshold)
    
    target_to_raws = {}
    for raw_h, target_h in mapping.items():
        target_to_raws.setdefault(target_h, []).append(raw_h)
        
    suggested_clusters = []
    for target_col, raw_list in target_to_raws.items():
        quality_info = inspect_cluster_quality(file_dfs, raw_list)
        recommendations = recommend_cleaning_rules(target_col, quality_info)
        suggested_clusters.append({
            'target': target_col,
            'raw_headers': raw_list,
            'quality_info': quality_info,
            'recommendations': recommendations
        })

    return {
        'total_files': len(files_data),
        'file_summaries': file_summaries,
        'all_raw_headers': all_headers_list,
        'suggested_clusters': suggested_clusters,
        'mapping': mapping
    }

def generate_blank_template(headers):
    """根据统一目标列名列表生成空数据 Excel 模板文件 (含列宽强行自适应)"""
    if not headers:
        headers = ["示例列1", "示例列2"]
    
    clean_headers = [h for h in headers if h != '_来源文件']
    empty_df = pd.DataFrame(columns=clean_headers)
    
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        empty_df.to_excel(writer, index=False, sheet_name="标准填写模板")
        auto_fit_excel_columns(writer, empty_df, sheet_name="标准填写模板")
        
    output_bytes = output_buffer.getvalue()
    
    download_id = str(uuid.uuid4())
    TEMP_DOWNLOADS[download_id] = {
        'filename': 'standard_excel_template.xlsx',
        'content': output_bytes
    }
    return download_id

def process_merge_excel_files(
    files_data,
    custom_rules=None,
    selected_columns=None,
    similarity_threshold=0.75,
    dedup_columns=None,
    dedup_keep='first',
    clean_rules=None
):
    analysis = analyze_excel_files(files_data, custom_rules, similarity_threshold)
    mapping = analysis['mapping']
    
    processed_dfs = []
    for file_info in files_data:
        fname = file_info['filename']
        content = file_info['content']
        try:
            df = read_excel_file_content(content, fname)
            df['_来源文件'] = fname
            renamed_df = df.rename(columns=mapping)
            renamed_df = renamed_df.T.groupby(level=0).first().T
            processed_dfs.append(renamed_df)
        except Exception as e:
            print(f"合并跳过异常文件 {fname}: {e}")
            
    if not processed_dfs:
        raise ValueError("没有可合并的高效文件数据")
        
    merged_df = pd.concat(processed_dfs, ignore_index=True, sort=False)
    
    all_cols = list(merged_df.columns)
    if selected_columns and len(selected_columns) > 0:
        cols_to_keep = [c for c in selected_columns if c in merged_df.columns]
        if '_来源文件' in merged_df.columns and '_来源文件' not in cols_to_keep:
            cols_to_keep.append('_来源文件')
        merged_df = merged_df[cols_to_keep]
    else:
        if '_来源文件' in all_cols:
            all_cols.remove('_来源文件')
            all_cols = ['_来源文件'] + all_cols
        merged_df = merged_df[all_cols]

    # 执行数据智能清洗与治理转换规则 (Apply Clean Rules Canvas Transformations)
    if clean_rules and isinstance(clean_rules, dict):
        for col_name, rule_cfg in clean_rules.items():
            if col_name in merged_df.columns and rule_cfg and isinstance(rule_cfg, dict):
                trim_space = rule_cfg.get('trim_space', True)
                remove_linebreaks = rule_cfg.get('remove_linebreaks', False)
                case_transform = rule_cfg.get('case_transform', 'none')
                fill_empty = rule_cfg.get('fill_empty', None)
                mask_type = rule_cfg.get('mask_type', None)
                date_format = rule_cfg.get('date_format', None)
                
                def apply_cleaning(val):
                    res = val
                    if date_format and date_format != 'none':
                        res = standardize_date_val(res, date_format)
                    if mask_type and mask_type != 'none':
                        res = mask_value(res, mask_type)
                    if trim_space or remove_linebreaks or case_transform != 'none' or fill_empty:
                        res = clean_text_val(res, trim_space, remove_linebreaks, case_transform, fill_empty)
                    return res
                    
                merged_df[col_name] = merged_df[col_name].apply(apply_cleaning)

    initial_row_count = len(merged_df)
    dedup_removed_count = 0
    
    if dedup_columns and len(dedup_columns) > 0 and dedup_keep in ['first', 'last']:
        valid_dedup_cols = [c for c in dedup_columns if c in merged_df.columns and c != '_来源文件']
        if valid_dedup_cols:
            merged_df = merged_df.drop_duplicates(subset=valid_dedup_cols, keep=dedup_keep)
            dedup_removed_count = initial_row_count - len(merged_df)

    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        merged_df.to_excel(writer, index=False, sheet_name="合并总表")
        auto_fit_excel_columns(writer, merged_df, sheet_name="合并总表")
        
    output_bytes = output_buffer.getvalue()
    
    template_download_id = generate_blank_template(list(merged_df.columns))

    preview_df = merged_df.head(10).fillna("")
    preview_data = {
        'headers': list(preview_df.columns),
        'rows': preview_df.values.tolist(),
        'total_rows': len(merged_df),
        'total_cols': len(merged_df.columns)
    }
    
    download_id = str(uuid.uuid4())
    TEMP_DOWNLOADS[download_id] = {
        'filename': 'merged_master_table.xlsx',
        'content': output_bytes
    }
    
    return {
        'download_id': download_id,
        'template_download_id': template_download_id,
        'preview': preview_data,
        'total_rows': len(merged_df),
        'total_cols': len(merged_df.columns),
        'initial_rows': initial_row_count,
        'dedup_removed_count': dedup_removed_count
    }

def process_clean_excel_file(files_data, clean_rules=None):
    """独立数据格式化与批量规范清洗引擎"""
    processed_dfs = []
    for file_info in files_data:
        fname = file_info['filename']
        content = file_info['content']
        df = read_excel_file_content(content, fname)
        
        if clean_rules and isinstance(clean_rules, dict):
            for col_name, rule_cfg in clean_rules.items():
                if col_name in df.columns and rule_cfg and isinstance(rule_cfg, dict):
                    trim_space = rule_cfg.get('trim_space', True)
                    remove_linebreaks = rule_cfg.get('remove_linebreaks', False)
                    case_transform = rule_cfg.get('case_transform', 'none')
                    fill_empty = rule_cfg.get('fill_empty', None)
                    mask_type = rule_cfg.get('mask_type', None)
                    date_format = rule_cfg.get('date_format', None)
                    
                    def apply_cleaning(val):
                        res = val
                        if date_format and date_format != 'none':
                            res = standardize_date_val(res, date_format)
                        if mask_type and mask_type != 'none':
                            res = mask_value(res, mask_type)
                        if trim_space or remove_linebreaks or case_transform != 'none' or fill_empty:
                            res = clean_text_val(res, trim_space, remove_linebreaks, case_transform, fill_empty)
                        return res
                        
                    df[col_name] = df[col_name].apply(apply_cleaning)
                    
        processed_dfs.append(df)
        
    if not processed_dfs:
        raise ValueError("没有可进行格式化清洗的有效文件数据")
        
    final_df = pd.concat(processed_dfs, ignore_index=True, sort=False) if len(processed_dfs) > 1 else processed_dfs[0]
    
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False, sheet_name="清洗规范结果")
        auto_fit_excel_columns(writer, final_df, sheet_name="清洗规范结果")
        
    output_bytes = output_buffer.getvalue()
    
    download_id = str(uuid.uuid4())
    TEMP_DOWNLOADS[download_id] = {
        'filename': 'formatted_cleaned_table.xlsx',
        'content': output_bytes
    }
    
    preview_df = final_df.head(10).fillna("")
    return {
        'download_id': download_id,
        'preview': {
            'headers': list(preview_df.columns),
            'rows': preview_df.values.tolist(),
            'total_rows': len(final_df),
            'total_cols': len(final_df.columns)
        },
        'total_rows': len(final_df),
        'total_cols': len(final_df.columns)
    }
