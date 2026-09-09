import os
import glob
import re
from difflib import SequenceMatcher
import pandas as pd

def normalize_title(title):
    """清理字符串：去除空格、标点符号、全半角规范化等"""
    if not isinstance(title, str):
        title = str(title)
    # 移除常见修饰符如（必填）、*号、空格
    title = re.sub(r'[\s\*\（\）\(\)\【\】\_\-\:\：]', '', title)
    return title.lower()

def title_similarity(t1, t2):
    """计算两个标题的相似度 (0.0 - 1.0)"""
    n1 = normalize_title(t1)
    n2 = normalize_title(t2)
    if n1 == n2:
        return 1.0
    if n1 in n2 or n2 in n1:
        # 子串匹配赋予较高权重
        return 0.85 + 0.15 * (min(len(n1), len(n2)) / max(len(n1), len(n2)))
    return SequenceMatcher(None, n1, n2).ratio()

def build_column_mapping(all_headers, custom_mapping=None, similarity_threshold=0.7):
    """
    根据自定义映射表和相似度分析，构建列名映射字典
    """
    mapping = {}
    mapped_originals = set()

    # 1. 优先处理自定义映射表参数
    # custom_mapping 格式例：{'客户姓名': ['姓名', '买家姓名', '客户'], '联系电话': ['电话', '手机', '手机号']}
    if custom_mapping:
        for std_col, aliases in custom_mapping.items():
            normalized_aliases = [normalize_title(a) for a in aliases] + [normalize_title(std_col)]
            for header in all_headers:
                norm_h = normalize_title(header)
                if norm_h in normalized_aliases or any(norm_h == a for a in normalized_aliases):
                    mapping[header] = std_col
                    mapped_originals.add(header)

    # 2. 对未手动映射的标题进行自动相似度聚类
    unmapped_headers = [h for h in all_headers if h not in mapped_originals]
    clusters = []
    
    for h in unmapped_headers:
        assigned = False
        for cluster in clusters:
            rep = cluster[0]
            if title_similarity(h, rep) >= similarity_threshold:
                cluster.append(h)
                assigned = True
                break
        if not assigned:
            clusters.append([h])

    for cluster in clusters:
        # 选择最短且最规范的名称作为代表标准列名
        std_col = sorted(cluster, key=lambda x: (len(normalize_title(x)), x))[0]
        for h in cluster:
            mapping[h] = std_col
            
    return mapping

def merge_excel_folder(
    folder_path,
    custom_mapping=None,
    selected_columns=None,
    similarity_threshold=0.70,
    output_file="merged_result.xlsx"
):
    """
    智能合并文件夹中的多个 Excel 文件
    
    :param folder_path: 存放 Excel 文件的文件夹路径
    :param custom_mapping: 自定义列名映射规则 dict，如 {'客户姓名': ['姓名', '买家'], '联系电话': ['手机号', '电话']}
    :param selected_columns: 需保留的输出列名 list，如果不填（None）则自动保留所有归并列
    :param similarity_threshold: 相似度匹配阈值 (0.0-1.0)，默认 0.70
    :param output_file: 合并结果保存路径
    """
    excel_files = glob.glob(os.path.join(folder_path, "*.xlsx")) + glob.glob(os.path.join(folder_path, "*.xls"))
    # 过滤掉输出文件本身
    excel_files = [f for f in excel_files if os.path.basename(f) != os.path.basename(output_file)]
    
    if not excel_files:
        print(f"在 {folder_path} 中未找到 Excel 文件")
        return None
        
    print(f"找到 {len(excel_files)} 个 Excel 文件，开始分析表头...")
    
    file_dfs = []
    all_headers = set()
    
    for fpath in excel_files:
        try:
            df = pd.read_excel(fpath)
            df['_来源文件'] = os.path.basename(fpath)
            file_dfs.append(df)
            for col in df.columns:
                if col != '_来源文件':
                    all_headers.add(str(col))
        except Exception as e:
            print(f"读取文件 {fpath} 失败: {e}")
            
    # 构建映射关系
    mapping = build_column_mapping(all_headers, custom_mapping, similarity_threshold)
    print("\n=== 列标题智能归并映射关系 ===")
    for original, std in mapping.items():
        print(f"  原始列名: [{original:15s}]  ===>  统一映射为: [{std}]")
    print("===============================\n")
    
    processed_dfs = []
    for df in file_dfs:
        renamed_df = df.rename(columns=mapping)
        # 如果重命名后产生重复列名，按列合并其非空数据
        renamed_df = renamed_df.T.groupby(level=0).first().T
        processed_dfs.append(renamed_df)
        
    merged_df = pd.concat(processed_dfs, ignore_index=True, sort=False)
    
    # 选择保留列
    if selected_columns:
        cols_to_keep = [c for c in selected_columns if c in merged_df.columns]
        if '_来源文件' in merged_df.columns and '_来源文件' not in cols_to_keep:
            cols_to_keep.append('_来源文件')
        merged_df = merged_df[cols_to_keep]
    else:
        cols = list(merged_df.columns)
        if '_来源文件' in cols:
            cols.remove('_来源文件')
            cols = ['_来源文件'] + cols
        merged_df = merged_df[cols]
        
    merged_df.to_excel(output_file, index=False)
    print(f"[SUCCESS] 合并成功！总计处理 {len(file_dfs)} 个文件，拼接 {len(merged_df)} 行数据。保存至: {output_file}")
    return merged_df

if __name__ == "__main__":
    # 创建测试数据演示
    os.makedirs("./sample_excels", exist_ok=True)
    
    df1 = pd.DataFrame({"姓名": ["张三", "李四"], "手机号": ["13800000001", "13800000002"], "成交金额": [100, 200]})
    df2 = pd.DataFrame({"客户姓名": ["王五", "赵六"], "联系电话": ["13800000003", "13800000004"], "金额(元)": [300, 400], "备注": ["老客户", "新客户"]})
    df3 = pd.DataFrame({"名字": ["钱七"], "电话": ["13800000005"], "销售额": [500]})
    
    df1.to_excel("./sample_excels/门店A.xlsx", index=False)
    df2.to_excel("./sample_excels/门店B.xlsx", index=False)
    df3.to_excel("./sample_excels/门店C.xlsx", index=False)
    
    print("--- 场景 1：完全自动分析合并（无需配置参数） ---")
    merge_excel_folder("./sample_excels", output_file="merged_auto.xlsx")
    
    print("\n--- 场景 2：自定义参数映射 + 选定保留列 ---")
    custom_map = {
        "客户名称": ["姓名", "客户姓名", "名字"],
        "联系电话": ["手机号", "联系电话", "电话"],
        "总金额": ["成交金额", "金额(元)", "销售额"]
    }
    select_cols = ["客户名称", "联系电话", "总金额"] # 仅保留这三列
    
    merge_excel_folder(
        "./sample_excels",
        custom_mapping=custom_map,
        selected_columns=select_cols,
        output_file="merged_custom.xlsx"
    )
