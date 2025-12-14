import os
import pandas as pd
import time
from collections import defaultdict

# 文件夹路径
folder_path = r"G:\优仙 工作\文档处理\input-gl"

# 支持的文件扩展名
SUPPORTED_EXTS = ['.xlsx', '.xls', '.csv', '.tsv']

# 存储所有表头信息
all_headers = []
total_rows = 0

print(f"开始分析文件夹: {folder_path}\n")

# 遍历文件夹中的所有文件
for root, dirs, files in os.walk(folder_path):
    for file in files:
        file_extension = os.path.splitext(file)[1].lower()
        file_path = os.path.join(root, file)

        if file_extension not in SUPPORTED_EXTS:
            continue

        try:
            # 获取文件大小（MB）
            file_size = os.path.getsize(file_path) / (1024 * 1024)

            # 根据文件扩展名选择读取方法
            if file_extension in ['.xlsx', '.xls']:
                # 读取第一行（表头）和行数
                df = pd.read_excel(file_path, nrows=1)
                headers = df.columns.tolist()
                num_rows = len(pd.read_excel(file_path, usecols=[0]))  # 仅读取第一列计算行数
            elif file_extension == '.csv':
                # 读取第一行（表头）和行数
                df = pd.read_csv(file_path, nrows=1)
                headers = df.columns.tolist()
                # 高效计算CSV行数
                with open(file_path, 'r') as f:
                    num_rows = sum(1 for line in f) - 1  # 减1是减去表头
            elif file_extension == '.tsv':
                # 读取第一行（表头）和行数
                df = pd.read_csv(file_path, sep='\t', nrows=1)
                headers = df.columns.tolist()
                # 高效计算TSV行数
                with open(file_path, 'r') as f:
                    num_rows = sum(1 for line in f) - 1  # 减1是减去表头

            # 记录总行数
            total_rows += num_rows

            # 记录表头信息
            all_headers.append({
                'file': file,
                'headers': headers,
                'num_columns': len(headers),
                'num_rows': num_rows
            })

            # 打印文件信息
            print(f"文件名: {file}")
            print(f"  - 第一行(表头): {headers[:5]}...")
            print(f"  - 列数: {len(headers)}")
            print(f"  - 文件大小: {file_size:.2f} MB")
            print(f"  - 行数: {num_rows:,}")
            print()

        except Exception as e:
            print(f"✗ 分析文件 {file} 时出错: {str(e)}\n")

# 分析表头一致性
if all_headers:
    # 计算表头完全一致的文件组
    header_groups = defaultdict(list)
    for info in all_headers:
        header_tuple = tuple(info['headers'])
        header_groups[header_tuple].append(info['file'])

    # 打印表头分析结果
    print("\n===== 表头分析结果 =====")
    print(f"总文件数: {len(all_headers)}")
    print(f"总记录数: {total_rows:,}")

    if len(header_groups) == 1:
        print("✅ 所有表格的表头完全一致！")
    else:
        print(f"❌ 发现 {len(header_groups)} 种不同的表头结构:")
        for i, (headers, files) in enumerate(header_groups.items(), 1):
            print(f"  {i}. 表头结构: {list(headers)[:5]}...")
            print(f"     包含文件数: {len(files)}")
            print(f"     示例文件: {files[:3]}...")
            print()
else:
    print("没有找到可分析的表格文件！")