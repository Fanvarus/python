import os
import pandas as pd
import shutil
from collections import defaultdict

# 输入和输出文件夹路径
input_folder = r"G:\优仙 工作\文档处理\input-gl"
output_folder = r"G:\优仙 工作\文档处理\output-gl"

# 确保输出文件夹存在
os.makedirs(output_folder, exist_ok=True)


def process_files():
    # 存储每个省份的数据
    province_data = defaultdict(list)

    # 存储每个省份的文件格式
    province_formats = {}

    # 处理每个文件
    for filename in os.listdir(input_folder):
        if filename.endswith(('.xls', '.xlsx', '.csv')):
            file_path = os.path.join(input_folder, filename)

            try:
                # 读取文件
                if filename.endswith('.csv'):
                    df = pd.read_csv(file_path)
                    file_format = 'csv'
                else:
                    df = pd.read_excel(file_path)
                    file_format = 'excel'

                # 检查是否有"所属省份"列
                if "所属省份" in df.columns:
                    # 按省份分组
                    for province, group in df.groupby("所属省份"):
                        # 处理省份名称中的非法字符
                        clean_province = province.replace('/', '_').replace('\\', '_').replace(':', '_').replace('*',
                                                                                                                 '_').replace(
                            '?', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_')

                        # 存储省份数据
                        province_data[clean_province].append(group)

                        # 记录文件格式（使用第一个文件的格式）
                        if clean_province not in province_formats:
                            province_formats[clean_province] = file_format

                    print(f"已处理: {filename}")
                else:
                    print(f"错误: 文件 {filename} 中没有'所属省份'列")
            except Exception as e:
                print(f"处理文件 {filename} 时出错: {str(e)}")

    # 保存每个省份的数据
    for province, data_list in province_data.items():
        # 合并该省份的所有数据
        combined_df = pd.concat(data_list)

        # 确定文件格式
        file_format = province_formats.get(province, 'csv')

        # 创建输出路径
        if file_format == 'csv':
            output_path = os.path.join(output_folder, f"{province}.csv")
            combined_df.to_csv(output_path, index=False, encoding='utf-8-sig')
        else:
            output_path = os.path.join(output_folder, f"{province}.xlsx")
            combined_df.to_excel(output_path, index=False)

        print(f"已创建省份文件: {province}.{'csv' if file_format == 'csv' else 'xlsx'}")

    print(f"\n处理完成! 共创建了 {len(province_data)} 个省份文件。")


if __name__ == "__main__":
    process_files()