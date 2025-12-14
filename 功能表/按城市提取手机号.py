import os
import pandas as pd
import time
from datetime import timedelta


def process_phone_numbers():
    # 路径配置
    source_dir = r"C:\Users\Administrator\Desktop\output-gl"
    target_dir = r"C:\Users\Administrator\Desktop\未筛选号码"

    if not os.path.exists(target_dir):
        os.makedirs(target_dir)

    start_total = time.time()
    processed_files = 0

    # 遍历源文件夹中的所有Excel文件
    for filename in os.listdir(source_dir):
        if filename.endswith(('.xlsx', '.xls')):
            file_start = time.time()
            province = os.path.splitext(filename)[0]
            province_dir = os.path.join(target_dir, province)

            if not os.path.exists(province_dir):
                os.makedirs(province_dir)

            file_path = os.path.join(source_dir, filename)
            print(f"\n⏳ 正在处理 [{province}] 数据...")

            try:
                # 读取Excel文件
                df = pd.read_excel(file_path, dtype={'有效手机号': str})

                # 检查必要列是否存在
                if '所属城市' not in df.columns or '有效手机号' not in df.columns:
                    print(f"⚠️ 文件 {filename} 缺少必要列，跳过处理")
                    continue

                # 处理空值
                df = df.dropna(subset=['所属城市', '有效手机号'])

                # 按城市分组处理
                city_stats = []
                for city, group in df.groupby('所属城市'):
                    # 手机号去重
                    unique_phones = group['有效手机号'].drop_duplicates().tolist()
                    count = len(unique_phones)

                    # 创建城市文件
                    city_filename = f"{province} {city}.txt"
                    city_path = os.path.join(province_dir, city_filename)

                    # 写入文件
                    with open(city_path, 'w', encoding='utf-8') as f:
                        for phone in unique_phones:
                            f.write(phone + '\n')

                    city_stats.append((city, count))

                # 显示城市统计
                file_time = timedelta(seconds=round(time.time() - file_start))
                print(f"✅ [{province}] 处理完成 | 用时: {file_time}")
                for city, count in city_stats:
                    print(f"  ├─ {city}: {count}个去重号码")

                processed_files += 1

            except Exception as e:
                print(f"❌ 处理 {filename} 时出错: {str(e)}")

    # 最终统计
    total_time = timedelta(seconds=round(time.time() - start_total))
    print(f"\n{'=' * 50}")
    print(f"📊 任务完成! 共处理 {processed_files} 个省份文件")
    print(f"⏱️ 总用时: {total_time}")
    print(f"📂 结果保存在: {target_dir}")
    print('=' * 50)


if __name__ == "__main__":
    process_phone_numbers()