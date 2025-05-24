import pandas as pd
import os

def optimize_time_format(df):
    # 定义需要处理的列名关键字
    time_column_keywords = ['就诊日期', '开始日期', '门诊日期', '就诊结束日期', '结束日期', '出院日期', '入院日期', '住院开始日期', '住院结束日期']
    # 遍历 DataFrame 的列
    for col in df.columns:
        # 检查列名是否包含时间关键字
        if any(keyword in col for keyword in time_column_keywords):
            try:
                # 尝试将列转换为日期时间类型
                df[col] = pd.to_datetime(df[col], errors='coerce')
                # 将 NaT 替换为空字符串
                df[col] = df[col].apply(lambda x: '' if pd.isna(x) else x)
                # 检查时间是否为 00:00:00
                mask = df[col].apply(lambda x: x.time() == pd.Timestamp('00:00:00').time() if isinstance(x, pd.Timestamp) else False)
                # 对于时间为 00:00:00 的数据，只保留日期部分
                df.loc[mask, col] = df.loc[mask, col].apply(lambda x: x.strftime('%Y-%m-%d') if isinstance(x, pd.Timestamp) else x)
                # 将整列转换为字符串类型
                df[col] = df[col].astype(str)
                # 对于其他数据，保持原日期时间格式
                df.loc[~mask, col] = df.loc[~mask, col].apply(lambda x: pd.Timestamp(x).strftime('%Y-%m-%d %H:%M:%S') if pd.notna(x) and x != 'NaT' and x != '' else x)
            except ValueError:
                continue
    return df

def merge_excel_data():
    # 获取脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # 定义文件路径
    source_file = os.path.join(script_dir, '测试原始数据.xlsx')
    target_file = os.path.join(script_dir, '测试合并.xlsx')

    try:
        # 读取源文件中的门急诊信息和住院信息
        excel_file = pd.ExcelFile(source_file)
        outpatient_df = excel_file.parse('门急诊信息')
        inpatient_df = excel_file.parse('住院信息')

        # 读取目标文件中的Lis01_就诊合并表单
        target_excel_file = pd.ExcelFile(target_file)
        target_df = target_excel_file.parse('Lis01_就诊合并')

        # 获取目标表单的表头（第一行）
        target_header = target_df.columns.tolist()

        # 定义需要查找的列名关键字
        column_keywords = {
            'E列': ['就诊类型', '类型', '来源'],
            'J列': ['就诊日期', '开始日期', '门诊日期'],
            'K列': ['就诊结束日期', '结束日期', '出院日期'],
            'U列': ['入院日期', '住院开始日期'],
            'V列': ['出院日期', '住院结束日期']
        }

        # 动态查找列索引
        column_indices = {}
        for col_name, keywords in column_keywords.items():
            found = False
            for idx, header in enumerate(target_header):
                if any(keyword in header for keyword in keywords):
                    column_indices[col_name] = idx
                    found = True
                    break
            if not found:
                raise ValueError(f"找不到包含以下关键词的{col_name}: {', '.join(keywords)}")

        # 确保目标表头包含所需列
        required_cols = ['J列', 'K列', 'U列', 'V列']
        if any(col not in column_indices for col in required_cols):
            missing = [col for col in required_cols if col not in column_indices]
            raise ValueError(f"找不到以下列: {', '.join(missing)}")

        # 处理门急诊信息
        outpatient_rows = []
        for _, row in outpatient_df.iterrows():
            new_row = {}
            for col in target_header:
                if col in outpatient_df.columns:
                    new_row[col] = row[col]
                else:
                    new_row[col] = None
            # 标记数据来源为门急诊
            new_row[target_header[column_indices['E列']]] = '门急诊'
            outpatient_rows.append(new_row)

        # 处理住院信息
        inpatient_rows = []
        for _, row in inpatient_df.iterrows():
            new_row = {}
            for col in target_header:
                if col in inpatient_df.columns:
                    new_row[col] = row[col]
                else:
                    new_row[col] = None
            # 标记数据来源为住院
            new_row[target_header[column_indices['E列']]] = '住院'

            # 迁移住院日期信息
            admission_date = row.get(target_header[column_indices['U列']])
            discharge_date = row.get(target_header[column_indices['V列']])

            new_row[target_header[column_indices['J列']]] = admission_date
            new_row[target_header[column_indices['K列']]] = discharge_date
            new_row[target_header[column_indices['U列']]] = None
            new_row[target_header[column_indices['V列']]] = None

            inpatient_rows.append(new_row)

        # 合并所有行
        all_rows = outpatient_rows + inpatient_rows

        # 如果有数据，直接创建DataFrame，不与空DataFrame拼接
        if all_rows:
            merged_df = pd.DataFrame(all_rows)
            # 确保列顺序与目标表头一致
            merged_df = merged_df[target_header]
        else:
            # 如果没有数据，创建一个只有表头的DataFrame
            merged_df = pd.DataFrame(columns=target_header)

        # 优化时间格式
        merged_df = optimize_time_format(merged_df)

        # 确保目标表单的第一行表头不被修改
        final_df = pd.DataFrame([target_header])
        final_df.columns = target_header
        if not merged_df.empty:
            final_df = pd.concat([final_df, merged_df], ignore_index=True)

        # 将合并后的数据写入目标文件的Lis01_就诊合并表单
        with pd.ExcelWriter(target_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            final_df.to_excel(writer, sheet_name='Lis01_就诊合并', index=False, header=False)

        print("数据合并完成！")

    except Exception as e:
        print(f"发生错误: {e}")

if __name__ == "__main__":
    merge_excel_data()