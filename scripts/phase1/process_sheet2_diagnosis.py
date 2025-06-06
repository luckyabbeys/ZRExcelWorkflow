# -*- coding: utf-8 -*-
"""
处理第二个sheet: Lis02_诊断

此脚本用于处理原始数据中的诊断信息，
合并到目标Excel的Lis02_诊断表单中。
"""

import pandas as pd
import os
import sys

# 添加项目根目录到系统路径，以便导入utils模块
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from utils.excel_utils import optimize_time_format, save_to_excel
from utils.data_utils import find_column_by_keywords, clean_column_names

def process(source_file, target_file):
    # 执行前删除已存在的目标文件，防止旧数据干扰
    if os.path.exists(target_file):
        os.remove(target_file)
    """
    处理原始数据中的诊断信息，合并到目标Excel的Lis02_诊断表单中
    
    参数:
        source_file (str): 源Excel文件路径
        target_file (str): 目标Excel文件路径
    """
    try:
        # 读取源文件中的门急诊信息和住院信息
        excel_file = pd.ExcelFile(source_file)
        
        # 检查源文件是否包含必要的表单
        required_sheets = ['门急诊信息', '住院信息']
        for sheet in required_sheets:
            if sheet not in excel_file.sheet_names:
                raise ValueError(f"源文件中缺少{sheet}表单")
        
        outpatient_df = excel_file.parse('门急诊信息')
        inpatient_df = excel_file.parse('住院信息')
        
        # 清理列名
        outpatient_df = clean_column_names(outpatient_df)
        inpatient_df = clean_column_names(inpatient_df)
        
        # 检查目标文件是否存在
        if os.path.exists(target_file):
            # 如果目标文件存在，读取Lis02_诊断表单
            target_excel_file = pd.ExcelFile(target_file)
            if 'Lis02_诊断' in target_excel_file.sheet_names:
                target_df = target_excel_file.parse('Lis02_诊断')
                # 获取目标表单的表头（第一行）
                target_header = target_df.columns.tolist()
            else:
                # 如果目标文件存在但没有Lis02_诊断表单，创建一个空的DataFrame
                target_header = create_default_header()
                target_df = pd.DataFrame(columns=target_header)
        else:
            # 如果目标文件不存在，创建一个空的DataFrame
            target_header = create_default_header()
            target_df = pd.DataFrame(columns=target_header)
        
        # 定义需要查找的列名关键字
        diagnosis_keywords = ['诊断', '疾病', '病名']
        diagnosis_code_keywords = ['诊断编码', '疾病编码', 'ICD']
        patient_id_keywords = ['患者ID', '病人ID', '就诊ID']
        visit_type_keywords = ['就诊类型', '类型', '来源']
        visit_date_keywords = ['就诊日期', '开始日期', '门诊日期']
        
        # 从门急诊数据中提取诊断信息
        outpatient_diagnosis = extract_diagnosis_info(
            outpatient_df, 
            target_header,
            diagnosis_keywords,
            diagnosis_code_keywords,
            patient_id_keywords,
            visit_type_keywords,
            visit_date_keywords,
            '门急诊'
        )
        
        # 从住院数据中提取诊断信息
        inpatient_diagnosis = extract_diagnosis_info(
            inpatient_df, 
            target_header,
            diagnosis_keywords,
            diagnosis_code_keywords,
            patient_id_keywords,
            visit_type_keywords,
            visit_date_keywords,
            '住院'
        )
        
        # 合并所有诊断信息
        all_diagnosis = outpatient_diagnosis + inpatient_diagnosis
        
        # 如果有数据，直接创建DataFrame，不与空DataFrame拼接
        if all_diagnosis:
            merged_df = pd.DataFrame(all_diagnosis)
            # 确保列顺序与目标表头一致
            for col in target_header:
                if col not in merged_df.columns:
                    merged_df[col] = None
            merged_df = merged_df[target_header]
        else:
            # 如果没有数据，创建一个只有表头的DataFrame
            merged_df = pd.DataFrame(columns=target_header)
        
        # 优化时间格式
        merged_df = optimize_time_format(merged_df)
        
        # 读取phase1sheet1处理后的数据
        phase1sheet1_file = 'data/output/原始数据合并phase1sheet1.xlsx'
        phase1sheet1_excel = pd.ExcelFile(phase1sheet1_file)
        lis01_df = phase1sheet1_excel.parse('Lis01_就诊合并')

        # 复制A列至N列到新的DataFrame
        lis02_df = lis01_df.iloc[:, 0:14].copy()

        # 添加O列：新冠患者
        lis02_df['新冠患者'] = lis01_df['诊断（文字）'].str.contains('新型冠状病毒感染').map({True: '是', False: '否'})

        # 添加P列：RS患者
        lis02_df['RS患者'] = lis01_df['诊断（ICD编码）'].apply(lambda x: '是' if any(code in str(x) for code in RS_CODES) else '否')

        # 添加Q列：诊断合并为肺炎
        lis02_df['诊断合并为肺炎'] = lis02_df.apply(lambda row: '是' if row['RS患者'] == '是' and ('肺部感染' in str(row['诊断（文字）']) or '肺炎' in str(row['诊断（文字）'])) else '否', axis=1)

        # 读取原目标文件的所有sheet
        if os.path.exists(target_file):
            original_excel = pd.ExcelFile(target_file)
            original_sheets = original_excel.sheet_names
        else:
            original_sheets = []

        # 写入所有原sheet和新的Lis02_诊断sheet
        with pd.ExcelWriter(target_file, engine='openpyxl') as writer:
            for sheet in original_sheets:
                original_excel.parse(sheet).to_excel(writer, sheet_name=sheet, index=False)
            lis02_df.to_excel(writer, sheet_name='Lis02_诊断', index=False)
        
        print("Lis02_诊断表处理完成！")
        return True
        
    except Exception as e:
        print(f"处理Lis02_诊断表时发生错误: {e}")
        return False

def extract_diagnosis_info(df, target_header, diagnosis_keywords, diagnosis_code_keywords, 
                          patient_id_keywords, visit_type_keywords, visit_date_keywords, visit_type):
    """
    从DataFrame中提取诊断信息
    
    参数:
        df (pandas.DataFrame): 源数据DataFrame
        target_header (list): 目标表头列表
        diagnosis_keywords (list): 诊断列关键字列表
        diagnosis_code_keywords (list): 诊断编码列关键字列表
        patient_id_keywords (list): 患者ID列关键字列表
        visit_type_keywords (list): 就诊类型列关键字列表
        visit_date_keywords (list): 就诊日期列关键字列表
        visit_type (str): 就诊类型值
        
    返回:
        list: 诊断信息行列表
    """
    diagnosis_info = []
    
    # 查找相关列
    diagnosis_cols = find_column_by_keywords(df, diagnosis_keywords)
    diagnosis_code_cols = find_column_by_keywords(df, diagnosis_code_keywords)
    patient_id_cols = find_column_by_keywords(df, patient_id_keywords)
    visit_date_cols = find_column_by_keywords(df, visit_date_keywords)
    
    # 如果找不到必要的列，返回空列表
    if not diagnosis_cols or not patient_id_cols:
        return []
    
    # 使用找到的第一个列
    diagnosis_col = diagnosis_cols[0] if diagnosis_cols else None
    diagnosis_code_col = diagnosis_code_cols[0] if diagnosis_code_cols else None
    patient_id_col = patient_id_cols[0] if patient_id_cols else None
    visit_date_col = visit_date_cols[0] if visit_date_cols else None
    
    # 遍历数据行
    for _, row in df.iterrows():
        # 创建新行
        new_row = {}
        
        # 填充目标表头中的列
        for col in target_header:
            # 根据关键字匹配列
            if any(keyword in col.lower() for keyword in ['诊断', '疾病', '病名']) and diagnosis_col:
                new_row[col] = row[diagnosis_col]
            elif any(keyword in col.lower() for keyword in ['编码', 'icd']) and diagnosis_code_col:
                new_row[col] = row[diagnosis_code_col]
            elif any(keyword in col.lower() for keyword in ['患者', '病人', '就诊']) and patient_id_col:
                new_row[col] = row[patient_id_col]
            elif any(keyword in col.lower() for keyword in ['就诊类型', '类型', '来源']):
                new_row[col] = visit_type
            elif any(keyword in col.lower() for keyword in ['就诊日期', '开始日期', '门诊日期']) and visit_date_col:
                new_row[col] = row[visit_date_col]
            elif col in df.columns:
                new_row[col] = row[col]
            else:
                new_row[col] = None
        
        # 添加到结果列表
        diagnosis_info.append(new_row)
    
    return diagnosis_info

def create_default_header():
    """
    创建默认的表头
    
    返回:
        list: 默认表头列表
    """
    # 这里可以根据实际需求定义默认表头
    # 以下是一个示例，实际应用中应根据业务需求调整
    return [
        '患者ID', '姓名', '就诊类型', '就诊日期', '诊断', 
        '诊断编码', '诊断类型', '诊断医生', '诊断科室', '诊断状态',
        '数据来源', '数据更新时间', '诊断（ICD编码）', '诊断（文字）', '新冠患者', 'RS患者', '诊断合并为肺炎'
    ]

import os
import pandas as pd

# 定义RS患者的ICD编码列表
RS_CODES = ['R50', 'J00', 'J01', 'J02', 'J03', 'J04', 'J06', 'J07', 'J08', 'J09', 'J10', 'J11', 'J12', 'J13', 'J14', 'J15', 'J16', 'J17', 'J18', 'J20', 'J21', 'J98.414', 'J98.802']

# 配置常量（根据项目规范）
RS_CODES = {'R50', 'J00', 'J01', 'J02', 'J03', 'J04', 'J06', 'J07', 'J08', 'J09', 'J10', 'J11', 'J12', 'J13', 'J14', 'J15', 'J16', 'J17', 'J18', 'J20', 'J21', 'J98.414', 'J98.802'}

def process_diagnosis_sheet(input_path, output_path):
    """
    处理Lis02_诊断sheet的主函数（保留原sheet并添加新sheet）
    参数:
        input_path: 输入文件路径（phase1sheet1输出文件）
        output_path: 输出文件路径（phase1sheet2目标文件）
    """
    try:
        # 读取输入文件的所有sheet（保留原数据）
        with pd.ExcelFile(input_path) as xls:
            sheets = {sheet_name: xls.parse(sheet_name) for sheet_name in xls.sheet_names}
            lis01_df = sheets['Lis01_就诊合并']  # 获取需要处理的源数据

        # 步骤1：复制A-N列生成Lis02_诊断基础数据（A=0列，N=13列）
        lis02_df = lis01_df.iloc[:, 0:14].copy()

        # 步骤2：添加O列-新冠患者判断（M列是诊断文字，索引12）
        lis02_df['新冠患者'] = lis01_df['诊断（文字）'].apply(
            lambda x: '是' if '新型冠状病毒感染' in str(x) else '否'
        )

        # 步骤3：添加P列-RS患者判断（L列是ICD编码，索引11）
        lis02_df['RS患者'] = lis01_df['诊断（ICD编码）'].apply(
            lambda x: '是' if str(x) in RS_CODES else '否'
        )

        # 步骤4：添加Q列-诊断合并为肺炎判断（仅当P列为'是'时检查）
        def check_pneumonia(row):
            if row['RS患者'] == '是':
                diag_text = str(row['诊断（文字）'])
                return '是' if ('肺部感染' in diag_text) or ('肺炎' in diag_text) else '否'
            return '否'  # P列不是'是'时填否
        lis02_df['诊断合并为肺炎'] = lis01_df.apply(check_pneumonia, axis=1)

        # 步骤5：写入新sheet并保留原sheet（使用openpyxl引擎追加模式）
        with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            # 写入原有的所有sheet
            for sheet_name, df in sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            # 写入新的Lis02_诊断sheet（覆盖或新增）
            lis02_df.to_excel(writer, sheet_name='Lis02_诊断', index=False)

        print(f"Lis02_诊断处理完成，结果保存至：{output_path}（含原sheet1和新sheet2）")

    except Exception as e:
        print(f"处理Lis02_诊断时发生错误：{str(e)}")

if __name__ == '__main__':
    # 获取项目根目录
    script_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    
    # 输入文件为phase1sheet1处理结果
    input_file = os.path.join(script_dir, 'data', 'output', '原始数据合并phase1sheet1.xlsx')
    
    # 生成输出文件名（phase1sheet2）
    input_file_name = os.path.splitext(os.path.basename(input_file))[0].replace('phase1sheet1', '')
    phase_step = 'phase1sheet2'
    output_file = os.path.join(script_dir, 'data', 'output', f'{input_file_name}合并{phase_step}.xlsx')

    # 执行处理（若输出文件已存在则先删除避免冲突）
    if os.path.exists(output_file):
        os.remove(output_file)
    process_diagnosis_sheet(input_path=input_file, output_path=output_file)
    
    # 如果直接运行此脚本，使用默认路径
    source_file = os.path.join(script_dir, 'data', 'input', '原始数据.xlsx')
    target_file = os.path.join(script_dir, 'data', 'output', '测试合并.xlsx')
    
    process(source_file, target_file)