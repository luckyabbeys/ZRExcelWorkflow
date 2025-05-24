# -*- coding: utf-8 -*-
"""
处理第三个sheet: Lis03_新冠感染

此脚本用于处理原始数据中的新冠感染相关信息，
合并到目标Excel的Lis03_新冠感染表单中。
"""

import pandas as pd
import os
import sys
import re

# 添加项目根目录到系统路径，以便导入utils模块
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from utils.excel_utils import optimize_time_format, save_to_excel
from utils.data_utils import find_column_by_keywords, clean_column_names

def process(source_file, target_file):
    """
    处理原始数据中的新冠感染相关信息，合并到目标Excel的Lis03_新冠感染表单中
    
    参数:
        source_file (str): 源Excel文件路径
        target_file (str): 目标Excel文件路径
    """
    try:
        # 读取源文件中的相关信息
        excel_file = pd.ExcelFile(source_file)
        
        # 检查源文件是否包含必要的表单
        required_sheets = ['门急诊信息', '住院信息', '检查信息']
        for sheet in required_sheets:
            if sheet not in excel_file.sheet_names:
                raise ValueError(f"源文件中缺少{sheet}表单")
        
        outpatient_df = excel_file.parse('门急诊信息')
        inpatient_df = excel_file.parse('住院信息')
        examination_df = excel_file.parse('检查信息')
        
        # 清理列名
        outpatient_df = clean_column_names(outpatient_df)
        inpatient_df = clean_column_names(inpatient_df)
        examination_df = clean_column_names(examination_df)
        
        # 检查目标文件是否存在
        if os.path.exists(target_file):
            # 如果目标文件存在，读取Lis03_新冠感染表单
            target_excel_file = pd.ExcelFile(target_file)
            if 'Lis03_新冠感染' in target_excel_file.sheet_names:
                target_df = target_excel_file.parse('Lis03_新冠感染')
                # 获取目标表单的表头（第一行）
                target_header = target_df.columns.tolist()
            else:
                # 如果目标文件存在但没有Lis03_新冠感染表单，创建一个空的DataFrame
                target_header = create_default_header()
                target_df = pd.DataFrame(columns=target_header)
        else:
            # 如果目标文件不存在，创建一个空的DataFrame
            target_header = create_default_header()
            target_df = pd.DataFrame(columns=target_header)
        
        # 从诊断信息中提取新冠感染相关信息
        covid_info = []
        
        # 处理门急诊信息
        covid_info.extend(extract_covid_info(outpatient_df, target_header, '门急诊'))
        
        # 处理住院信息
        covid_info.extend(extract_covid_info(inpatient_df, target_header, '住院'))
        
        # 处理检查信息中的新冠检测结果
        covid_info.extend(extract_covid_test_info(examination_df, target_header))
        
        # 如果有数据，直接创建DataFrame，不与空DataFrame拼接
        if covid_info:
            merged_df = pd.DataFrame(covid_info)
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
        
        # 将合并后的数据保存到目标文件
        save_to_excel(merged_df, target_file, 'Lis03_新冠感染', index=False, header=True)
        
        print("Lis03_新冠感染表处理完成！")
        return True
        
    except Exception as e:
        print(f"处理Lis03_新冠感染表时发生错误: {e}")
        return False

def extract_covid_info(df, target_header, visit_type):
    """
    从DataFrame中提取新冠感染相关信息
    
    参数:
        df (pandas.DataFrame): 源数据DataFrame
        target_header (list): 目标表头列表
        visit_type (str): 就诊类型值
        
    返回:
        list: 新冠感染信息行列表
    """
    covid_info = []
    
    # 定义新冠相关关键词
    covid_keywords = ['新冠', '冠状病毒', 'COVID', 'SARS-CoV-2', '新型冠状']
    
    # 查找诊断列
    diagnosis_cols = find_column_by_keywords(df, ['诊断', '疾病', '病名'])
    if not diagnosis_cols:
        return []
    
    # 查找患者ID列
    patient_id_cols = find_column_by_keywords(df, ['患者ID', '病人ID', '就诊ID'])
    if not patient_id_cols:
        return []
    
    # 查找就诊日期列
    visit_date_cols = find_column_by_keywords(df, ['就诊日期', '开始日期', '门诊日期'])
    
    # 使用找到的第一个列
    diagnosis_col = diagnosis_cols[0]
    patient_id_col = patient_id_cols[0]
    visit_date_col = visit_date_cols[0] if visit_date_cols else None
    
    # 遍历数据行
    for _, row in df.iterrows():
        # 检查诊断是否包含新冠相关关键词
        diagnosis = str(row[diagnosis_col]) if pd.notna(row[diagnosis_col]) else ''
        
        if any(keyword in diagnosis for keyword in covid_keywords):
            # 创建新行
            new_row = {}
            
            # 填充目标表头中的列
            for col in target_header:
                # 根据关键字匹配列
                if any(keyword in col.lower() for keyword in ['患者', '病人', '就诊']):
                    new_row[col] = row[patient_id_col]
                elif any(keyword in col.lower() for keyword in ['就诊类型', '类型', '来源']):
                    new_row[col] = visit_type
                elif any(keyword in col.lower() for keyword in ['就诊日期', '开始日期', '门诊日期']) and visit_date_col:
                    new_row[col] = row[visit_date_col]
                elif any(keyword in col.lower() for keyword in ['诊断', '疾病', '病名']):
                    new_row[col] = diagnosis
                elif any(keyword in col.lower() for keyword in ['感染状态', '感染情况', '感染结果']):
                    new_row[col] = '确诊'
                elif col in df.columns:
                    new_row[col] = row[col]
                else:
                    new_row[col] = None
            
            # 添加到结果列表
            covid_info.append(new_row)
    
    return covid_info

def extract_covid_test_info(df, target_header):
    """
    从检查信息中提取新冠检测结果
    
    参数:
        df (pandas.DataFrame): 检查信息DataFrame
        target_header (list): 目标表头列表
        
    返回:
        list: 新冠检测信息行列表
    """
    covid_test_info = []
    
    # 定义新冠检测相关关键词
    covid_test_keywords = ['新冠', '冠状病毒', 'COVID', 'SARS-CoV-2', '核酸', 'PCR', '抗原']
    
    # 查找检查名称列
    test_name_cols = find_column_by_keywords(df, ['检查名称', '检验名称', '项目名称'])
    if not test_name_cols:
        return []
    
    # 查找检查结果列
    test_result_cols = find_column_by_keywords(df, ['检查结果', '检验结果', '结果'])
    if not test_result_cols:
        return []
    
    # 查找患者ID列
    patient_id_cols = find_column_by_keywords(df, ['患者ID', '病人ID', '就诊ID'])
    if not patient_id_cols:
        return []
    
    # 查找检查日期列
    test_date_cols = find_column_by_keywords(df, ['检查日期', '检验日期', '日期'])
    
    # 使用找到的第一个列
    test_name_col = test_name_cols[0]
    test_result_col = test_result_cols[0]
    patient_id_col = patient_id_cols[0]
    test_date_col = test_date_cols[0] if test_date_cols else None
    
    # 遍历数据行
    for _, row in df.iterrows():
        # 检查名称是否包含新冠相关关键词
        test_name = str(row[test_name_col]) if pd.notna(row[test_name_col]) else ''
        
        if any(keyword in test_name for keyword in covid_test_keywords):
            # 创建新行
            new_row = {}
            
            # 获取检测结果
            test_result = str(row[test_result_col]) if pd.notna(row[test_result_col]) else ''
            
            # 判断感染状态
            infection_status = '未知'
            if re.search(r'阳性|检出|positive', test_result, re.IGNORECASE):
                infection_status = '确诊'
            elif re.search(r'阴性|未检出|negative', test_result, re.IGNORECASE):
                infection_status = '排除'
            
            # 填充目标表头中的列
            for col in target_header:
                # 根据关键字匹配列
                if any(keyword in col.lower() for keyword in ['患者', '病人', '就诊']):
                    new_row[col] = row[patient_id_col]
                elif any(keyword in col.lower() for keyword in ['就诊类型', '类型', '来源']):
                    new_row[col] = '检查'
                elif any(keyword in col.lower() for keyword in ['检查日期', '检验日期']) and test_date_col:
                    new_row[col] = row[test_date_col]
                elif any(keyword in col.lower() for keyword in ['就诊日期', '开始日期', '门诊日期']) and test_date_col:
                    new_row[col] = row[test_date_col]
                elif any(keyword in col.lower() for keyword in ['检查名称', '检验名称', '项目名称']):
                    new_row[col] = test_name
                elif any(keyword in col.lower() for keyword in ['检查结果', '检验结果', '结果']):
                    new_row[col] = test_result
                elif any(keyword in col.lower() for keyword in ['感染状态', '感染情况', '感染结果']):
                    new_row[col] = infection_status
                elif col in df.columns:
                    new_row[col] = row[col]
                else:
                    new_row[col] = None
            
            # 添加到结果列表
            covid_test_info.append(new_row)
    
    return covid_test_info

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
        '感染状态', '检测方法', '检测结果', '检测日期', '症状',
        '严重程度', '治疗方案', '数据来源', '数据更新时间'
    ]

if __name__ == "__main__":
    # 如果直接运行此脚本，使用默认路径
    script_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    source_file = os.path.join(script_dir, 'data', 'input', '测试原始数据.xlsx')
    target_file = os.path.join(script_dir, 'data', 'output', '测试合并.xlsx')
    
    process(source_file, target_file)