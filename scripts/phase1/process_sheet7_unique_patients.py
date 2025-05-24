# -*- coding: utf-8 -*-
"""
处理第七个sheet: LisA1_唯一患者

此脚本用于处理原始数据中的患者信息，
提取唯一患者记录并合并到目标Excel的LisA1_唯一患者表单中。
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
    处理原始数据中的患者信息，提取唯一患者记录并合并到目标Excel的LisA1_唯一患者表单中
    
    参数:
        source_file (str): 源Excel文件路径
        target_file (str): 目标Excel文件路径
    """
    try:
        # 读取源文件中的相关信息
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
            # 如果目标文件存在，读取LisA1_唯一患者表单
            target_excel_file = pd.ExcelFile(target_file)
            if 'LisA1_唯一患者' in target_excel_file.sheet_names:
                target_df = target_excel_file.parse('LisA1_唯一患者')
                # 获取目标表单的表头（第一行）
                target_header = target_df.columns.tolist()
            else:
                # 如果目标文件存在但没有LisA1_唯一患者表单，创建一个空的DataFrame
                target_header = create_default_header()
                target_df = pd.DataFrame(columns=target_header)
        else:
            # 如果目标文件不存在，创建一个空的DataFrame
            target_header = create_default_header()
            target_df = pd.DataFrame(columns=target_header)
        
        # 合并门急诊和住院信息
        combined_df = pd.concat([outpatient_df, inpatient_df], ignore_index=True)
        
        # 提取唯一患者信息
        unique_patients = extract_unique_patients(combined_df, target_header)
        
        # 如果有数据，直接创建DataFrame，不与空DataFrame拼接
        if unique_patients:
            merged_df = pd.DataFrame(unique_patients)
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
        save_to_excel(merged_df, target_file, 'LisA1_唯一患者', index=False, header=True)
        
        print("LisA1_唯一患者表处理完成！")
        return True
        
    except Exception as e:
        print(f"处理LisA1_唯一患者表时发生错误: {e}")
        return False

def extract_unique_patients(combined_df, target_header):
    """
    从合并的数据中提取唯一患者信息
    
    参数:
        combined_df (pandas.DataFrame): 合并的门急诊和住院信息DataFrame
        target_header (list): 目标表头列表
        
    返回:
        list: 唯一患者信息行列表
    """
    unique_patients = []
    
    # 查找患者ID列
    patient_id_cols = find_column_by_keywords(combined_df, ['患者ID', '病人ID', '唯一标识'])
    if not patient_id_cols:
        return []
    
    # 使用找到的第一个列作为患者ID列
    patient_id_col = patient_id_cols[0]
    
    # 查找其他相关列
    name_cols = find_column_by_keywords(combined_df, ['姓名', '患者姓名', '病人姓名'])
    age_cols = find_column_by_keywords(combined_df, ['年龄'])
    gender_cols = find_column_by_keywords(combined_df, ['性别'])
    birth_cols = find_column_by_keywords(combined_df, ['出生', '生日', '出生日期'])
    contact_cols = find_column_by_keywords(combined_df, ['联系', '电话', '手机', '联系方式'])
    address_cols = find_column_by_keywords(combined_df, ['地址', '住址', '家庭住址'])
    
    # 使用找到的第一个列
    name_col = name_cols[0] if name_cols else None
    age_col = age_cols[0] if age_cols else None
    gender_col = gender_cols[0] if gender_cols else None
    birth_col = birth_cols[0] if birth_cols else None
    contact_col = contact_cols[0] if contact_cols else None
    address_col = address_cols[0] if address_cols else None
    
    # 获取唯一患者ID
    unique_patient_ids = combined_df[patient_id_col].dropna().unique()
    
    # 遍历唯一患者ID
    for patient_id in unique_patient_ids:
        # 获取该患者的所有记录
        patient_records = combined_df[combined_df[patient_id_col] == patient_id]
        
        # 如果有多条记录，选择最新的一条
        if len(patient_records) > 1:
            # 查找日期列
            date_cols = find_column_by_keywords(patient_records, ['日期', '时间', '就诊日期', '入院日期'])
            if date_cols:
                date_col = date_cols[0]
                # 尝试将日期列转换为日期类型
                try:
                    patient_records[date_col] = pd.to_datetime(patient_records[date_col], errors='coerce')
                    # 按日期排序，获取最新记录
                    patient_record = patient_records.sort_values(by=date_col, ascending=False).iloc[0]
                except:
                    # 如果日期转换失败，取第一条记录
                    patient_record = patient_records.iloc[0]
            else:
                # 如果没有日期列，取第一条记录
                patient_record = patient_records.iloc[0]
        else:
            # 如果只有一条记录，直接使用
            patient_record = patient_records.iloc[0]
        
        # 创建新行
        new_row = {}
        
        # 获取患者基本信息
        name = patient_record[name_col] if name_col and pd.notna(patient_record[name_col]) else ''
        age = patient_record[age_col] if age_col and pd.notna(patient_record[age_col]) else None
        gender = patient_record[gender_col] if gender_col and pd.notna(patient_record[gender_col]) else ''
        birth_date = patient_record[birth_col] if birth_col and pd.notna(patient_record[birth_col]) else None
        contact = patient_record[contact_col] if contact_col and pd.notna(patient_record[contact_col]) else ''
        address = patient_record[address_col] if address_col and pd.notna(patient_record[address_col]) else ''
        
        # 填充目标表头中的列
        for col in target_header:
            # 根据关键字匹配列
            if any(keyword in col.lower() for keyword in ['患者id', '病人id', '唯一标识']):
                new_row[col] = patient_id
            elif any(keyword in col.lower() for keyword in ['姓名', '患者姓名', '病人姓名']):
                new_row[col] = name
            elif any(keyword in col.lower() for keyword in ['年龄']):
                new_row[col] = age
            elif any(keyword in col.lower() for keyword in ['性别']):
                new_row[col] = gender
            elif any(keyword in col.lower() for keyword in ['出生', '生日', '出生日期']):
                new_row[col] = birth_date
            elif any(keyword in col.lower() for keyword in ['联系', '电话', '手机', '联系方式']):
                new_row[col] = contact
            elif any(keyword in col.lower() for keyword in ['地址', '住址', '家庭住址']):
                new_row[col] = address
            elif col in patient_record.index:
                new_row[col] = patient_record[col]
            else:
                new_row[col] = None
        
        # 添加到结果列表
        unique_patients.append(new_row)
    
    return unique_patients

def create_default_header():
    """
    创建默认的表头
    
    返回:
        list: 默认表头列表
    """
    # 这里可以根据实际需求定义默认表头
    # 以下是一个示例，实际应用中应根据业务需求调整
    return [
        '患者ID', '姓名', '年龄', '性别', '出生日期', 
        '联系方式', '家庭住址', '首次就诊日期', '最近就诊日期',
        '就诊次数', '数据来源', '数据更新时间'
    ]

if __name__ == "__main__":
    # 如果直接运行此脚本，使用默认路径
    script_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    source_file = os.path.join(script_dir, 'data', 'input', '测试原始数据.xlsx')
    target_file = os.path.join(script_dir, 'data', 'output', '测试合并.xlsx')
    
    process(source_file, target_file)