# -*- coding: utf-8 -*-
"""
处理第四个sheet: Lis04_抗病毒药物

此脚本用于处理原始数据中的抗病毒药物相关信息，
合并到目标Excel的Lis04_抗病毒药物表单中。
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
    处理原始数据中的抗病毒药物相关信息，合并到目标Excel的Lis04_抗病毒药物表单中
    
    参数:
        source_file (str): 源Excel文件路径
        target_file (str): 目标Excel文件路径
    """
    try:
        # 读取源文件中的相关信息
        excel_file = pd.ExcelFile(source_file)
        
        # 检查源文件是否包含必要的表单
        required_sheets = ['药物医嘱信息']
        for sheet in required_sheets:
            if sheet not in excel_file.sheet_names:
                raise ValueError(f"源文件中缺少{sheet}表单")
        
        medication_df = excel_file.parse('药物医嘱信息')
        
        # 清理列名
        medication_df = clean_column_names(medication_df)
        
        # 检查目标文件是否存在
        if os.path.exists(target_file):
            # 如果目标文件存在，读取Lis04_抗病毒药物表单
            target_excel_file = pd.ExcelFile(target_file)
            if 'Lis04_抗病毒药物' in target_excel_file.sheet_names:
                target_df = target_excel_file.parse('Lis04_抗病毒药物')
                # 获取目标表单的表头（第一行）
                target_header = target_df.columns.tolist()
            else:
                # 如果目标文件存在但没有Lis04_抗病毒药物表单，创建一个空的DataFrame
                target_header = create_default_header()
                target_df = pd.DataFrame(columns=target_header)
        else:
            # 如果目标文件不存在，创建一个空的DataFrame
            target_header = create_default_header()
            target_df = pd.DataFrame(columns=target_header)
        
        # 从药物医嘱信息中提取抗病毒药物相关信息
        antiviral_info = extract_antiviral_info(medication_df, target_header)
        
        # 如果有数据，直接创建DataFrame，不与空DataFrame拼接
        if antiviral_info:
            merged_df = pd.DataFrame(antiviral_info)
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
        save_to_excel(merged_df, target_file, 'Lis04_抗病毒药物', index=False, header=True)
        
        print("Lis04_抗病毒药物表处理完成！")
        return True
        
    except Exception as e:
        print(f"处理Lis04_抗病毒药物表时发生错误: {e}")
        return False

def extract_antiviral_info(df, target_header):
    """
    从药物医嘱信息中提取抗病毒药物相关信息
    
    参数:
        df (pandas.DataFrame): 药物医嘱信息DataFrame
        target_header (list): 目标表头列表
        
    返回:
        list: 抗病毒药物信息行列表
    """
    antiviral_info = []
    
    # 定义抗病毒药物关键词
    antiviral_keywords = [
        '利巴韦林', '阿比多尔', '奥司他韦', '扎那米韦', '帕拉米韦', '法匹拉韦',
        '瑞德西韦', '洛匹那韦', '利托那韦', '达芦那韦', '克立芝', '克力芝',
        '抗病毒', '抗病毒药', '抗病毒药物', '抗病毒治疗',
        'Ribavirin', 'Arbidol', 'Oseltamivir', 'Zanamivir', 'Peramivir', 'Favipiravir',
        'Remdesivir', 'Lopinavir', 'Ritonavir', 'Darunavir', 'Kaletra'
    ]
    
    # 查找药物名称列
    drug_name_cols = find_column_by_keywords(df, ['药物名称', '药品名称', '医嘱内容', '药名'])
    if not drug_name_cols:
        return []
    
    # 查找患者ID列
    patient_id_cols = find_column_by_keywords(df, ['患者ID', '病人ID', '就诊ID'])
    if not patient_id_cols:
        return []
    
    # 查找用药日期列
    medication_date_cols = find_column_by_keywords(df, ['用药日期', '医嘱日期', '开始日期'])
    
    # 查找用药剂量列
    dosage_cols = find_column_by_keywords(df, ['剂量', '用量', '单次用量'])
    
    # 查找用药频次列
    frequency_cols = find_column_by_keywords(df, ['频次', '用药频次', '给药频次'])
    
    # 查找用药途径列
    route_cols = find_column_by_keywords(df, ['途径', '给药途径', '用药途径'])
    
    # 使用找到的第一个列
    drug_name_col = drug_name_cols[0]
    patient_id_col = patient_id_cols[0]
    medication_date_col = medication_date_cols[0] if medication_date_cols else None
    dosage_col = dosage_cols[0] if dosage_cols else None
    frequency_col = frequency_cols[0] if frequency_cols else None
    route_col = route_cols[0] if route_cols else None
    
    # 遍历数据行
    for _, row in df.iterrows():
        # 检查药物名称是否包含抗病毒药物关键词
        drug_name = str(row[drug_name_col]) if pd.notna(row[drug_name_col]) else ''
        
        if any(keyword in drug_name for keyword in antiviral_keywords):
            # 创建新行
            new_row = {}
            
            # 填充目标表头中的列
            for col in target_header:
                # 根据关键字匹配列
                if any(keyword in col.lower() for keyword in ['患者', '病人', '就诊']):
                    new_row[col] = row[patient_id_col]
                elif any(keyword in col.lower() for keyword in ['药物名称', '药品名称', '医嘱内容', '药名']):
                    new_row[col] = drug_name
                elif any(keyword in col.lower() for keyword in ['用药日期', '医嘱日期', '开始日期']) and medication_date_col:
                    new_row[col] = row[medication_date_col]
                elif any(keyword in col.lower() for keyword in ['剂量', '用量', '单次用量']) and dosage_col:
                    new_row[col] = row[dosage_col]
                elif any(keyword in col.lower() for keyword in ['频次', '用药频次', '给药频次']) and frequency_col:
                    new_row[col] = row[frequency_col]
                elif any(keyword in col.lower() for keyword in ['途径', '给药途径', '用药途径']) and route_col:
                    new_row[col] = row[route_col]
                elif any(keyword in col.lower() for keyword in ['药物类型', '药品类型', '类型']):
                    new_row[col] = '抗病毒药物'
                elif col in df.columns:
                    new_row[col] = row[col]
                else:
                    new_row[col] = None
            
            # 添加到结果列表
            antiviral_info.append(new_row)
    
    return antiviral_info

def create_default_header():
    """
    创建默认的表头
    
    返回:
        list: 默认表头列表
    """
    # 这里可以根据实际需求定义默认表头
    # 以下是一个示例，实际应用中应根据业务需求调整
    return [
        '患者ID', '姓名', '药物名称', '药物类型', '用药日期', 
        '剂量', '单位', '频次', '用药途径', '用药天数',
        '医嘱医生', '医嘱科室', '数据来源', '数据更新时间'
    ]

if __name__ == "__main__":
    # 如果直接运行此脚本，使用默认路径
    script_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    source_file = os.path.join(script_dir, 'data', 'input', '测试原始数据.xlsx')
    target_file = os.path.join(script_dir, 'data', 'output', '测试合并.xlsx')
    
    process(source_file, target_file)