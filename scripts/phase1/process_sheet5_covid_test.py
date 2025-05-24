# -*- coding: utf-8 -*-
"""
处理第五个sheet: Lis05_新冠检测

此脚本用于处理原始数据中的新冠检测相关信息，
合并到目标Excel的Lis05_新冠检测表单中。
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
    处理原始数据中的新冠检测相关信息，合并到目标Excel的Lis05_新冠检测表单中
    
    参数:
        source_file (str): 源Excel文件路径
        target_file (str): 目标Excel文件路径
    """
    try:
        # 读取源文件中的相关信息
        excel_file = pd.ExcelFile(source_file)
        
        # 检查源文件是否包含必要的表单
        required_sheets = ['检查信息']
        for sheet in required_sheets:
            if sheet not in excel_file.sheet_names:
                raise ValueError(f"源文件中缺少{sheet}表单")
        
        examination_df = excel_file.parse('检查信息')
        
        # 清理列名
        examination_df = clean_column_names(examination_df)
        
        # 检查目标文件是否存在
        if os.path.exists(target_file):
            # 如果目标文件存在，读取Lis05_新冠检测表单
            target_excel_file = pd.ExcelFile(target_file)
            if 'Lis05_新冠检测' in target_excel_file.sheet_names:
                target_df = target_excel_file.parse('Lis05_新冠检测')
                # 获取目标表单的表头（第一行）
                target_header = target_df.columns.tolist()
            else:
                # 如果目标文件存在但没有Lis05_新冠检测表单，创建一个空的DataFrame
                target_header = create_default_header()
                target_df = pd.DataFrame(columns=target_header)
        else:
            # 如果目标文件不存在，创建一个空的DataFrame
            target_header = create_default_header()
            target_df = pd.DataFrame(columns=target_header)
        
        # 从检查信息中提取新冠检测相关信息
        covid_test_info = extract_covid_test_info(examination_df, target_header)
        
        # 如果有数据，直接创建DataFrame，不与空DataFrame拼接
        if covid_test_info:
            merged_df = pd.DataFrame(covid_test_info)
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
        save_to_excel(merged_df, target_file, 'Lis05_新冠检测', index=False, header=True)
        
        print("Lis05_新冠检测表处理完成！")
        return True
        
    except Exception as e:
        print(f"处理Lis05_新冠检测表时发生错误: {e}")
        return False

def extract_covid_test_info(df, target_header):
    """
    从检查信息中提取新冠检测相关信息
    
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
    
    # 查找检查方法列
    test_method_cols = find_column_by_keywords(df, ['检查方法', '检验方法', '方法'])
    
    # 查找检查部门列
    test_dept_cols = find_column_by_keywords(df, ['检查部门', '检验部门', '部门', '科室'])
    
    # 使用找到的第一个列
    test_name_col = test_name_cols[0]
    test_result_col = test_result_cols[0]
    patient_id_col = patient_id_cols[0]
    test_date_col = test_date_cols[0] if test_date_cols else None
    test_method_col = test_method_cols[0] if test_method_cols else None
    test_dept_col = test_dept_cols[0] if test_dept_cols else None
    
    # 遍历数据行
    for _, row in df.iterrows():
        # 检查名称是否包含新冠相关关键词
        test_name = str(row[test_name_col]) if pd.notna(row[test_name_col]) else ''
        
        if any(keyword in test_name for keyword in covid_test_keywords):
            # 创建新行
            new_row = {}
            
            # 获取检测结果
            test_result = str(row[test_result_col]) if pd.notna(row[test_result_col]) else ''
            
            # 判断检测结果类型
            result_type = '未知'
            if re.search(r'阳性|检出|positive', test_result, re.IGNORECASE):
                result_type = '阳性'
            elif re.search(r'阴性|未检出|negative', test_result, re.IGNORECASE):
                result_type = '阴性'
            
            # 判断检测方法
            test_method = ''
            if test_method_col and pd.notna(row[test_method_col]):
                test_method = row[test_method_col]
            else:
                # 从检查名称推断检测方法
                if re.search(r'核酸|PCR|RT-PCR', test_name, re.IGNORECASE):
                    test_method = '核酸检测'
                elif re.search(r'抗原', test_name, re.IGNORECASE):
                    test_method = '抗原检测'
                elif re.search(r'抗体|IgM|IgG', test_name, re.IGNORECASE):
                    test_method = '抗体检测'
            
            # 填充目标表头中的列
            for col in target_header:
                # 根据关键字匹配列
                if any(keyword in col.lower() for keyword in ['患者', '病人', '就诊']):
                    new_row[col] = row[patient_id_col]
                elif any(keyword in col.lower() for keyword in ['检查名称', '检验名称', '项目名称']):
                    new_row[col] = test_name
                elif any(keyword in col.lower() for keyword in ['检查结果', '检验结果', '结果']):
                    new_row[col] = test_result
                elif any(keyword in col.lower() for keyword in ['结果类型', '结果判断']):
                    new_row[col] = result_type
                elif any(keyword in col.lower() for keyword in ['检查日期', '检验日期', '日期']) and test_date_col:
                    new_row[col] = row[test_date_col]
                elif any(keyword in col.lower() for keyword in ['检查方法', '检验方法', '方法']):
                    new_row[col] = test_method
                elif any(keyword in col.lower() for keyword in ['检查部门', '检验部门', '部门', '科室']) and test_dept_col:
                    new_row[col] = row[test_dept_col]
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
        '患者ID', '姓名', '检查名称', '检查方法', '检查日期', 
        '检查结果', '结果类型', 'CT值', '检查部门', '采样部位',
        '备注', '数据来源', '数据更新时间'
    ]

if __name__ == "__main__":
    # 如果直接运行此脚本，使用默认路径
    script_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    source_file = os.path.join(script_dir, 'data', 'input', '测试原始数据.xlsx')
    target_file = os.path.join(script_dir, 'data', 'output', '测试合并.xlsx')
    
    process(source_file, target_file)