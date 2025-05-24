# -*- coding: utf-8 -*-
"""
处理第一个sheet: Lis01_就诊合并

此脚本用于处理原始数据中的门急诊信息和住院信息，
合并到目标Excel的Lis01_就诊合并表单中。
"""

import pandas as pd
import os
import sys

# 添加项目根目录到系统路径，以便导入utils模块
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from utils.excel_utils import optimize_time_format, save_to_excel
from utils.data_utils import find_column_by_keywords

def create_default_header():
    """
    创建默认的表头，与目标文件保持一致
    
    返回:
        list: 默认表头列表
    """
    return [
        '医院编码', '医院名称', '患者编号', '患者唯一编码', '来源',
        '年龄（周岁）', '性别', '就诊科室', '就诊类别', '就诊日期',
        '就诊结束日期', '诊断（ICD编码）', '诊断（文字）', '主诉',
        '现病史（临床症状）', '既往新冠阳性次数', '上次感染日期', '疫苗接种次数',
        '末次接种日期', '是否收入院', '入院日期', '出院日期', '是否收入ICU',
        '入ICU日期', '出ICU日期', '是否死亡', '死亡诊断', '死亡日期'
    ]

def process(source_file, target_file):
    """
    处理原始数据中的门急诊信息和住院信息，合并到目标Excel的Lis01_就诊合并表单中
    
    参数:
        source_file (str): 源Excel文件路径
        target_file (str): 目标Excel文件路径
    """
    try:
        # 读取源文件中的门急诊信息和住院信息
        excel_file = pd.ExcelFile(source_file)
        outpatient_df = excel_file.parse('门急诊信息')
        inpatient_df = excel_file.parse('住院信息')
        
        # 获取标准表头
        target_header = create_default_header()
        
        # 检查目标文件是否存在
        if os.path.exists(target_file):
            # 如果目标文件存在，读取Lis01_就诊合并表单
            target_excel_file = pd.ExcelFile(target_file)
            if 'Lis01_就诊合并' in target_excel_file.sheet_names:
                target_df = target_excel_file.parse('Lis01_就诊合并')
                # 如果目标文件的列数与标准表头不匹配，使用标准表头创建新的DataFrame
                if len(target_df.columns) != len(target_header):
                    print(f"警告：目标文件的列数({len(target_df.columns)})与标准表头列数({len(target_header)})不匹配，将使用标准表头")
                    target_df = pd.DataFrame(columns=target_header)
        else:
            # 如果目标文件不存在，创建一个空的DataFrame
            target_df = pd.DataFrame(columns=target_header)
        
        # 定义需要查找的列名关键字
        column_keywords = {
            '来源列': ['来源', '就诊类别', '就诊类型'],
            '就诊日期列': ['就诊日期', '门诊日期'],
            '就诊结束列': ['就诊结束日期', '出院日期'],
            '入院日期列': ['入院日期'],
            '出院日期列': ['出院日期']
        }
        
        # 动态查找列索引
        column_indices = {}
        for col_name, keywords in column_keywords.items():
            for idx, header in enumerate(target_header):
                if any(keyword in header for keyword in keywords):
                    column_indices[col_name] = idx
                    break
            if col_name not in column_indices:
                raise ValueError(f"找不到包含以下关键词的{col_name}: {', '.join(keywords)}")
        
        # 确保目标表头包含所需列
        required_cols = ['就诊日期列', '就诊结束列', '入院日期列', '出院日期列']
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
            new_row[target_header[column_indices['来源列']]] = '门急诊'
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
            new_row[target_header[column_indices['来源列']]] = '住院'
            
            # 迁移住院日期信息
            admission_date = row.get(target_header[column_indices['入院日期列']])
            discharge_date = row.get(target_header[column_indices['出院日期列']])
            
            new_row[target_header[column_indices['就诊日期列']]] = admission_date
            new_row[target_header[column_indices['就诊结束列']]] = discharge_date
            
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
        
        # 将合并后的数据保存到目标文件
        if not save_to_excel(final_df, target_file, 'Lis01_就诊合并', index=False, header=False):
            raise Exception("保存数据到Excel文件失败")
        
        print("Lis01_就诊合并表处理完成！")
        return True
        
    except Exception as e:
        print(f"处理Lis01_就诊合并表时发生错误: {e}")
        return False

def print_header_info(file_path):
    """
    打印Excel文件中Lis01_就诊合并表单的表头信息
    
    参数:
        file_path (str): Excel文件路径
    """
    try:
        # 读取Excel文件
        excel_file = pd.ExcelFile(file_path)
        if 'Lis01_就诊合并' in excel_file.sheet_names:
            df = excel_file.parse('Lis01_就诊合并')
            print("\n表头信息:")
            for idx, col in enumerate(df.columns):
                print(f"{idx + 1}. {col}")
        else:
            print("未找到'Lis01_就诊合并'表单")
    except Exception as e:
        print(f"读取表头信息时发生错误: {e}")

if __name__ == "__main__":
    # 如果直接运行此脚本，使用默认路径
    script_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    source_file = os.path.join(script_dir, 'data', 'input', '测试原始数据.xlsx')
    target_file = os.path.join(script_dir, '测试合并.xlsx')
    
    # 打印目标文件的表头信息
    print("\n目标文件的表头信息:")
    print_header_info(target_file)
    
    # 处理数据
    process(source_file, target_file)