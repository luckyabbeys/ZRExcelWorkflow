# -*- coding: utf-8 -*-
"""
处理第六个sheet: Lis06_人群划分

此脚本用于处理原始数据中的人群划分相关信息，
合并到目标Excel的Lis06_人群划分表单中。
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
    处理原始数据中的人群划分相关信息，合并到目标Excel的Lis06_人群划分表单中
    
    参数:
        source_file (str): 源Excel文件路径
        target_file (str): 目标Excel文件路径
    """
    try:
        # 读取源文件中的相关信息
        excel_file = pd.ExcelFile(source_file)
        
        # 检查源文件是否包含必要的表单
        required_sheets = ['门急诊信息', '住院信息', '统计数据']
        for sheet in required_sheets:
            if sheet not in excel_file.sheet_names:
                raise ValueError(f"源文件中缺少{sheet}表单")
        
        outpatient_df = excel_file.parse('门急诊信息')
        inpatient_df = excel_file.parse('住院信息')
        statistics_df = excel_file.parse('统计数据')
        
        # 清理列名
        outpatient_df = clean_column_names(outpatient_df)
        inpatient_df = clean_column_names(inpatient_df)
        statistics_df = clean_column_names(statistics_df)
        
        # 检查目标文件是否存在
        if os.path.exists(target_file):
            # 如果目标文件存在，读取Lis06_人群划分表单
            target_excel_file = pd.ExcelFile(target_file)
            if 'Lis06_人群划分' in target_excel_file.sheet_names:
                target_df = target_excel_file.parse('Lis06_人群划分')
                # 获取目标表单的表头（第一行）
                target_header = target_df.columns.tolist()
            else:
                # 如果目标文件存在但没有Lis06_人群划分表单，创建一个空的DataFrame
                target_header = create_default_header()
                target_df = pd.DataFrame(columns=target_header)
        else:
            # 如果目标文件不存在，创建一个空的DataFrame
            target_header = create_default_header()
            target_df = pd.DataFrame(columns=target_header)
        
        # 合并门急诊和住院信息，用于人群划分
        combined_df = pd.concat([outpatient_df, inpatient_df], ignore_index=True)
        
        # 从统计数据中提取人群划分信息
        population_info = extract_population_info(combined_df, statistics_df, target_header)
        
        # 如果有数据，直接创建DataFrame，不与空DataFrame拼接
        if population_info:
            merged_df = pd.DataFrame(population_info)
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
        save_to_excel(merged_df, target_file, 'Lis06_人群划分', index=False, header=True)
        
        print("Lis06_人群划分表处理完成！")
        return True
        
    except Exception as e:
        print(f"处理Lis06_人群划分表时发生错误: {e}")
        return False

def extract_population_info(combined_df, statistics_df, target_header):
    """
    从合并的数据和统计数据中提取人群划分信息
    
    参数:
        combined_df (pandas.DataFrame): 合并的门急诊和住院信息DataFrame
        statistics_df (pandas.DataFrame): 统计数据DataFrame
        target_header (list): 目标表头列表
        
    返回:
        list: 人群划分信息行列表
    """
    population_info = []
    
    # 查找患者ID列
    patient_id_cols = find_column_by_keywords(combined_df, ['患者ID', '病人ID', '就诊ID'])
    if not patient_id_cols:
        return []
    
    # 查找年龄列
    age_cols = find_column_by_keywords(combined_df, ['年龄'])
    
    # 查找性别列
    gender_cols = find_column_by_keywords(combined_df, ['性别'])
    
    # 查找地区列
    region_cols = find_column_by_keywords(combined_df, ['地区', '区域', '地址', '住址'])
    
    # 使用找到的第一个列
    patient_id_col = patient_id_cols[0]
    age_col = age_cols[0] if age_cols else None
    gender_col = gender_cols[0] if gender_cols else None
    region_col = region_cols[0] if region_cols else None
    
    # 从统计数据中提取人群分类信息
    population_categories = extract_population_categories(statistics_df)
    
    # 遍历合并数据的每一行
    for _, row in combined_df.iterrows():
        # 创建新行
        new_row = {}
        
        # 获取患者基本信息
        patient_id = row[patient_id_col] if pd.notna(row[patient_id_col]) else ''
        age = row[age_col] if age_col and pd.notna(row[age_col]) else None
        gender = row[gender_col] if gender_col and pd.notna(row[gender_col]) else ''
        region = row[region_col] if region_col and pd.notna(row[region_col]) else ''
        
        # 确定年龄段
        age_group = determine_age_group(age)
        
        # 确定人群类别
        population_category = determine_population_category(patient_id, age, gender, region, population_categories)
        
        # 填充目标表头中的列
        for col in target_header:
            # 根据关键字匹配列
            if any(keyword in col.lower() for keyword in ['患者', '病人', '就诊']):
                new_row[col] = patient_id
            elif any(keyword in col.lower() for keyword in ['年龄']):
                new_row[col] = age
            elif any(keyword in col.lower() for keyword in ['性别']):
                new_row[col] = gender
            elif any(keyword in col.lower() for keyword in ['地区', '区域', '地址', '住址']):
                new_row[col] = region
            elif any(keyword in col.lower() for keyword in ['年龄段', '年龄组', '年龄分组']):
                new_row[col] = age_group
            elif any(keyword in col.lower() for keyword in ['人群类别', '人群分类', '人群划分']):
                new_row[col] = population_category
            elif col in combined_df.columns:
                new_row[col] = row[col]
            else:
                new_row[col] = None
        
        # 添加到结果列表
        population_info.append(new_row)
    
    return population_info

def extract_population_categories(statistics_df):
    """
    从统计数据中提取人群分类信息
    
    参数:
        statistics_df (pandas.DataFrame): 统计数据DataFrame
        
    返回:
        dict: 人群分类信息字典
    """
    # 这里可以根据实际需求从统计数据中提取人群分类信息
    # 以下是一个示例，实际应用中应根据业务需求调整
    population_categories = {
        'elderly': {'name': '老年人', 'criteria': lambda age, gender, region: age >= 65 if age is not None else False},
        'adult': {'name': '成年人', 'criteria': lambda age, gender, region: 18 <= age < 65 if age is not None else False},
        'child': {'name': '儿童', 'criteria': lambda age, gender, region: age < 18 if age is not None else False},
        'male': {'name': '男性', 'criteria': lambda age, gender, region: gender == '男'},
        'female': {'name': '女性', 'criteria': lambda age, gender, region: gender == '女'},
    }
    
    return population_categories

def determine_age_group(age):
    """
    根据年龄确定年龄段
    
    参数:
        age: 年龄值
        
    返回:
        str: 年龄段描述
    """
    if age is None:
        return '未知'
    
    try:
        age_value = float(age)
        if age_value < 3:
            return '婴幼儿(0-2岁)'
        elif age_value < 6:
            return '学龄前(3-5岁)'
        elif age_value < 12:
            return '儿童(6-11岁)'
        elif age_value < 18:
            return '青少年(12-17岁)'
        elif age_value < 35:
            return '青年(18-34岁)'
        elif age_value < 60:
            return '中年(35-59岁)'
        elif age_value < 80:
            return '老年(60-79岁)'
        else:
            return '高龄老人(80岁以上)'
    except (ValueError, TypeError):
        return '未知'

def determine_population_category(patient_id, age, gender, region, population_categories):
    """
    根据患者信息确定人群类别
    
    参数:
        patient_id: 患者ID
        age: 年龄
        gender: 性别
        region: 地区
        population_categories: 人群分类信息字典
        
    返回:
        str: 人群类别描述
    """
    # 根据年龄确定主要人群类别
    if age is not None:
        if age >= 65:
            return '老年人'
        elif 18 <= age < 65:
            return '成年人'
        else:
            return '儿童青少年'
    
    # 如果没有年龄信息，根据性别分类
    if gender:
        if gender == '男':
            return '男性'
        elif gender == '女':
            return '女性'
    
    return '未分类'

def create_default_header():
    """
    创建默认的表头
    
    返回:
        list: 默认表头列表
    """
    # 这里可以根据实际需求定义默认表头
    # 以下是一个示例，实际应用中应根据业务需求调整
    return [
        '患者ID', '姓名', '年龄', '性别', '地区', 
        '年龄段', '人群类别', '特殊人群标记', '备注', '数据来源',
        '数据更新时间'
    ]

if __name__ == "__main__":
    # 如果直接运行此脚本，使用默认路径
    script_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    source_file = os.path.join(script_dir, 'data', 'input', '测试原始数据.xlsx')
    target_file = os.path.join(script_dir, 'data', 'output', '测试合并.xlsx')
    
    process(source_file, target_file)