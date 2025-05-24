# -*- coding: utf-8 -*-
"""
数据处理工具函数

此模块包含数据处理的通用函数，可被各个阶段的脚本复用。
"""

import pandas as pd
import numpy as np
import re

def clean_column_names(df):
    """
    清理DataFrame的列名
    
    去除列名中的空格、特殊字符，并统一格式
    
    参数:
        df (pandas.DataFrame): 需要处理的DataFrame
        
    返回:
        pandas.DataFrame: 处理后的DataFrame，列名已清理
    """
    # 复制DataFrame，避免修改原始数据
    df_cleaned = df.copy()
    
    # 清理列名
    cleaned_columns = {}
    for col in df_cleaned.columns:
        # 去除前后空格
        cleaned_col = col.strip()
        # 替换中间的多个空格为单个空格
        cleaned_col = re.sub(r'\s+', ' ', cleaned_col)
        # 替换特殊字符
        cleaned_col = re.sub(r'[^\w\s\u4e00-\u9fff]', '', cleaned_col)
        # 存储映射关系
        cleaned_columns[col] = cleaned_col
    
    # 重命名列
    df_cleaned.rename(columns=cleaned_columns, inplace=True)
    
    return df_cleaned

def fill_missing_values(df, strategy='default'):
    """
    填充DataFrame中的缺失值
    
    参数:
        df (pandas.DataFrame): 需要处理的DataFrame
        strategy (str, optional): 填充策略，可选值为'default'、'mean'、'median'、'mode'，默认为'default'
            - default: 数值型列用0填充，字符串列用空字符串填充
            - mean: 数值型列用均值填充
            - median: 数值型列用中位数填充
            - mode: 用众数填充
        
    返回:
        pandas.DataFrame: 处理后的DataFrame，缺失值已填充
    """
    # 复制DataFrame，避免修改原始数据
    df_filled = df.copy()
    
    # 根据不同的策略填充缺失值
    if strategy == 'default':
        # 对于数值型列，用0填充
        numeric_cols = df_filled.select_dtypes(include=['number']).columns
        df_filled[numeric_cols] = df_filled[numeric_cols].fillna(0)
        
        # 对于字符串列，用空字符串填充
        string_cols = df_filled.select_dtypes(include=['object']).columns
        df_filled[string_cols] = df_filled[string_cols].fillna('')
    
    elif strategy == 'mean':
        # 对于数值型列，用均值填充
        numeric_cols = df_filled.select_dtypes(include=['number']).columns
        for col in numeric_cols:
            df_filled[col] = df_filled[col].fillna(df_filled[col].mean())
        
        # 对于字符串列，用空字符串填充
        string_cols = df_filled.select_dtypes(include=['object']).columns
        df_filled[string_cols] = df_filled[string_cols].fillna('')
    
    elif strategy == 'median':
        # 对于数值型列，用中位数填充
        numeric_cols = df_filled.select_dtypes(include=['number']).columns
        for col in numeric_cols:
            df_filled[col] = df_filled[col].fillna(df_filled[col].median())
        
        # 对于字符串列，用空字符串填充
        string_cols = df_filled.select_dtypes(include=['object']).columns
        df_filled[string_cols] = df_filled[string_cols].fillna('')
    
    elif strategy == 'mode':
        # 用众数填充
        for col in df_filled.columns:
            mode_value = df_filled[col].mode()[0] if not df_filled[col].mode().empty else ''
            df_filled[col] = df_filled[col].fillna(mode_value)
    
    return df_filled

def find_column_by_keywords(df, keywords):
    """
    根据关键字查找DataFrame中的列
    
    参数:
        df (pandas.DataFrame): 需要处理的DataFrame
        keywords (list): 关键字列表
        
    返回:
        list: 匹配的列名列表
    """
    matched_columns = []
    for col in df.columns:
        if any(keyword in col for keyword in keywords):
            matched_columns.append(col)
    return matched_columns