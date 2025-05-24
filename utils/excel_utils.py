# -*- coding: utf-8 -*-
"""
Excel处理工具函数

此模块包含处理Excel文件的通用函数，可被各个阶段的脚本复用。
"""

import pandas as pd
import os

def optimize_time_format(df):
    """
    优化DataFrame中的时间格式
    
    将包含时间关键字的列转换为标准日期时间格式，对于时间为00:00:00的数据，只保留日期部分
    
    参数:
        df (pandas.DataFrame): 需要处理的DataFrame
        
    返回:
        pandas.DataFrame: 处理后的DataFrame
    """
    # 定义需要处理的列名关键字
    time_column_keywords = ['就诊日期', '开始日期', '门诊日期', '就诊结束日期', '结束日期', 
                          '出院日期', '入院日期', '住院开始日期', '住院结束日期']
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

def get_excel_files(directory, pattern=None):
    """
    获取指定目录下的所有Excel文件
    
    参数:
        directory (str): 目录路径
        pattern (str, optional): 文件名匹配模式，默认为None，表示获取所有Excel文件
        
    返回:
        list: Excel文件路径列表
    """
    excel_files = []
    for file in os.listdir(directory):
        if file.endswith(('.xlsx', '.xls')):
            if pattern is None or pattern in file:
                excel_files.append(os.path.join(directory, file))
    return excel_files

def save_to_excel(df, file_path, sheet_name, index=False, header=True):
    """
    将DataFrame保存到Excel文件
    
    参数:
        df (pandas.DataFrame): 需要保存的DataFrame
        file_path (str): 保存路径
        sheet_name (str): 表单名称
        index (bool, optional): 是否保存索引，默认为False
        header (bool, optional): 是否保存表头，默认为True
    
    返回:
        bool: 保存成功返回True，失败返回False
    """
    try:
        # 确保目标目录存在
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        
        # 如果文件存在，尝试检查文件是否可写
        if os.path.exists(file_path):
            try:
                # 尝试以写入模式打开文件
                with open(file_path, 'ab') as f:
                    pass
            except PermissionError:
                print(f"错误：文件 {file_path} 正在被其他程序使用，请关闭后重试")
                return False
            
            # 如果文件可写，使用追加模式
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=index, header=header)
        else:
            # 如果文件不存在，创建新文件
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=index, header=header)
        
        print(f"数据已保存到 {file_path} 的 {sheet_name} 表单")
        return True
        
    except PermissionError as e:
        print(f"错误：没有权限写入文件 {file_path}，请检查文件权限")
        return False
    except Exception as e:
        print(f"保存数据时发生错误: {e}")
        return False