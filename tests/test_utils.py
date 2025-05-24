# -*- coding: utf-8 -*-
"""
测试utils模块中的函数
"""

import os
import sys
import unittest
import pandas as pd
from datetime import datetime

# 添加项目根目录到系统路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 导入要测试的模块
from utils.excel_utils import optimize_time_format, get_excel_files, save_to_excel
from utils.data_utils import clean_column_names, fill_missing_values, find_column_by_keywords


class TestExcelUtils(unittest.TestCase):
    """
    测试excel_utils模块中的函数
    """
    
    def setUp(self):
        """
        测试前的准备工作
        """
        # 创建测试数据
        self.test_df = pd.DataFrame({
            '日期': ['2023-01-01', '2023/01/02', '2023.01.03', '20230104'],
            '时间': ['12:30:45', '12:30', '12-30-45', None]
        })
        
        # 创建临时目录用于测试
        self.test_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_data')
        os.makedirs(self.test_dir, exist_ok=True)
        
        # 创建测试Excel文件
        self.test_file1 = os.path.join(self.test_dir, 'test1.xlsx')
        self.test_file2 = os.path.join(self.test_dir, 'test2.xlsx')
        self.test_df.to_excel(self.test_file1, index=False)
        self.test_df.to_excel(self.test_file2, index=False)
    
    def tearDown(self):
        """
        测试后的清理工作
        """
        # 删除测试文件
        if os.path.exists(self.test_file1):
            os.remove(self.test_file1)
        if os.path.exists(self.test_file2):
            os.remove(self.test_file2)
        
        # 删除测试目录
        if os.path.exists(self.test_dir):
            os.rmdir(self.test_dir)
    
    def test_optimize_time_format(self):
        """
        测试optimize_time_format函数
        """
        # 调用函数
        result_df = optimize_time_format(self.test_df)
        
        # 验证结果
        self.assertIsInstance(result_df, pd.DataFrame)
        self.assertEqual(len(result_df), len(self.test_df))
        
        # 检查日期列是否被正确格式化
        for date_str in result_df['日期']:
            if pd.notna(date_str):
                # 尝试解析日期，如果成功则通过测试
                try:
                    datetime.strptime(date_str, '%Y-%m-%d')
                    self.assertTrue(True)
                except ValueError:
                    self.fail(f"日期格式化失败: {date_str}")
    
    def test_get_excel_files(self):
        """
        测试get_excel_files函数
        """
        # 调用函数
        excel_files = get_excel_files(self.test_dir, "*.xlsx")
        
        # 验证结果
        self.assertIsInstance(excel_files, list)
        self.assertEqual(len(excel_files), 2)
        self.assertTrue(self.test_file1 in excel_files or os.path.abspath(self.test_file1) in excel_files)
        self.assertTrue(self.test_file2 in excel_files or os.path.abspath(self.test_file2) in excel_files)
    
    def test_save_to_excel(self):
        """
        测试save_to_excel函数
        """
        # 准备测试数据
        test_output_file = os.path.join(self.test_dir, 'test_output.xlsx')
        test_sheet_name = 'TestSheet'
        
        # 调用函数
        save_to_excel(self.test_df, test_output_file, test_sheet_name)
        
        # 验证结果
        self.assertTrue(os.path.exists(test_output_file))
        
        # 读取保存的文件并验证内容
        saved_df = pd.read_excel(test_output_file, sheet_name=test_sheet_name)
        self.assertEqual(len(saved_df), len(self.test_df))
        self.assertEqual(list(saved_df.columns), list(self.test_df.columns))
        
        # 清理
        if os.path.exists(test_output_file):
            os.remove(test_output_file)


class TestDataUtils(unittest.TestCase):
    """
    测试data_utils模块中的函数
    """
    
    def setUp(self):
        """
        测试前的准备工作
        """
        # 创建测试数据
        self.test_df = pd.DataFrame({
            ' 姓名 ': ['张三', '李四', '王五', None],
            '年龄 ': [20, 30, None, 50],
            ' 性别': ['男', None, '女', '男']
        })
    
    def test_clean_column_names(self):
        """
        测试clean_column_names函数
        """
        # 调用函数
        result_df = clean_column_names(self.test_df)
        
        # 验证结果
        self.assertIsInstance(result_df, pd.DataFrame)
        self.assertEqual(len(result_df), len(self.test_df))
        
        # 检查列名是否被正确清理
        expected_columns = ['姓名', '年龄', '性别']
        self.assertEqual(list(result_df.columns), expected_columns)
    
    def test_fill_missing_values(self):
        """
        测试fill_missing_values函数
        """
        # 调用函数
        result_df = fill_missing_values(self.test_df)
        
        # 验证结果
        self.assertIsInstance(result_df, pd.DataFrame)
        self.assertEqual(len(result_df), len(self.test_df))
        
        # 检查缺失值是否被填充
        self.assertFalse(result_df[' 姓名 '].isnull().any())
        self.assertFalse(result_df['年龄 '].isnull().any())
        self.assertFalse(result_df[' 性别'].isnull().any())
    
    def test_find_column_by_keywords(self):
        """
        测试find_column_by_keywords函数
        """
        # 调用函数
        name_cols = find_column_by_keywords(self.test_df, ['姓名'])
        age_cols = find_column_by_keywords(self.test_df, ['年龄'])
        gender_cols = find_column_by_keywords(self.test_df, ['性别'])
        not_exist_cols = find_column_by_keywords(self.test_df, ['地址'])
        
        # 验证结果
        self.assertIsInstance(name_cols, list)
        self.assertEqual(len(name_cols), 1)
        self.assertEqual(name_cols[0], ' 姓名 ')
        
        self.assertIsInstance(age_cols, list)
        self.assertEqual(len(age_cols), 1)
        self.assertEqual(age_cols[0], '年龄 ')
        
        self.assertIsInstance(gender_cols, list)
        self.assertEqual(len(gender_cols), 1)
        self.assertEqual(gender_cols[0], ' 性别')
        
        self.assertIsInstance(not_exist_cols, list)
        self.assertEqual(len(not_exist_cols), 0)


if __name__ == '__main__':
    unittest.main()