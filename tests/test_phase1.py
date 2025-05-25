# -*- coding: utf-8 -*-
"""
测试第一阶段处理脚本
"""

import os
import sys
import unittest
import pandas as pd
import shutil
from unittest.mock import patch, MagicMock

# 添加项目根目录到系统路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 导入要测试的模块
from scripts.phase1.process_sheet1_attendance import process as process_sheet1
from scripts.phase1.process_sheet2_diagnosis import process as process_sheet2
from scripts.phase1.process_sheet3_covid import process as process_sheet3
from scripts.phase1.process_sheet4_antiviral import process as process_sheet4
from scripts.phase1.process_sheet5_covid_test import process as process_sheet5
from scripts.phase1.process_sheet6_population import process as process_sheet6
from scripts.phase1.process_sheet7_unique_patients import process as process_sheet7


class TestPhase1(unittest.TestCase):
    """
    测试第一阶段处理脚本
    """
    
    def setUp(self):
        """
        测试前的准备工作
        """
        # 创建测试目录
        self.test_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_data')
        self.input_dir = os.path.join(self.test_dir, 'input')
        self.output_dir = os.path.join(self.test_dir, 'output')
        
        os.makedirs(self.input_dir, exist_ok=True)
        os.makedirs(self.output_dir, exist_ok=True)
        
        # 创建测试源文件路径
        # 修改测试源文件名
        self.source_file = os.path.join(self.input_dir, '原始数据.xlsx')
        
        self.target_file = os.path.join(self.output_dir, '测试合并.xlsx')
        
        # 创建测试数据
        self.create_test_data()
    
    def tearDown(self):
        """
        测试后的清理工作
        """
        # 删除测试目录
        if os.path.exists(self.test_dir):
            shutil.rmtree(self.test_dir)
    
    def create_test_data(self):
        """
        创建测试数据
        """
        # 创建门急诊信息表
        outpatient_df = pd.DataFrame({
            '患者ID': ['P001', 'P002', 'P003'],
            '姓名': ['张三', '李四', '王五'],
            '年龄': [30, 40, 50],
            '性别': ['男', '女', '男'],
            '就诊日期': ['2023-01-01', '2023-01-02', '2023-01-03'],
            '就诊科室': ['内科', '外科', '内科'],
            '诊断': ['感冒', '骨折', '高血压']
        })
        
        # 创建住院信息表
        inpatient_df = pd.DataFrame({
            '患者ID': ['P004', 'P005', 'P006'],
            '姓名': ['赵六', '钱七', '孙八'],
            '年龄': [60, 70, 80],
            '性别': ['女', '男', '女'],
            '入院日期': ['2023-01-04', '2023-01-05', '2023-01-06'],
            '出院日期': ['2023-01-10', '2023-01-15', '2023-01-20'],
            '住院科室': ['内科', '外科', '内科'],
            '诊断': ['肺炎', '骨折', '心脏病']
        })
        
        # 创建药物医嘱信息表
        medication_df = pd.DataFrame({
            '患者ID': ['P001', 'P002', 'P003', 'P004', 'P005', 'P006'],
            '药物名称': ['阿莫西林', '布洛芬', '降压药', '抗生素', '止痛药', '强心剂'],
            '用法用量': ['一日三次', '一日两次', '一日一次', '一日三次', '一日两次', '一日一次'],
            '开始日期': ['2023-01-01', '2023-01-02', '2023-01-03', '2023-01-04', '2023-01-05', '2023-01-06'],
            '结束日期': ['2023-01-07', '2023-01-08', '2023-01-09', '2023-01-10', '2023-01-15', '2023-01-20']
        })
        
        # 创建吸氧信息表
        oxygen_df = pd.DataFrame({
            '患者ID': ['P004', 'P005', 'P006'],
            '吸氧方式': ['鼻导管', '面罩', '鼻导管'],
            '吸氧流量': [2, 5, 3],
            '开始日期': ['2023-01-04', '2023-01-05', '2023-01-06'],
            '结束日期': ['2023-01-10', '2023-01-15', '2023-01-20']
        })
        
        # 创建检查信息表
        examination_df = pd.DataFrame({
            '患者ID': ['P001', 'P002', 'P003', 'P004', 'P005', 'P006'],
            '检查项目': ['血常规', 'X光', '心电图', '肺CT', 'X光', '心脏彩超'],
            '检查结果': ['正常', '骨折', '异常', '肺炎', '骨折', '心功能不全'],
            '检查日期': ['2023-01-01', '2023-01-02', '2023-01-03', '2023-01-04', '2023-01-05', '2023-01-06']
        })
        
        # 创建统计数据表
        statistics_df = pd.DataFrame({
            '统计项目': ['总患者数', '门诊患者数', '住院患者数', '男性患者数', '女性患者数', '平均年龄'],
            '统计值': [6, 3, 3, 3, 3, 55]
        })
        
        # 将数据保存到Excel文件
        with pd.ExcelWriter(self.source_file) as writer:
            outpatient_df.to_excel(writer, sheet_name='门急诊信息', index=False)
            inpatient_df.to_excel(writer, sheet_name='住院信息', index=False)
            medication_df.to_excel(writer, sheet_name='药物医嘱信息', index=False)
            oxygen_df.to_excel(writer, sheet_name='吸氧信息', index=False)
            examination_df.to_excel(writer, sheet_name='检查信息', index=False)
            statistics_df.to_excel(writer, sheet_name='统计数据', index=False)
    
    @patch('scripts.phase1.process_sheet1_attendance.optimize_time_format')
    @patch('scripts.phase1.process_sheet1_attendance.save_to_excel')
    def test_process_sheet1(self, mock_save_to_excel, mock_optimize_time_format):
        """
        测试process_sheet1函数
        """
        # 设置模拟函数的行为
        mock_optimize_time_format.side_effect = lambda df: df
        mock_save_to_excel.return_value = None
        
        # 调用函数
        result = process_sheet1(self.source_file, self.target_file)
        
        # 验证结果
        self.assertTrue(result)
        mock_optimize_time_format.assert_called()
        mock_save_to_excel.assert_called_once()
    
    @patch('scripts.phase1.process_sheet2_diagnosis.optimize_time_format')
    @patch('scripts.phase1.process_sheet2_diagnosis.save_to_excel')
    def test_process_sheet2(self, mock_save_to_excel, mock_optimize_time_format):
        """
        测试process_sheet2函数
        """
        # 设置模拟函数的行为
        mock_optimize_time_format.side_effect = lambda df: df
        mock_save_to_excel.return_value = None
        
        # 调用函数
        result = process_sheet2(self.source_file, self.target_file)
        
        # 验证结果
        self.assertTrue(result)
        mock_optimize_time_format.assert_called()
        mock_save_to_excel.assert_called_once()
    
    @patch('scripts.phase1.process_sheet3_covid.optimize_time_format')
    @patch('scripts.phase1.process_sheet3_covid.save_to_excel')
    def test_process_sheet3(self, mock_save_to_excel, mock_optimize_time_format):
        """
        测试process_sheet3函数
        """
        # 设置模拟函数的行为
        mock_optimize_time_format.side_effect = lambda df: df
        mock_save_to_excel.return_value = None
        
        # 调用函数
        result = process_sheet3(self.source_file, self.target_file)
        
        # 验证结果
        self.assertTrue(result)
        mock_optimize_time_format.assert_called()
        mock_save_to_excel.assert_called_once()
    
    @patch('scripts.phase1.process_sheet4_antiviral.optimize_time_format')
    @patch('scripts.phase1.process_sheet4_antiviral.save_to_excel')
    def test_process_sheet4(self, mock_save_to_excel, mock_optimize_time_format):
        """
        测试process_sheet4函数
        """
        # 设置模拟函数的行为
        mock_optimize_time_format.side_effect = lambda df: df
        mock_save_to_excel.return_value = None
        
        # 调用函数
        result = process_sheet4(self.source_file, self.target_file)
        
        # 验证结果
        self.assertTrue(result)
        mock_optimize_time_format.assert_called()
        mock_save_to_excel.assert_called_once()
    
    @patch('scripts.phase1.process_sheet5_covid_test.optimize_time_format')
    @patch('scripts.phase1.process_sheet5_covid_test.save_to_excel')
    def test_process_sheet5(self, mock_save_to_excel, mock_optimize_time_format):
        """
        测试process_sheet5函数
        """
        # 设置模拟函数的行为
        mock_optimize_time_format.side_effect = lambda df: df
        mock_save_to_excel.return_value = None
        
        # 调用函数
        result = process_sheet5(self.source_file, self.target_file)
        
        # 验证结果
        self.assertTrue(result)
        mock_optimize_time_format.assert_called()
        mock_save_to_excel.assert_called_once()
    
    @patch('scripts.phase1.process_sheet6_population.optimize_time_format')
    @patch('scripts.phase1.process_sheet6_population.save_to_excel')
    def test_process_sheet6(self, mock_save_to_excel, mock_optimize_time_format):
        """
        测试process_sheet6函数
        """
        # 设置模拟函数的行为
        mock_optimize_time_format.side_effect = lambda df: df
        mock_save_to_excel.return_value = None
        
        # 调用函数
        result = process_sheet6(self.source_file, self.target_file)
        
        # 验证结果
        self.assertTrue(result)
        mock_optimize_time_format.assert_called()
        mock_save_to_excel.assert_called_once()
    
    @patch('scripts.phase1.process_sheet7_unique_patients.optimize_time_format')
    @patch('scripts.phase1.process_sheet7_unique_patients.save_to_excel')
    def test_process_sheet7(self, mock_save_to_excel, mock_optimize_time_format):
        """
        测试process_sheet7函数
        """
        # 设置模拟函数的行为
        mock_optimize_time_format.side_effect = lambda df: df
        mock_save_to_excel.return_value = None
        
        # 调用函数
        result = process_sheet7(self.source_file, self.target_file)
        
        # 验证结果
        self.assertTrue(result)
        mock_optimize_time_format.assert_called()
        mock_save_to_excel.assert_called_once()


if __name__ == '__main__':
    unittest.main()