# -*- coding: utf-8 -*-
"""
第二阶段：批量处理多个Excel文件

此脚本用于批量处理指定目录下的所有Excel文件，
将处理结果保存到输出目录中。
"""

import os
import sys
import glob
import pandas as pd
import logging
from concurrent.futures import ProcessPoolExecutor, as_completed

# 添加项目根目录到系统路径，以便导入其他模块
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

# 导入第一阶段的处理脚本
from scripts.phase1.process_sheet1_attendance import process as process_sheet1
from scripts.phase1.process_sheet2_diagnosis import process as process_sheet2
from scripts.phase1.process_sheet3_covid import process as process_sheet3
from scripts.phase1.process_sheet4_antiviral import process as process_sheet4
from scripts.phase1.process_sheet5_covid_test import process as process_sheet5
from scripts.phase1.process_sheet6_population import process as process_sheet6
from scripts.phase1.process_sheet7_unique_patients import process as process_sheet7

# 导入工具函数
from utils.excel_utils import get_excel_files

def setup_logging():
    """
    设置日志记录
    """
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 'batch_process.log')),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)

def process_single_file(source_file, output_dir, sheets_to_process=None):
    """
    处理单个Excel文件
    
    参数:
        source_file (str): 源Excel文件路径
        output_dir (str): 输出目录路径
        sheets_to_process (list, optional): 要处理的sheet列表，如果为None则处理所有sheet
        
    返回:
        tuple: (文件名, 成功处理的sheet列表, 失败处理的sheet列表)
    """
    logger = logging.getLogger(__name__)
    
    # 获取文件名（不含扩展名）
    file_name = os.path.splitext(os.path.basename(source_file))[0]
    
    # 创建输出文件路径
    target_file = os.path.join(output_dir, f"{file_name}_processed.xlsx")
    
    # 记录处理结果
    success_sheets = []
    failed_sheets = []
    
    # 定义处理函数字典
    process_functions = {
        'sheet1': process_sheet1,
        'sheet2': process_sheet2,
        'sheet3': process_sheet3,
        'sheet4': process_sheet4,
        'sheet5': process_sheet5,
        'sheet6': process_sheet6,
        'sheet7': process_sheet7
    }
    
    # 如果未指定要处理的sheet，则处理所有sheet
    if sheets_to_process is None:
        sheets_to_process = list(process_functions.keys())
    
    # 处理每个sheet
    for sheet in sheets_to_process:
        if sheet in process_functions:
            try:
                logger.info(f"开始处理文件 {file_name} 的 {sheet}")
                result = process_functions[sheet](source_file, target_file)
                if result:
                    success_sheets.append(sheet)
                    logger.info(f"文件 {file_name} 的 {sheet} 处理成功")
                else:
                    failed_sheets.append(sheet)
                    logger.error(f"文件 {file_name} 的 {sheet} 处理失败")
            except Exception as e:
                failed_sheets.append(sheet)
                logger.error(f"处理文件 {file_name} 的 {sheet} 时发生错误: {e}")
        else:
            logger.warning(f"未找到处理 {sheet} 的函数，跳过")
    
    return (file_name, success_sheets, failed_sheets)

def batch_process(input_dir, output_dir, file_pattern="*.xlsx", sheets_to_process=None, max_workers=None):
    """
    批量处理指定目录下的所有Excel文件
    
    参数:
        input_dir (str): 输入目录路径
        output_dir (str): 输出目录路径
        file_pattern (str, optional): 文件匹配模式，默认为"*.xlsx"
        sheets_to_process (list, optional): 要处理的sheet列表，如果为None则处理所有sheet
        max_workers (int, optional): 最大工作进程数，如果为None则使用默认值
        
    返回:
        dict: 处理结果统计
    """
    logger = logging.getLogger(__name__)
    
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    # 获取所有Excel文件
    excel_files = get_excel_files(input_dir, file_pattern)
    
    if not excel_files:
        logger.warning(f"在目录 {input_dir} 中未找到匹配 {file_pattern} 的Excel文件")
        return {
            "total_files": 0,
            "processed_files": 0,
            "success_files": 0,
            "failed_files": 0,
            "details": []
        }
    
    # 统计结果
    results = {
        "total_files": len(excel_files),
        "processed_files": 0,
        "success_files": 0,
        "failed_files": 0,
        "details": []
    }
    
    # 使用进程池并行处理文件
    with ProcessPoolExecutor(max_workers=max_workers) as executor:
        # 提交所有任务
        future_to_file = {executor.submit(process_single_file, file, output_dir, sheets_to_process): file for file in excel_files}
        
        # 处理完成的任务
        for future in as_completed(future_to_file):
            file = future_to_file[future]
            try:
                file_name, success_sheets, failed_sheets = future.result()
                results["processed_files"] += 1
                
                # 记录详细结果
                file_result = {
                    "file_name": file_name,
                    "source_file": file,
                    "target_file": os.path.join(output_dir, f"{file_name}_processed.xlsx"),
                    "success_sheets": success_sheets,
                    "failed_sheets": failed_sheets,
                    "status": "success" if failed_sheets == [] else "partial" if success_sheets else "failed"
                }
                
                results["details"].append(file_result)
                
                # 更新统计信息
                if file_result["status"] == "success":
                    results["success_files"] += 1
                elif file_result["status"] == "failed":
                    results["failed_files"] += 1
                
                logger.info(f"文件 {file_name} 处理完成，状态: {file_result['status']}")
                
            except Exception as e:
                results["processed_files"] += 1
                results["failed_files"] += 1
                
                # 记录错误详情
                file_result = {
                    "file_name": os.path.splitext(os.path.basename(file))[0],
                    "source_file": file,
                    "error": str(e),
                    "status": "error"
                }
                
                results["details"].append(file_result)
                
                logger.error(f"处理文件 {file} 时发生错误: {e}")
    
    # 生成处理报告
    generate_report(results, output_dir)
    
    return results

def generate_report(results, output_dir):
    """
    生成处理报告
    
    参数:
        results (dict): 处理结果统计
        output_dir (str): 输出目录路径
    """
    logger = logging.getLogger(__name__)
    
    try:
        # 创建报告文件路径
        report_file = os.path.join(output_dir, "batch_process_report.xlsx")
        
        # 创建一个Excel写入器
        with pd.ExcelWriter(report_file, engine='openpyxl') as writer:
            # 创建总体统计表
            summary_data = {
                "指标": ["总文件数", "处理文件数", "成功文件数", "失败文件数", "成功率"],
                "数值": [
                    results["total_files"],
                    results["processed_files"],
                    results["success_files"],
                    results["failed_files"],
                    f"{results['success_files'] / results['total_files'] * 100:.2f}%" if results["total_files"] > 0 else "0.00%"
                ]
            }
            
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name="总体统计", index=False)
            
            # 创建详细结果表
            details_data = []
            for detail in results["details"]:
                details_data.append({
                    "文件名": detail["file_name"],
                    "源文件路径": detail["source_file"],
                    "目标文件路径": detail.get("target_file", ""),
                    "状态": detail["status"],
                    "成功处理的sheet": ", ".join(detail.get("success_sheets", [])),
                    "失败处理的sheet": ", ".join(detail.get("failed_sheets", [])),
                    "错误信息": detail.get("error", "")
                })
            
            details_df = pd.DataFrame(details_data)
            details_df.to_excel(writer, sheet_name="详细结果", index=False)
        
        logger.info(f"处理报告已生成: {report_file}")
        
    except Exception as e:
        logger.error(f"生成处理报告时发生错误: {e}")

def main():
    """
    主函数
    """
    # 设置日志记录
    logger = setup_logging()
    
    # 获取项目根目录
    root_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    
    # 设置输入和输出目录
    input_dir = os.path.join(root_dir, "data", "input")
    output_dir = os.path.join(root_dir, "data", "output")
    
    logger.info("开始批量处理Excel文件")
    logger.info(f"输入目录: {input_dir}")
    logger.info(f"输出目录: {output_dir}")
    
    # 批量处理文件
    results = batch_process(input_dir, output_dir)
    
    # 输出处理结果
    logger.info(f"批量处理完成，总文件数: {results['total_files']}, 成功: {results['success_files']}, 失败: {results['failed_files']}")
    
    return results

if __name__ == "__main__":
    main()