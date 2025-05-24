# -*- coding: utf-8 -*-
"""
第三阶段：合并处理多个Excel文件的结果

此脚本用于合并第二阶段处理后的多个Excel文件，
将所有文件的相同sheet合并到一个最终的Excel文件中。
"""

import os
import sys
import glob
import pandas as pd
import logging
from datetime import datetime

# 添加项目根目录到系统路径，以便导入其他模块
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

# 导入工具函数
from utils.excel_utils import get_excel_files, save_to_excel
from utils.data_utils import clean_column_names

def setup_logging():
    """
    设置日志记录
    """
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 'merge_results.log')),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)

def get_sheet_names():
    """
    获取所有需要处理的sheet名称
    
    返回:
        list: sheet名称列表
    """
    return [
        'Lis01_就诊合并',
        'Lis02_诊断',
        'Lis03_新冠感染',
        'Lis04_抗病毒药物',
        'Lis05_新冠检测',
        'Lis06_人群划分',
        'LisA1_唯一患者'
    ]

def merge_sheet_data(input_files, sheet_name):
    """
    合并多个Excel文件中的同名sheet数据
    
    参数:
        input_files (list): 输入文件路径列表
        sheet_name (str): 要合并的sheet名称
        
    返回:
        pandas.DataFrame: 合并后的DataFrame，如果没有数据则返回None
    """
    logger = logging.getLogger(__name__)
    
    # 存储所有文件的sheet数据
    all_data = []
    
    # 遍历所有输入文件
    for file in input_files:
        try:
            # 检查文件是否存在
            if not os.path.exists(file):
                logger.warning(f"文件 {file} 不存在，跳过")
                continue
            
            # 检查文件是否包含指定sheet
            excel_file = pd.ExcelFile(file)
            if sheet_name not in excel_file.sheet_names:
                logger.warning(f"文件 {file} 中不包含sheet {sheet_name}，跳过")
                continue
            
            # 读取sheet数据
            df = excel_file.parse(sheet_name)
            
            # 如果DataFrame为空，跳过
            if df.empty:
                logger.warning(f"文件 {file} 中的sheet {sheet_name} 为空，跳过")
                continue
            
            # 清理列名
            df = clean_column_names(df)
            
            # 添加文件来源列
            df['数据来源'] = os.path.basename(file)
            
            # 添加数据更新时间列
            df['数据更新时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            # 将数据添加到列表中
            all_data.append(df)
            
            logger.info(f"成功读取文件 {file} 中的sheet {sheet_name}，共 {len(df)} 行数据")
            
        except Exception as e:
            logger.error(f"处理文件 {file} 中的sheet {sheet_name} 时发生错误: {e}")
    
    # 如果没有数据，返回None
    if not all_data:
        logger.warning(f"没有找到包含sheet {sheet_name} 的有效数据")
        return None
    
    # 合并所有数据
    merged_df = pd.concat(all_data, ignore_index=True)
    
    # 去除重复行
    if not merged_df.empty:
        # 尝试找到唯一标识列
        id_columns = [col for col in merged_df.columns if any(keyword in col.lower() for keyword in ['id', '编号', '标识'])]
        
        if id_columns:
            # 使用找到的第一个ID列作为去重依据
            id_column = id_columns[0]
            merged_df.drop_duplicates(subset=[id_column], keep='first', inplace=True)
            logger.info(f"使用列 {id_column} 去除重复行，剩余 {len(merged_df)} 行数据")
        else:
            # 如果没有找到ID列，使用所有列去重
            merged_df.drop_duplicates(keep='first', inplace=True)
            logger.info(f"使用所有列去除重复行，剩余 {len(merged_df)} 行数据")
    
    return merged_df

def merge_results(input_dir, output_file, file_pattern="*_processed.xlsx"):
    """
    合并处理多个Excel文件的结果
    
    参数:
        input_dir (str): 输入目录路径
        output_file (str): 输出文件路径
        file_pattern (str, optional): 文件匹配模式，默认为"*_processed.xlsx"
        
    返回:
        dict: 处理结果统计
    """
    logger = logging.getLogger(__name__)
    
    # 确保输出目录存在
    output_dir = os.path.dirname(output_file)
    os.makedirs(output_dir, exist_ok=True)
    
    # 获取所有Excel文件
    input_files = get_excel_files(input_dir, file_pattern)
    
    if not input_files:
        logger.warning(f"在目录 {input_dir} 中未找到匹配 {file_pattern} 的Excel文件")
        return {
            "status": "error",
            "message": f"在目录 {input_dir} 中未找到匹配 {file_pattern} 的Excel文件",
            "sheets_processed": 0,
            "sheets_failed": 0,
            "details": []
        }
    
    # 获取所有需要处理的sheet名称
    sheet_names = get_sheet_names()
    
    # 统计结果
    results = {
        "status": "success",
        "message": "",
        "sheets_processed": 0,
        "sheets_failed": 0,
        "details": []
    }
    
    # 创建一个Excel写入器
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 处理每个sheet
        for sheet_name in sheet_names:
            try:
                logger.info(f"开始合并sheet {sheet_name}")
                
                # 合并sheet数据
                merged_df = merge_sheet_data(input_files, sheet_name)
                
                if merged_df is not None and not merged_df.empty:
                    # 将合并后的数据写入输出文件
                    merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # 更新统计信息
                    results["sheets_processed"] += 1
                    results["details"].append({
                        "sheet_name": sheet_name,
                        "status": "success",
                        "rows": len(merged_df)
                    })
                    
                    logger.info(f"成功合并sheet {sheet_name}，共 {len(merged_df)} 行数据")
                else:
                    # 创建一个空的DataFrame
                    empty_df = pd.DataFrame()
                    empty_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # 更新统计信息
                    results["sheets_failed"] += 1
                    results["details"].append({
                        "sheet_name": sheet_name,
                        "status": "empty",
                        "rows": 0
                    })
                    
                    logger.warning(f"sheet {sheet_name} 没有有效数据，创建空sheet")
            
            except Exception as e:
                # 更新统计信息
                results["sheets_failed"] += 1
                results["details"].append({
                    "sheet_name": sheet_name,
                    "status": "error",
                    "error": str(e)
                })
                
                logger.error(f"合并sheet {sheet_name} 时发生错误: {e}")
    
    # 更新结果状态
    if results["sheets_failed"] > 0 and results["sheets_processed"] > 0:
        results["status"] = "partial"
        results["message"] = f"部分sheet合并成功，{results['sheets_processed']}个成功，{results['sheets_failed']}个失败"
    elif results["sheets_failed"] > 0 and results["sheets_processed"] == 0:
        results["status"] = "error"
        results["message"] = f"所有sheet合并失败"
    else:
        results["status"] = "success"
        results["message"] = f"所有sheet合并成功，共{results['sheets_processed']}个sheet"
    
    # 生成合并报告
    generate_report(results, input_files, output_file)
    
    return results

def generate_report(results, input_files, output_file):
    """
    生成合并报告
    
    参数:
        results (dict): 处理结果统计
        input_files (list): 输入文件路径列表
        output_file (str): 输出文件路径
    """
    logger = logging.getLogger(__name__)
    
    try:
        # 创建报告文件路径
        report_file = os.path.join(os.path.dirname(output_file), "merge_report.xlsx")
        
        # 创建一个Excel写入器
        with pd.ExcelWriter(report_file, engine='openpyxl') as writer:
            # 创建总体统计表
            summary_data = {
                "指标": ["合并状态", "合并信息", "输入文件数", "成功处理的sheet数", "失败处理的sheet数", "输出文件路径"],
                "数值": [
                    results["status"],
                    results["message"],
                    len(input_files),
                    results["sheets_processed"],
                    results["sheets_failed"],
                    output_file
                ]
            }
            
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name="总体统计", index=False)
            
            # 创建输入文件表
            input_files_data = []
            for i, file in enumerate(input_files):
                input_files_data.append({
                    "序号": i + 1,
                    "文件名": os.path.basename(file),
                    "文件路径": file
                })
            
            input_files_df = pd.DataFrame(input_files_data)
            input_files_df.to_excel(writer, sheet_name="输入文件列表", index=False)
            
            # 创建详细结果表
            details_data = []
            for detail in results["details"]:
                details_data.append({
                    "sheet名称": detail["sheet_name"],
                    "状态": detail["status"],
                    "行数": detail.get("rows", 0),
                    "错误信息": detail.get("error", "")
                })
            
            details_df = pd.DataFrame(details_data)
            details_df.to_excel(writer, sheet_name="详细结果", index=False)
        
        logger.info(f"合并报告已生成: {report_file}")
        
    except Exception as e:
        logger.error(f"生成合并报告时发生错误: {e}")

def main():
    """
    主函数
    """
    # 设置日志记录
    logger = setup_logging()
    
    # 获取项目根目录
    root_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    
    # 设置输入和输出目录
    input_dir = os.path.join(root_dir, "data", "output")
    output_file = os.path.join(root_dir, "data", "final", "merged_results.xlsx")
    
    # 确保输出目录存在
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    
    logger.info("开始合并Excel文件")
    logger.info(f"输入目录: {input_dir}")
    logger.info(f"输出文件: {output_file}")
    
    # 合并文件
    results = merge_results(input_dir, output_file)
    
    # 输出处理结果
    logger.info(f"合并完成，状态: {results['status']}, {results['message']}")
    
    return results

if __name__ == "__main__":
    main()