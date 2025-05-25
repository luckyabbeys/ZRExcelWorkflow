# -*- coding: utf-8 -*-
"""
Excel自动化处理工作流主程序

此脚本作为整个工作流的入口，可以按顺序调用各个阶段的脚本，
也可以通过命令行参数指定只运行特定阶段或特定sheet的处理。
"""

import os
import sys
import argparse
import importlib
import logging
from datetime import datetime

# 设置日志
def setup_logging():
    """
    设置日志配置
    """
    log_dir = 'logs'
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    log_file = os.path.join(log_dir, f'workflow_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log')
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    
    return logging.getLogger('workflow')

# 解析命令行参数
def parse_arguments():
    """
    解析命令行参数
    
    返回:
        argparse.Namespace: 解析后的参数
    """
    parser = argparse.ArgumentParser(description='Excel自动化处理工作流')
    
    parser.add_argument(
        '--phase',
        type=int,
        choices=[1, 2, 3],
        help='指定要运行的阶段 (1, 2, 或 3)'
    )
    
    parser.add_argument(
        '--sheet',
        type=int,
        choices=range(1, 8),  # 1-7对应7个sheet
        help='指定要处理的sheet (1-7)'
    )
    
    parser.add_argument(
        '--input',
        type=str,
        help='指定输入文件或目录路径'
    )
    
    parser.add_argument(
        '--output',
        type=str,
        help='指定输出文件或目录路径'
    )
    
    return parser.parse_args()

# 运行第一阶段处理
def run_phase1(logger, sheet=None, input_file=None, output_file=None):
    """
    运行第一阶段处理
    
    参数:
        logger (logging.Logger): 日志记录器
        sheet (int, optional): 指定要处理的sheet编号，默认为None表示处理所有sheet
        input_file (str, optional): 输入文件路径，默认为None使用默认路径
        output_file (str, optional): 输出文件路径，默认为None使用默认路径
    """
    logger.info('开始运行第一阶段处理')
    
    # 设置默认文件路径
    if input_file is None:
        input_file = os.path.join('data', 'input', '原始数据.xlsx')
    
    if output_file is None:
        # 获取输入文件名，去掉扩展名
        input_file_name = os.path.splitext(os.path.basename(input_file))[0]
        # 生成输出文件名
        output_file = os.path.join('data', 'output', f'{input_file_name}合并phase1sheet1.xlsx')
    
    # 确保输入文件存在
    if not os.path.exists(input_file):
        logger.error(f'输入文件不存在: {input_file}')
        return False
    
    # 确保输出目录存在
    output_dir = os.path.dirname(output_file)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # 处理指定的sheet或所有sheet
    sheet_modules = {
        1: 'scripts.phase1.process_sheet1_attendance',
        2: 'scripts.phase1.process_sheet2_diagnosis',
        3: 'scripts.phase1.process_sheet3_covid',
        4: 'scripts.phase1.process_sheet4_antiviral',
        5: 'scripts.phase1.process_sheet5_covid_test',
        6: 'scripts.phase1.process_sheet6_population',
        7: 'scripts.phase1.process_sheet7_unique_patients'
    }
    
    if sheet is not None:
        # 处理指定的sheet
        if sheet not in sheet_modules:
            logger.error(f'无效的sheet编号: {sheet}')
            return False
        
        try:
            module = importlib.import_module(sheet_modules[sheet])
            logger.info(f'运行模块: {sheet_modules[sheet]}')
            module.process(input_file, output_file)
            logger.info(f'Sheet {sheet} 处理完成')
        except Exception as e:
            logger.error(f'处理Sheet {sheet}时发生错误: {e}')
            return False
    else:
        # 处理所有sheet
        success = True
        for sheet_num, module_name in sheet_modules.items():
            try:
                module = importlib.import_module(module_name)
                logger.info(f'运行模块: {module_name}')
                module.process(input_file, output_file)
                logger.info(f'Sheet {sheet_num} 处理完成')
            except Exception as e:
                logger.error(f'处理Sheet {sheet_num}时发生错误: {e}')
                success = False
        
        if not success:
            return False
    
    logger.info('第一阶段处理完成')
    return True

# 运行第二阶段处理
def run_phase2(logger, input_dir=None, output_dir=None):
    """
    运行第二阶段处理
    
    参数:
        logger (logging.Logger): 日志记录器
        input_dir (str, optional): 输入目录路径，默认为None使用默认路径
        output_dir (str, optional): 输出目录路径，默认为None使用默认路径
    """
    logger.info('开始运行第二阶段处理')
    
    # 设置默认目录路径
    if input_dir is None:
        input_dir = os.path.join('data', 'input')
    
    if output_dir is None:
        output_dir = os.path.join('data', 'output')
    
    # 确保目录存在
    if not os.path.exists(input_dir):
        logger.error(f'输入目录不存在: {input_dir}')
        return False
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    try:
        # 导入并运行第二阶段处理模块
        module = importlib.import_module('scripts.phase2.batch_process')
        logger.info('运行批量处理模块')
        module.batch_process(input_dir, output_dir)
        logger.info('第二阶段处理完成')
        return True
    except Exception as e:
        logger.error(f'第二阶段处理时发生错误: {e}')
        return False

# 运行第三阶段处理
def run_phase3(logger, input_dir=None, output_file=None):
    """
    运行第三阶段处理
    
    参数:
        logger (logging.Logger): 日志记录器
        input_dir (str, optional): 输入目录路径，默认为None使用默认路径
        output_file (str, optional): 输出文件路径，默认为None使用默认路径
    """
    logger.info('开始运行第三阶段处理')
    
    # 设置默认路径
    if input_dir is None:
        input_dir = os.path.join('data', 'output')
    
    if output_file is None:
        # 获取输入文件名，去掉扩展名
        input_file_name = os.path.splitext(os.path.basename(input_file))[0]
        # 生成输出文件名
        output_file = os.path.join('data', 'output', f'{input_file_name}合并phase3.xlsx')
    
    # 确保目录存在
    if not os.path.exists(input_dir):
        logger.error(f'输入目录不存在: {input_dir}')
        return False
    
    output_dir = os.path.dirname(output_file)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    try:
        # 导入并运行第三阶段处理模块
        module = importlib.import_module('scripts.phase3.merge_results')
        logger.info('运行文件合并模块')
        module.merge_results(input_dir, output_file)
        logger.info('第三阶段处理完成')
        return True
    except Exception as e:
        logger.error(f'第三阶段处理时发生错误: {e}')
        return False

# 主函数
def main():
    """
    主函数，程序入口
    """
    # 设置日志
    logger = setup_logging()
    logger.info('Excel自动化处理工作流开始运行')
    
    # 解析命令行参数
    args = parse_arguments()
    
    # 根据参数运行相应的阶段
    if args.phase == 1 or args.phase is None:
        success = run_phase1(logger, args.sheet, args.input, args.output)
        if not success:
            logger.error('第一阶段处理失败')
            sys.exit(1)
    
    if args.phase == 2 or args.phase is None:
        success = run_phase2(logger, args.input, args.output)
        if not success:
            logger.error('第二阶段处理失败')
            sys.exit(1)
    
    if args.phase == 3 or args.phase is None:
        success = run_phase3(logger, args.input, args.output)
        if not success:
            logger.error('第三阶段处理失败')
            sys.exit(1)
    
    logger.info('Excel自动化处理工作流运行完成')

if __name__ == "__main__":
    main()