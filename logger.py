import os
import sys
import logging
from datetime import datetime


def setup_logging():
    """配置日志系统"""
    # 创建logs目录
    log_dir = 'logs'
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    # 设置日志文件名（按日期）
    log_filename = datetime.now().strftime('neepshop_%Y%m%d-%H%M%S.log')
    log_filepath = os.path.join(log_dir, log_filename)

    # 配置日志格式
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(funcName)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filepath, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)  # 同时输出到控制台
        ]
    )

    return logging.getLogger(__name__)

# 初始化日志系统
setup_logging()

# 创建根日志记录器
logger = logging.getLogger()
