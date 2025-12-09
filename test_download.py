from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import os, sys, re, json
from datetime import datetime, timedelta
import logging


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


project_name = "内蒙古能源集团有限公司内蒙古长城发电有限公司治安反恐系统建设项目询价采购"
pdf_url = 'https://www.neep.shop/dist/index.html#/purchaserNoticeIndex#/purchaserNoticeIndex?autoId=290201'
cookie_json = "neepshop.json"
logger = setup_logging()

with sync_playwright() as p:
    # 启动浏览器，假设get_browser_object函数已实现或直接启动
    browser = p.chromium.launch(headless=False)  # 设为True则无头模式运行
    context = browser.new_context(storage_state=cookie_json)
    page3 = context.new_page()

    # ------------------------下载pdf-----------------------
    page3.goto(pdf_url)
    page3.get_by_role("textbox", name="请输入采购单名称").click(timeout=6000)
    page3.get_by_role("textbox", name="请输入采购单名称").fill(project_name)
    page3.get_by_role("button", name="搜索").click()
    page3.get_by_role("row", name="序号 采购单名称 采购单编号 收到的澄清 日期/周期 发布时间 报价(名)截止时间 采购机构 采购类别").get_by_label("").check(timeout=2000)

    try:
        # 设置导航等待超时
        with page3.expect_navigation(timeout=10000):
            page3.get_by_role("button", name="我要参与").click()
            page3.get_by_role("button", name="确定").click()
            try:
                page3.get_by_role("button", name="确定").click(timeout=1000)
            except Exception as e:
                logger.error(f"没有多余弹窗按钮A: {e}")

        logger.info("页面发生了跳转, 加载页面A:报编")
        try:
            # page3.get_by_role("button", name="关闭").wait_for(timeout=2000)
            page3.get_by_role("button", name="关闭").click(timeout=10000)
            logger.info("已关闭弹窗")
        except PlaywrightTimeoutError:
            logger.info("页面没有关闭按钮，直接下载")

        download_button = page3.get_by_role("button", name=" 下载采购文件")
        with page3.expect_download() as download_info:
            download_button.click()

        # 获取下载对象
        download = download_info.value
        # 等待下载文件完成并获取建议的文件名
        suggested_filename = download.suggested_filename
        file_path = os.path.join(os.path.join(os.getcwd(), '临时文件'), suggested_filename)
        # 将文件保存到指定路径（如果已有同名文件，可能会覆盖）
        download.save_as(file_path)
        logger.info(f"PDF文件已下载到: {file_path}")

    except Exception as e:
        try:
            with context.expect_page(timeout=10000) as new_page_info:
                page3.get_by_role("row", name="序号 采购单名称 采购单编号 收到的澄清 日期/周期 发布时间 报价(名)截止时间 采购机构 采购类别").get_by_label("").check()
                page3.get_by_role("button", name="我要参与").click()
                page3.get_by_role("button", name="确定").click()
                try:
                    page3.get_by_role("button", name="确定").click(timeout=5000)
                except Exception as e:
                    logger.error(f"没有多余弹窗按钮B: {e}")

            new_page = new_page_info.value
            logger.info("加载页面B:供应商询比价管理")

            try:
                new_page.wait_for_selector('a.fileOperation-btn', timeout=10000)
                download_elements = new_page.query_selector_all('a.fileOperation-btn')
                logger.info(f"找到 {len(download_elements)} 个下载链接")

                for download_element in download_elements:
                    with new_page.expect_download() as download_info:
                        download_element.click()
                        # 获取下载对象
                        download = download_info.value
                        # 等待下载文件完成并获取建议的文件名
                        suggested_filename = download.suggested_filename
                        file_path = os.path.join(
                            os.path.join(os.getcwd(), '临时文件'),
                            suggested_filename)
                        # 将文件保存到指定路径（如果已有同名文件，可能会覆盖）
                        download.save_as(file_path)
                        logger.info(f"WORD文件已下载到: {file_path}")
            except Exception as e:
                logger.info(f"页面下载按钮版面变化导致错误: {e}")
            new_page.close()

        except Exception as e:
            logger.info(f"未知错误: {e}")
            logger.info(f"未搜索到 {project_name} 相关下载文件网页")
    # ------------------------下载pdf----------------------
    page3.close()
    context.close()
    browser.close()
