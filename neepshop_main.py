from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import os, sys
import pandas as pd
from neepshop_UI import CustomDialog
from datetime import datetime, timedelta
import logging
from markitdown import MarkItDown
import zipfile
import requests


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
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filepath, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)  # 同时输出到控制台
        ]
    )

    return logging.getLogger(__name__)

cookie_json = 'neepshop.json'
# cookie_json2 = 'neepshop_shixiao.json'
success_url = 'https://www.neep.shop/'
login_url = 'https://cooperation.ceic.com/login/index?client_id=oauth-neep&redirect_uri=https%3A%2F%2Fwww.neep.shop%2Frest%2Fstsso&response_type=code'
excel_url = 'https://www.neep.shop/html/portal/notice.html?type=enquiryOrderAnnc&nodeurl=callback_list_enquiry_order&noticeMoreUrl=https://gd-prod.cn-beijing.oss.aliyuncs.com/upload/cms/column/inquireListOne/index.html&pageTag=undefined&menu_code=&parent_menu_code=&root_menu_code='
pdf_url = 'https://www.neep.shop/dist/index.html#/purchaserNoticeIndex#/purchaserNoticeIndex?autoId=290201'
fabu_time_file = 'public_time.txt'
logger = setup_logging()

api_key = "fastgpt-uktl6lsmWuE6ocGg2adSC2CXPWlB2TLXp87LOHCxq9zRfljK4sPO"
base_url = "http://192.168.50.81:3100/ragai"  # 例如: "https://your-domain.com"
workflow_id = "68cbc237fd26a9e5197e6730"
chat_id = "chat_id"  # 你可以生成一个UUID或使用固定值进行测试
# PDF_FILE_PATH = "宁夏煤业清水营煤矿2025年数字化智能运维管理平台研究与应用技术服务询价采购-商务文件.pdf"
# ZIP_FILE_PATH = "宁夏煤业清水营煤矿2025年数字化智能运维管理平台研究与应用技术服务询价采购-商务文件.zip"
file_id = True


def write_excel(xmmc, bjrzgtj, xjfs, wzfl, fwsj, bjjzsj, fbsj, kw, is_down_pdf):
    try:
        data = {
            '项目名称': [xmmc],
            '报价人资格条件': [bjrzgtj],
            '询价方式': [xjfs],
            '物资分类': [wzfl],
            '服务时间': [fwsj],
            '报价截止时间': [bjjzsj],
            '发布时间': [fbsj],
            '是否下载pdf': [is_down_pdf]
        }
        file_path = os.path.join(os.getcwd(), kw) + '/' + '信息.xlsx'
        try:
            existing_df = pd.read_excel(file_path)
            new_df = pd.DataFrame(data)
            # 将新数据追加到现有的DataFrame中
            combined_df = pd.concat([existing_df, new_df], ignore_index=True)
            # 将合并后的数据写回Excel文件
            combined_df.to_excel(file_path, index=False, sheet_name='信息表')
            logger.info("数据已通过Pandas成功追加并保存！")
        except FileNotFoundError:
            logger.info("文件不存在，创建新文件")
            df = pd.DataFrame(data)
            df.to_excel(file_path, index=False, sheet_name='信息表')
            logger.info("Excel文件已生成！")
    except Exception as e:
        logger.error(f"写入Excel文件时发生错误: {str(e)}")
        raise

def write_excel2(xmmc, ai_read_text, kw):
    try:
        data = {
            '项目名称': [xmmc],
            'AI文本理解提取结果': [ai_read_text]
        }
        file_path = os.path.join(os.getcwd(), kw) + '/' + 'AI文本提取理解.xlsx'
        try:
            existing_df = pd.read_excel(file_path)
            new_df = pd.DataFrame(data)
            # 将新数据追加到现有的DataFrame中
            combined_df = pd.concat([existing_df, new_df], ignore_index=True)
            # 将合并后的数据写回Excel文件
            combined_df.to_excel(file_path, index=False, sheet_name='信息表')
            # logger.info("数据已通过Pandas成功追加并保存！")
        except FileNotFoundError:
            # logger.info("文件不存在，创建新文件")
            df = pd.DataFrame(data)
            df.to_excel(file_path, index=False, sheet_name='信息表')
            # logger.info("Excel文件已生成！")
    except Exception as e:
        # logger.error(f"写入Excel文件时发生错误: {str(e)}")
        raise

def update_content(project_name_list, iframe_locator, context, keyword, old_time):
    project_name = None
    baojiarenzigetiaojian = None
    xunjiafangshi = None
    wuzifenlei = None
    fuwushijian = None
    baojiajiezhishijian = None
    tmp_new_time = None

    logger.info(f"开始处理关键词 '{keyword}' 的项目列表，共 {len(project_name_list)} 个项目")

    for i in range(len(project_name_list)):
        try:
            logger.info(f"处理第 {i + 1} 个项目: {project_name_list[i]}")

            link_locator = iframe_locator.get_by_role('link', name=project_name_list[i])
            href = link_locator.get_attribute('href')
            if not href:
                logger.warning(f"项目 '{project_name_list[i]}' 没有找到链接，跳过")
                continue
            page3 = context.new_page()
            page3.goto(href)
            items = page3.query_selector_all('div.notice-detail div.item')
            items2 = page3.query_selector_all('div.notice-db div.item-db div.item')
            public_times = page3.query_selector_all('div.content_right div.right-main-content div.help-detail-title div.title-other')

            for public_time in public_times:
                if i == 0:
                    tmp_new_time = public_time.inner_text().strip()
                fabushijian = public_time.inner_text().strip()
                try:
                    new_time = datetime.strptime(fabushijian, "%Y-%m-%d %H:%M:%S")
                    logger.info(f"项目发布时间: {fabushijian}")
                except ValueError as e:
                    logger.error(f"时间格式解析错误: {fabushijian}, 错误: {e}")
                    page3.close()
                    continue

            if old_time < new_time:
                logger.info("发现新数据，开始提取信息")
                for item in items:
                    text_content = item.inner_text()
                    if "项目名称" in text_content:
                        project_name = text_content.split("：", 1)[1].strip()
                        logger.info(f"项目名称: {project_name}")
                    elif "报价人资格条件" in text_content:
                        baojiarenzigetiaojian = text_content.split("：", 1)[1].strip()
                        logger.info(f"报价人资格条件: {baojiarenzigetiaojian}")
                    elif "公开询价" in text_content:
                        xunjiafangshi = text_content.split("：", 1)[1].strip()
                        logger.info(f"询价方式: {xunjiafangshi}")
                    elif "服务类->综合服务" in text_content:
                        wuzifenlei = text_content.split("：", 1)[1].strip()
                        logger.info(f"物资分类: {wuzifenlei}")

                for item2 in items2:
                    text_content2 = item2.inner_text()
                    if "服务时间" in text_content2 or "交货时间" in text_content2:
                        fuwushijian = text_content2.split("：", 1)[1].strip()
                        logger.info(f"交货时间或服务时间: {fuwushijian}")
                    elif "报价截止时间" in text_content2:
                        baojiajiezhishijian = text_content2.split("：", 1)[1].strip()
                        logger.info(f"报价截止时间: {baojiajiezhishijian}")

                if project_name and baojiarenzigetiaojian and xunjiafangshi and wuzifenlei and fuwushijian and baojiajiezhishijian:
                    logger.info("项目信息完整，符合条件，开始下载PDF")
                    # ------------------------下载pdf-----------------------
                    page3.goto(pdf_url)
                    try:
                        page3.get_by_role("textbox", name="请输入采购单名称").click(timeout=6000)
                        page3.get_by_role("textbox", name="请输入采购单名称").fill(project_name)
                        page3.get_by_role("button", name="搜索").click()
                        logger.info(f"已搜索采购单: {project_name}")
                    except PlaywrightTimeoutError:
                        logger.warning("账户未登录，无法下载PDF文件")
                        dialog = CustomDialog("账户登录提示", "账户未登录,不能下载pdf文件,请登陆账户,再重新执行程序")
                        dialog.exec()

                    try:
                        # 设置导航等待超时
                        with page3.expect_navigation(timeout=10000):
                            page3.get_by_role("row", name="序号 采购单名称 采购单编号 收到的澄清 日期/周期 发布时间 报价(名)截止时间 采购机构 采购类别").get_by_label("").check()
                            page3.get_by_role("button", name="我要参与").click()
                            page3.get_by_role("button", name="确定").click()

                        title = page3.title()
                        logger.info(f"页面标题: {title}")

                        if '报编' in title:
                            try:
                                page3.get_by_role("button", name="关闭").wait_for(state="visible", timeout=10000)
                                page3.get_by_role("button", name="关闭").click()
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
                            file_path = os.path.join(os.path.join(os.getcwd(), keyword), suggested_filename)
                            # 将文件保存到指定路径（如果已有同名文件，可能会覆盖）
                            download.save_as(file_path)
                            logger.info(f"PDF文件已下载到: {file_path}")
                            write_excel(project_name, baojiarenzigetiaojian, xunjiafangshi, wuzifenlei, fuwushijian,
                                        baojiajiezhishijian, fabushijian, keyword, '是')
                        else:
                            logger.info("页面标题不包含'报编'，跳过下载")
                            write_excel(project_name, baojiarenzigetiaojian, xunjiafangshi, wuzifenlei, fuwushijian,
                                        baojiajiezhishijian, fabushijian, keyword, '否')
                            continue
                    except PlaywrightTimeoutError:
                        logger.warning("导航超时，跳过PDF下载")
                        write_excel(project_name, baojiarenzigetiaojian, xunjiafangshi, wuzifenlei, fuwushijian,
                                    baojiajiezhishijian, fabushijian, keyword, '否')
                        continue
                    # ------------------------下载pdf----------------------
                    # write_excel(project_name, baojiarenzigetiaojian, xunjiafangshi, wuzifenlei, fuwushijian, baojiajiezhishijian, fabushijian, keyword)
                    logger.info("新数据更新成功")
                    # page3.close()
                else:
                    logger.warning("项目信息不完整，不符合条件")
                    # page3.close()
            else:
                logger.info("数据无需更新，跳过处理")
                page3.close()
                break
            page3.close()

        except Exception as e:
            logger.error(f"处理项目 '{project_name_list[i]}' 时发生错误: {str(e)}")
            if 'page3' in locals():
                page3.close()
            continue

    logger.info(f"关键词 '{keyword}' 处理完成")
    return tmp_new_time


def download_excel_pdf(main_folder, fabu_time_file, cookie_json, excel_url):
    logger.info("开始执行下载任务")
    with sync_playwright() as p:
        try:
            bro = p.chromium.launch(headless=False, slow_mo=1000)
            context = bro.new_context(storage_state=cookie_json)
            page = context.new_page()
            logger.info("浏览器启动成功")

            for keyword in main_folder:
                logger.info(f"开始处理关键词: {keyword}")
                main_path = os.path.join(os.getcwd(), keyword)
                if not os.path.exists(main_path):
                    os.mkdir(main_path)
                    logger.info(f"创建主文件夹: {main_path}")
                else:
                    logger.info(f"主文件夹已存在: {main_path}")

                page.goto(excel_url)
                logger.info(f"已打开excel项目地址页面: {excel_url}")
                page.locator("#notice iframe").content_frame.get_by_role("textbox", name="项目名称").click()
                page.locator("#notice iframe").content_frame.get_by_role("textbox", name="项目名称").fill(keyword)
                page.locator("#notice iframe").content_frame.get_by_text("搜索").click()
                logger.info(f"已搜索关键词: {keyword}")

                iframe_locator = page.frame_locator("iframe[src='https://gd-prod.cn-beijing.oss.aliyuncs.com/upload/cms/column/inquireListOne/index.html']")
                project_name_list = iframe_locator.locator('[class="c_href"]').all_text_contents()
                # next_page = iframe_locator.locator('//*[@id="next_page"]').inner_text()
                pagenum = iframe_locator.locator('//*[@id="pageNum"]').inner_text()
                total_page = int(pagenum.lstrip('共').rstrip('页'))
                logger.info(f"找到 {len(project_name_list)} 个项目，共 {total_page} 页")

                if len(project_name_list) == 0:
                    logger.info("没有网页数据")
                    continue
                else:
                    if not os.path.exists(os.path.join(os.getcwd(), keyword, fabu_time_file)):
                        logger.info("时间文件不存在，创建初始时间")
                        # 获取当前日期（时间部分设为00:00:00）
                        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
                        # 计算5天前的日期
                        five_days_ago = today - timedelta(days=5)
                        # 按照指定格式输出
                        today_str = today.strftime("%Y-%m-%d %H:%M:%S")
                        five_days_ago_str = five_days_ago.strftime("%Y-%m-%d %H:%M:%S")
                        logger.info(f"设置初始时间: {five_days_ago_str}")

                        # 创建标记文件
                        with open(os.path.join(os.getcwd(), "public_time", keyword, fabu_time_file), 'w') as f:
                            f.write(five_days_ago_str)

                        with open(os.path.join(os.getcwd(), "public_time", keyword, fabu_time_file), 'r') as file:
                            lines = file.readlines()
                            old_fabu_time = lines[0].strip() if lines else None  # 处理空文件
                            old_time = datetime.strptime(old_fabu_time, "%Y-%m-%d %H:%M:%S")
                            logger.info(f"读取上次更新时间: {old_fabu_time}")

                        # 正常遍历内容
                        tmp_time = update_content(project_name_list, iframe_locator, context, keyword, old_time)
                        # 下一页功能实现
                        for i in range(total_page - 1):
                            logger.info(f"处理第 {i + 2} 页")
                            page.locator("#notice iframe").content_frame.get_by_text("下一页").click()
                            iframe_locator = page.frame_locator("iframe[src='https://gd-prod.cn-beijing.oss.aliyuncs.com/upload/cms/column/inquireListOne/index.html']")
                            project_name_list = iframe_locator.locator('[class="c_href"]').all_text_contents()
                            update_content(project_name_list, iframe_locator, context, keyword, old_time)
                        with open(os.path.join(os.getcwd(), "public_time", keyword, fabu_time_file), 'w', encoding='utf-8') as file:
                            file.write(tmp_time)
                        logger.info(f"更新时间文件: {tmp_time}")

                    # 如果不为空，则取出txt中旧时间，进行对比
                    else:
                        with open(os.path.join(os.getcwd(), keyword, fabu_time_file), 'r') as file:
                            lines = file.readlines()
                            old_fabu_time = lines[0].strip() if lines else None  # 处理空文件
                            old_time = datetime.strptime(old_fabu_time, "%Y-%m-%d %H:%M:%S")

                        # 正常遍历内容
                        tmp_time2 = update_content(project_name_list, iframe_locator, context, keyword, old_time)
                        # 下一页功能实现
                        for i in range(total_page - 1):
                            page.locator("#notice iframe").content_frame.get_by_text("下一页").click()
                            iframe_locator = page.frame_locator(
                                "iframe[src='https://gd-prod.cn-beijing.oss.aliyuncs.com/upload/cms/column/inquireListOne/index.html']")
                            project_name_list = iframe_locator.locator('[class="c_href"]').all_text_contents()
                            update_content(project_name_list, iframe_locator, context, keyword, old_time)
                        with open(os.path.join(os.getcwd(), "public_time", keyword, fabu_time_file), 'w', encoding='utf-8') as file:
                            file.write(tmp_time2)
            logger.info("所有关键词处理完成")

        except Exception as e:
            logger.error(f"执行下载任务时发生错误: {str(e)}")
            raise
        finally:
            page.close()
            context.close()
            bro.close()
            logger.info("浏览器已关闭")


def main(keywords_list):
    is_login_require = False
    main_folder = keywords_list
    # sub_folders = ['excel', 'pdf']
    # 账户cookie登录逻辑
    if os.path.exists(os.path.join(os.getcwd(), cookie_json)):
        with sync_playwright() as p:
            # 启动浏览器，假设get_browser_object函数已实现或直接启动
            browser = p.chromium.launch(headless=False)  # 设为True则无头模式运行
            context = browser.new_context(storage_state=cookie_json)
            page = context.new_page()
            page.goto(login_url)
            # page.wait_for_load_state("networkidle")
            try:
                page.wait_for_url(success_url)
                dialog = CustomDialog("账户登录提示", "账号登陆成功，关闭弹窗，继续执行程序")
                dialog.exec()
                page.close()
                context.close()
                browser.close()
                is_login_require = True
                logger.info("账户登录成功，开始执行程序")
            except Exception as e:
                logger.info("cookie过期失效,请重新登录账户")
                page.close()
                context.close()
                browser.close()
                dialog = CustomDialog("账户登录提示", "cookie过期失效,请重新登录账户")
                dialog.exec()
                os.remove(cookie_json)
                is_login_require = False
            finally:
                page.close()
                context.close()
                browser.close()

    else:
        dialog = CustomDialog("账户登录提示", "请登陆账户")
        dialog.exec()
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)
            context = browser.new_context()
            page = context.new_page()
            page.goto(login_url)
            try:
                page.wait_for_url(success_url, timeout=5*60*1000)
                dialog = CustomDialog("账户登录提示", "账户登录成功,关闭弹窗,继续执行程序")
                dialog.exec()
                storage = context.storage_state(path=cookie_json)
                page.close()
                context.close()
                browser.close()
                is_login_require = True
                logger.info("账户cookie创建成功,开始执行程序")
            except Exception as e:
                logger.info("登录超时或出错")
                page.close()
                context.close()
                browser.close()
                is_login_require = False
                dialog = CustomDialog("账户登录提示", "登陆账户超时或出错,重新登陆")
                dialog.exec()
            finally:
                page.close()
                context.close()
                browser.close()

    if is_login_require:
        dialog = CustomDialog("账户登陆状态", "账户登陆成功True,开始收集数据和下载pdf")
        dialog.exec()
        download_excel_pdf(main_folder, fabu_time_file, cookie_json, excel_url)
        dialog = CustomDialog("账户登陆状态", "收集数据和下载pdf完成")
        dialog.exec()
    else:
        dialog = CustomDialog("账户登陆状态", "账户登陆失败False")
        dialog.exec()

# def main2():
#     api_key = "fastgpt-uktl6lsmWuE6ocGg2adSC2CXPWlB2TLXp87LOHCxq9zRfljK4sPO"
#     base_url = "http://192.168.50.81:3100/ragai"  # 例如: "https://your-domain.com"
#     workflow_id = "68cbc237fd26a9e5197e6730"
#     chat_id = "chat_id"  # 你可以生成一个UUID或使用固定值进行测试
#     # PDF_FILE_PATH = "宁夏煤业清水营煤矿2025年数字化智能运维管理平台研究与应用技术服务询价采购-商务文件.pdf"
#     # ZIP_FILE_PATH = "宁夏煤业清水营煤矿2025年数字化智能运维管理平台研究与应用技术服务询价采购-商务文件.zip"
#     file_id = True
#
#     with zipfile.ZipFile("宁夏煤业清水营煤矿2025年数字化智能运维管理平台研究与应用技术服务询价采购-商务文件.zip", 'r') as zip_ref:
#         zip_ref.extractall("临时文件")
#         print(f"成功解压到: {"临时文件"}")
#
#     md = MarkItDown(docintel_endpoint="<document_intelligence_endpoint>")
#     result = md.convert("宁夏煤业清水营煤矿2025年数字化智能运维管理平台研究与应用技术服务询价采购-商务文件.pdf")
#     input_text = result.text_content[:9999]
#     """
#     步骤2：调用工作流API，传入文件ID和输入文本
#     """
#     url = f"{base_url}/api/v1/chat/completions"
#     headers = {
#         'Authorization': f'Bearer {api_key}',
#         'Content-Type': 'application/json'
#     }
#     # 构建请求数据体
#     data = {
#         "model": "fastgpt-workflow",  # 或者其他指定的模型名
#         "chatId": chat_id,  # 用于保持会话的连续性:cite[9]
#         "workflowId": workflow_id,  # 指定要运行的工作流
#         "messages": [
#             {
#                 "role": "user",
#                 "content": input_text,
#                 # 此处是关键：在消息中关联已上传的文件
#                 "files": [file_id]  # 假设API支持通过`files`字段传递文件ID列表
#             }
#         ]
#     }
#
#     try:
#         response = requests.post(url, json=data, headers=headers)
#         response.raise_for_status()
#         result = response.json()
#         print("工作流调用成功！")
#         # 写入到excel文件
#         write_excel2('宁夏煤业清水营煤矿2025年数字化智能运维管理平台研究与应用技术服务询价采购-商务文件.pdf',
#                      result["choices"][0]["message"]["content"], '软件')
#         # 提取并返回模型的回复内容
#         # return result["choices"][0]["message"]["content"]
#     except Exception as e:
#         print(f"工作流调用失败: {e}")
#         print(f"响应状态码: {response.status_code}")
#         print(f"响应内容: {response.text}")
#         # return None
#
#     # 4. 删除临时文件中的所有文件
#     for file in os.listdir(os.path.join(os.getcwd(), '临时文件')):
#         os.remove(os.path.join(os.getcwd(), '临时文件') + '/' + file)


def main2():
    key_word_list = ['软件', '维保', '运维']
    for keyword in key_word_list:
        # zip_file_path = os.path.join(os.getcwd(), keyword)
        for f in os.listdir(keyword):
            if f.endswith('.zip'):
                print('zip名字: ', f)
                with zipfile.ZipFile(os.path.join(keyword, f), 'r') as zip_ref:
                    zip_ref.extractall("临时文件")
                    print(f"成功解压到: {"临时文件"}")

                for pdf_file in os.listdir('临时文件'):
                    if pdf_file.endswith('.pdf') and '商务' in pdf_file.split('.')[0]:
                        md = MarkItDown(docintel_endpoint="<document_intelligence_endpoint>")
                        result = md.convert('临时文件' + '/' + pdf_file)
                        input_text = result.text_content[:9999]
                        """
                        步骤2：调用工作流API，传入文件ID和输入文本
                        """
                        url = f"{base_url}/api/v1/chat/completions"
                        headers = {
                            'Authorization': f'Bearer {api_key}',
                            'Content-Type': 'application/json'
                        }
                        # 构建请求数据体
                        data = {
                            "model": "fastgpt-workflow",  # 或者其他指定的模型名
                            "chatId": chat_id,  # 用于保持会话的连续性:cite[9]
                            "workflowId": workflow_id,  # 指定要运行的工作流
                            "messages": [
                                {
                                    "role": "user",
                                    "content": input_text,
                                    # 此处是关键：在消息中关联已上传的文件
                                    "files": [file_id]  # 假设API支持通过`files`字段传递文件ID列表
                                }
                            ]
                        }

                        try:
                            response = requests.post(url, json=data, headers=headers)
                            response.raise_for_status()
                            result = response.json()
                            print("工作流调用成功！")
                            # 写入到excel文件
                            write_excel2((pdf_file), result["choices"][0]["message"]["content"], keyword)
                            # 提取并返回模型的回复内容
                            # return result["choices"][0]["message"]["content"]
                        except Exception as e:
                            print(f"工作流调用失败: {e}")
                            print(f"响应状态码: {response.status_code}")
                            print(f"响应内容: {response.text}")
                            # return None

                        # 4. 删除临时文件中的所有文件
                        for file in os.listdir('临时文件'):
                            os.remove('临时文件' + '/' + file)


if __name__ == '__main__':

    keywords_list = ['软件', '运维', '维保']
    try:
        main(keywords_list)
        main2()

    except KeyboardInterrupt:
        logger.info("程序被用户中断")
    except Exception as e:
        logger.error(f"程序异常退出: {str(e)}")
        sys.exit(1)