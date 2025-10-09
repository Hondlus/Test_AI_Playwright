import requests
import json
import os
import pandas as pd
from markitdown import MarkItDown
import zipfile


def extract_zip(zip_path, extract_to="临时文件"):
    """
    解压ZIP文件
    Args:
        zip_path: ZIP文件路径
        extract_to: 解压目标目录（默认为ZIP文件所在目录）
    """

    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)
        print(f"成功解压到: {extract_to}")

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


def upload_pdf_to_fastgpt(api_key, base_url, file_path):
    md = MarkItDown(docintel_endpoint="<document_intelligence_endpoint>")
    result = md.convert("宁夏煤业清水营煤矿2025年数字化智能运维管理平台研究与应用技术服务询价采购-商务文件.pdf")
    # print(result.text_content)
    return result.text_content


def call_fastgpt_workflow(api_key, base_url, workflow_id, chat_id, file_id, input_text):
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
        # 提取并返回模型的回复内容
        return result["choices"][0]["message"]["content"]
    except Exception as e:
        print(f"工作流调用失败: {e}")
        print(f"响应状态码: {response.status_code}")
        print(f"响应内容: {response.text}")
        return None


# 使用示例
if __name__ == "__main__":
    # 配置你的信息
    API_KEY = "fastgpt-uktl6lsmWuE6ocGg2adSC2CXPWlB2TLXp87LOHCxq9zRfljK4sPO"
    BASE_URL = "http://192.168.50.81:3100/ragai"  # 例如: "https://your-domain.com"
    WORKFLOW_ID = "68cbc237fd26a9e5197e6730"
    CHAT_ID = "chat_id"  # 你可以生成一个UUID或使用固定值进行测试
    PDF_FILE_PATH = "宁夏煤业清水营煤矿2025年数字化智能运维管理平台研究与应用技术服务询价采购-商务文件.pdf"
    ZIP_FILE_PATH = "宁夏煤业清水营煤矿2025年数字化智能运维管理平台研究与应用技术服务询价采购-商务文件.zip"
    FILE_ID = True

    # 1. 解压缩zip文件
    extract_zip(ZIP_FILE_PATH)

    # 2. 上传PDF给markitdown解析,获取markdown内容
    markdown_result = upload_pdf_to_fastgpt(API_KEY, BASE_URL, PDF_FILE_PATH)
    INPUT_TEXT = markdown_result[:9999] # 截取前10000字符
    # 3. 调用工作流,使用ai获取文本提取内容，并写入excel中
    workflow_result = call_fastgpt_workflow(API_KEY, BASE_URL, WORKFLOW_ID, CHAT_ID, FILE_ID, INPUT_TEXT)
    if workflow_result:
        print("\n--- FastGPT 工作流返回结果 ---")
        print(workflow_result)
        write_excel2('宁夏煤业清水营煤矿2025年数字化智能运维管理平台研究与应用技术服务询价采购-商务文件.pdf', workflow_result, '软件')
    # 4. 删除临时文件中的所有文件
    for file in os.listdir(os.path.join(os.getcwd(), '临时文件')):
        # print(file)
        os.remove(os.path.join(os.getcwd(), '临时文件') + '/' + file)