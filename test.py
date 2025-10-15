import os
import pandas as pd
import requests
import re
import json

def test():

    pattern = r'\{[^{}]*\}'

    input_text = "'项称': '国能程序', 'AI文本理取结果': 'nihao', '间': '111' 改为json格式"
    url = f"http://192.168.10.45:3000/api/v1/chat/completions"
    headers = {
        'Authorization': f'Bearer fastgpt-kRiYi9vHUzZZF55wKnkXjdVmSj4VF1IYaruRgEC59V2cijN0HWZY5CcO4dE7c',
        'Content-Type': 'application/json'
    }
    # 构建请求数据体
    data = {
        "model": "fastgpt-workflow",  # 或者其他指定的模型名
        "chatId": "chat_id",  # 用于保持会话的连续性:cite[9]
        "workflowId": "68ece7bc933a66674938d73a",  # 指定要运行的工作流
        "messages": [
            {
                "role": "user",
                "content": input_text,
                # 此处是关键：在消息中关联已上传的文件
                "files": [True]  # 假设API支持通过`files`字段传递文件ID列表
            }
        ]
    }

    try:
        response = requests.post(url, json=data, headers=headers)
        response.raise_for_status()
        result = response.json()
        # logger.info("工作流调用成功！")
        print('1: ', result)
        print('2:', result["choices"][0]["message"]["content"])
        print('1: ', type(result))
        print('2:', type(result["choices"][0]["message"]["content"]))

        matches = re.findall(pattern, result["choices"][0]["message"]["content"])

        json_str = matches[0].replace('\n', '')
        print(json_str)
        json_dict = json.loads(json_str)
        return json_dict
        # 写入到excel文件
        # write_excel2((pdf_file), result["choices"][0]["message"]["content"], keyword)
        # 提取并返回模型的回复内容
        # return result["choices"][0]["message"]["content"]
    except Exception as e:
        # logger.info(f"工作流调用失败: {e}")
        # logger.info(f"响应状态码: {response.status_code}")
        # logger.info(f"响应内容: {response.text}")
        return None



def write_excel2(text):
    try:

        data = {key: [value] for key, value in text.items()}
        print(data)

        file_path = os.path.join(os.getcwd(), '软件') + '/' + 'AI文本提取理解.xlsx'
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

    # existing_df = pd.read_excel(file_path)
    # trans_df = existing_df.T.reset_index()
    # trans_df.to_excel(file_path, index=False, sheet_name='信息表')

if __name__ == '__main__':
    json = test()
    print(type(json))
    write_excel2(json)

    # file_path = os.path.join(os.getcwd(), '软件') + '/' + 'AI文本提取理解.xlsx'
    # existing_df = pd.read_excel(file_path)
    # trans_df = existing_df.T.reset_index()
    # trans_df.to_excel(file_path, index=False, sheet_name='信息表')



