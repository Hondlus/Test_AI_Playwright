import os

os.system("powershell -Command \"pip install playwright==1.54.0 pandas==2.3.1 openpyxl==3.1.5 PyQt6==6.9.1 'markitdown[pdf, docx, pptx, az-doc-intel]==0.1.3' magika==0.6.2 -i https://mirrors.aliyun.com/pypi/simple\"")
os.system("playwright install chromium")
# pyinstaller==6.16.0 pyinstaller-hooks-contrib==2025.8
# 打包命令:pyinstaller -w -F .\neepshop_UI.py .\neepshop_main.py .\logger.py --add-data D:\ProgramData\Anaconda3\envs\py31293\Lib\site-packages\playwright:playwright/ --add-data D:\ProgramData\Anaconda3\envs\py31293\Lib\site-packages\markitdown:markitdown/ --add-data D:\ProgramData\Anaconda3\envs\py31293\Lib\site-packages\magika:magika
# 打包命令2:pyinstaller -w -F .\neepshop_UI.py .\neepshop_main.py .\logger.py --add-data C:\Users\dxw-user\AppData\Local\ms-playwright:playwright --add-data C:\Users\DXW\AppData\Local\Programs\Python\Python312\Lib\site-packages\markitdown:markitdown/ --add-data C:\Users\DXW\AppData\Local\Programs\Python\Python312\Lib\site-packages\magika:magika

