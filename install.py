import os

os.system("powershell -Command \"pip install playwright pandas openpyxl PyQt6 'markitdown[pdf, docx, pptx, az-doc-intel]' magika -i https://mirrors.aliyun.com/pypi/simple\"")
os.system("playwright install chromium")
