import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QLineEdit, QPushButton, QLabel, QMessageBox, QDialog)
import re, time
import neepshop_main
from PyQt6.QtCore import QTimer


# 自定义弹窗
class CustomDialog(QDialog):
    def __init__(self, windowtitle, textlabel):
        super().__init__()
        self.setWindowTitle(windowtitle)
        layout = QVBoxLayout()
        label = QLabel(textlabel)
        button = QPushButton("关闭")
        button.clicked.connect(self.accept)  # 点击按钮关闭弹窗
        layout.addWidget(label)
        layout.addWidget(button)
        self.setLayout(layout)


class MainWindow(QMainWindow, QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("招标网站智能搜索解析工具")
        self.setGeometry(100, 100, 400, 200)

        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 创建布局
        layout = QVBoxLayout()
        central_widget.setLayout(layout)

        # 创建标签
        label = QLabel("step1. 请输入关键字（用逗号、空格或分号分隔）：")
        layout.addWidget(label)

        # 创建输入框
        self.keyword_input = QLineEdit()
        self.keyword_input.setText("软件, 运维, 维保")
        self.keyword_input.setPlaceholderText("例如: 软件, 运维, 维保")
        layout.addWidget(self.keyword_input)

        # 创建按钮
        self.execute_button = QPushButton("01_招标网站智能搜索")
        self.execute_button.clicked.connect(self.execute_main_function)
        layout.addWidget(self.execute_button)

        # 创建标签
        label = QLabel("step2. AI提取并理解pdf内容：(在完成step1的前提下执行step2)")
        layout.addWidget(label)

        # 创建按钮
        self.execute_button2 = QPushButton("02_招标文件智能解析")
        self.execute_button2.clicked.connect(self.execute_main_function2)
        layout.addWidget(self.execute_button2)

        # 初始化关键字列表变量
        self.keyword_list = []

    def parse_keywords(self, text):
        """将输入字符串转换为关键字列表"""

        # 使用正则表达式匹配逗号、分号、空格等作为分隔符
        keywords = re.split(r'[,;\s]+', text)
        # 过滤空字符串
        keywords = [keyword.strip() for keyword in keywords if keyword.strip()]
        return keywords

    def main_function(self, keyword_list):
        """主函数程序 - 这里可以替换为你的实际功能"""
        # print(f"执行主函数，关键字列表: {keyword_list}")

        neepshop_main.main(keyword_list)
        # 显示结果（在实际应用中，你可以根据需要修改这部分）
        # QMessageBox.information(self, "执行结果", "程序执行结束")

    def main_function2(self, keyword_list):
        """主函数程序 - 这里可以替换为你的实际功能"""
        neepshop_main.main2(keyword_list)
        # 显示结果（在实际应用中，你可以根据需要修改这部分）
        QMessageBox.information(self, "执行结果", "AI文档解析执行完毕")

    def execute_main_function(self):
        """执行主函数"""
        # 获取输入文本
        input_text = self.keyword_input.text().strip()

        if not input_text:
            QMessageBox.warning(self, "输入错误", "请输入关键字！")
            return

        # 转换为关键字列表
        self.keyword_list = self.parse_keywords(input_text)

        # 执行主函数
        try:
            self.main_function(self.keyword_list)
            print(f"主函数1-执行完成")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"执行过程中出现错误: {str(e)}")

    def execute_main_function2(self):
        """执行主函数2"""
        # 获取输入文本
        input_text = self.keyword_input.text().strip()

        if not input_text:
            QMessageBox.warning(self, "输入错误", "请输入关键字！")
            return

        # 转换为关键字列表
        self.keyword_list = self.parse_keywords(input_text)

        try:
            self.main_function2(self.keyword_list)
            print(f"主函数2执行完成")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"执行过程中出现错误: {str(e)}")

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
