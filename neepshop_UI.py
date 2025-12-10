import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QLineEdit, QPushButton, QLabel, QMessageBox, QDialog, QHBoxLayout)
import re
import neepshop_main
import traceback


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
        screen = QApplication.primaryScreen().availableGeometry()
        x = (screen.width() - self.width()) // 2
        y = (screen.height() - self.height()) // 2
        self.setGeometry(x, y, 400, 200)

        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 创建布局
        layout = QVBoxLayout()
        central_widget.setLayout(layout)

        # 搜索关键词水平布局
        keyword_layout = QHBoxLayout()
        # 创建标签
        label = QLabel("搜索关键字(多个关键字中间用逗号隔开): ")
        # 创建输入框
        self.keyword_input = QLineEdit()
        self.keyword_input.setText("系统, 软件, 运维, 维保")
        self.keyword_input.setPlaceholderText("例如: 系统, 软件, 运维, 维保")
        keyword_layout.addWidget(label)
        keyword_layout.addWidget(self.keyword_input)
        layout.addLayout(keyword_layout)

        layout.addSpacing(15)

        # 分步按钮垂直布局
        step_layout = QVBoxLayout()
        # 创建按钮
        self.execute_button = QPushButton("分步执行(Step1: 招标网站智能搜索)")
        self.execute_button.clicked.connect(self.execute_main_function)
        step_layout.addWidget(self.execute_button)

        # step_layout.addSpacing(5)

        # 创建按钮
        self.execute_button2 = QPushButton("分步执行(Step2: 招标文件智能解析)")
        self.execute_button2.clicked.connect(self.execute_main_function2)
        step_layout.addWidget(self.execute_button2)
        layout.addLayout(step_layout)

        layout.addSpacing(15)

        self.onekey_button = QPushButton("一键执行(Step1 + Step2)")
        self.onekey_button.clicked.connect(self.execute_onekey_function)
        layout.addWidget(self.onekey_button)

        layout.addSpacing(15)

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
        QMessageBox.information(self, "执行结果", "程序执行结束")

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
            # QMessageBox.critical(self, "错误", f"主函数1执行过程中出现错误: {str(e)}")
            error_msg = traceback.format_exc()
            QMessageBox.critical(self, "错误", f"主函数1执行过程中出现错误: {str(error_msg)}")

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
            # QMessageBox.critical(self, "错误", f"主函数2执行过程中出现错误: {str(e)}")
            error_msg = traceback.format_exc()
            QMessageBox.critical(self, "错误", f"主函数2执行过程中出现错误: {str(error_msg)}")

    def execute_onekey_function(self):
        """一键执行函数"""
        # 获取输入文本
        input_text = self.keyword_input.text().strip()

        if not input_text:
            QMessageBox.warning(self, "输入错误", "请输入关键字！")
            return

        # 转换为关键字列表
        self.keyword_list = self.parse_keywords(input_text)

        try:
            self.main_function(self.keyword_list)
            self.main_function2(self.keyword_list)
            print(f"一键执行完成")
        except Exception as e:
            # QMessageBox.critical(self, "错误", f"一键执行过程中出现错误: {str(e)}")
            error_msg = traceback.format_exc()
            QMessageBox.critical(self, "错误", f"执行过程中出现错误: {str(error_msg)}")


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
