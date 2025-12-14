import sys
import subprocess
import os
from PySide2.QtWidgets import QApplication, QMessageBox
from card_verify import TechStyleVerifyDialog

# 目标程序配置 - 修改为需要运行的Python脚本文件名
TARGET_SCRIPT = "wechat_monitor.py"  # 例如："wechat_monitor.py"


class VerifyAndLaunch:
    def __init__(self):
        self.app = QApplication(sys.argv)
        # 确保中文显示正常
        font = self.app.font()
        font.setFamily("Microsoft YaHei")
        self.app.setFont(font)

        self.verify_dialog = TechStyleVerifyDialog()
        self.verify_dialog.verify_success.connect(self.launch_target)

    def launch_target(self):
        """验证通过后启动目标程序"""
        # 检查目标文件是否存在
        if not os.path.exists(TARGET_SCRIPT):
            QMessageBox.critical(
                self.verify_dialog,
                "文件错误",
                f"目标程序 {TARGET_SCRIPT} 不存在！\n请检查TARGET_SCRIPT配置是否正确"
            )
            self.verify_dialog.close()
            self.app.quit()
            return

        # 检查是否为Python文件
        if not TARGET_SCRIPT.endswith(".py"):
            QMessageBox.warning(
                self.verify_dialog,
                "格式错误",
                f"{TARGET_SCRIPT} 不是Python脚本文件（.py）"
            )
            self.verify_dialog.close()
            self.app.quit()
            return

        try:
            # 启动目标程序（使用当前Python环境）
            subprocess.Popen([sys.executable, TARGET_SCRIPT])
            # 关闭验证窗口并退出当前程序
            self.verify_dialog.close()
            self.app.quit()
        except Exception as e:
            QMessageBox.critical(
                self.verify_dialog,
                "启动失败",
                f"无法启动程序: {str(e)}"
            )

    def run(self):
        """运行验证流程"""
        self.verify_dialog.show()
        sys.exit(self.app.exec_())


if __name__ == "__main__":
    launcher = VerifyAndLaunch()
    launcher.run()