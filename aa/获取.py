import sys
import time
import requests
import json
import logging
import os
import ctypes
import uuid
from datetime import datetime, timezone
from typing import Dict, Optional, Tuple

# ===================== 1. 核心配置（无修改需求） =====================
# 隐藏控制台黑窗（关键）
try:
    ctypes.windll.kernel32.FreeConsole()  # 释放并隐藏控制台窗口
except:
    pass

# 鸽子云接口配置
PIGEON_CONFIG = {
    "base_url": "https://www.geziyun.cn/api.php",
    "app_id": "945",  # 替换为实际APPID
    "time_check": True,
    "sign_check": False
}

# 日志配置（无控制台输出）
LOG_DIR = os.path.join(os.environ.get("LOCALAPPDATA"), "PigeonCard")
os.makedirs(LOG_DIR, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler(os.path.join(LOG_DIR, "card.log"), encoding="utf-8")]
)
logger = logging.getLogger(__name__)

# ===================== 2. 依赖导入（PySide2） =====================
from PySide2.QtWidgets import (
    QApplication, QDialog, QLabel, QLineEdit,
    QPushButton, QVBoxLayout, QWidget, QMessageBox
)
from PySide2.QtCore import Qt, Signal, QThread, QTimer
from PySide2.QtGui import QFont, QColor, QPalette


# ===================== 3. 核心工具类 =====================
class DeviceTool:
    @staticmethod
    def get_device_id() -> str:
        """生成唯一设备ID"""
        try:
            mac = uuid.getnode()
            hostname = os.environ.get("COMPUTERNAME", "unknown")
            return f"DEV_{mac}_{hostname}".replace(" ", "")[:32]
        except:
            return f"TEMP_{int(time.time() * 1000)}"


class PigeonCloudAPI:
    def __init__(self):
        self.base_url = PIGEON_CONFIG["base_url"]
        self.app_id = PIGEON_CONFIG["app_id"]
        self.device_id = DeviceTool.get_device_id()

    def _request(self, api: str, params: dict = None) -> Tuple[bool, Dict, str]:
        """通用接口请求"""
        params = params or {}
        params["api"] = api
        params["app"] = self.app_id
        if PIGEON_CONFIG["time_check"]:
            params["t"] = int(time.time())
        if api in ["heartbeat", "kmlogon", "kmunmachine"]:
            params["markcode"] = self.device_id

        try:
            resp = requests.get(
                self.base_url, params=params, timeout=8,
                headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
            )
            resp.raise_for_status()
            result = resp.json()
            if result.get("code") != 200:
                return False, result, result.get("msg", "接口调用失败")
            return True, result, ""
        except Exception as e:
            logger.error(f"API[{api}]错误: {str(e)}")
            return False, {}, f"接口错误：{str(e)[:20]}"

    def kmlogon(self, card_key: str) -> Tuple[bool, Dict, str]:
        """卡密登录（核心验证）"""
        return self._request("kmlogon", {"kami": card_key})

    def heartbeat(self, card_key: str, quit: bool = False) -> Tuple[bool, Dict, str]:
        """心跳维持"""
        return self._request("heartbeat", {"kami": card_key, "quit": 1 if quit else 0})


class CardCache:
    """极简缓存管理"""

    def __init__(self):
        self.cache_path = os.path.join(LOG_DIR, "cache.json")
        self.device_id = DeviceTool.get_device_id()

    def save(self, card_key: str, card_info: Dict):
        """保存缓存"""
        try:
            with open(self.cache_path, "w", encoding="utf-8") as f:
                json.dump({
                    "key": card_key,
                    "info": card_info,
                    "device": self.device_id,
                    "time": datetime.now().isoformat()
                }, f)
        except:
            pass

    def load(self) -> Tuple[Optional[str], Optional[Dict]]:
        """加载缓存"""
        try:
            if not os.path.exists(self.cache_path):
                return None, None
            with open(self.cache_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if data.get("device") != self.device_id:
                return None, None
            return data.get("key"), data.get("info")
        except:
            return None, None

    def clear(self):
        """清空缓存"""
        try:
            if os.path.exists(self.cache_path):
                os.remove(self.cache_path)
        except:
            pass


class TimeThread(QThread):
    """极简倒计时线程"""
    time_update = Signal(str, str)
    time_expire = Signal(str)

    def __init__(self, card_key: str, expire_ts: int):
        super().__init__()
        self.card_key = card_key
        self.expire_time = datetime.fromtimestamp(expire_ts).astimezone()
        self.running = True

    def run(self):
        while self.running:
            now = datetime.now(timezone.utc).astimezone()
            remain = (self.expire_time - now).total_seconds()
            if remain <= 0:
                self.running = False
                self.time_expire.emit(self.card_key)
                CardCache().clear()
                remain = 0

            # 格式化时间
            d = int(remain // 86400)
            h = int((remain % 86400) // 3600)
            m = int((remain % 3600) // 60)
            s = int(remain % 60)
            self.time_update.emit(
                f"{d:02d}天{h:02d}时{m:02d}分{s:02d}秒",
                self.expire_time.strftime("%Y-%m-%d %H:%M:%S")
            )
            time.sleep(1)

    def stop(self):
        self.running = False
        self.wait(500)


# ===================== 4. 科技风登录窗口（核心优化） =====================
class LoginDialog(QDialog):
    login_success = Signal(str, int)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("卡密验证")
        self.setFixedSize(320, 280)  # 小巧尺寸
        self.setStyleSheet("""
            QDialog {
                background: #0a0e17;  /* 深空黑底色 */
                border: 1px solid #00f0ff;  /* 霓虹蓝边框 */
                border-radius: 8px;
            }
            QLabel {
                color: #ffffff;
                font-family: "Microsoft YaHei", "微软雅黑";
            }
            QLabel#title {
                color: #00f0ff;
                font-size: 16px;
                font-weight: bold;
                letter-spacing: 1px;
            }
            QLabel#time {
                color: #70e0ff;
                font-size: 11px;
                margin-top: 5px;
            }
            QLabel#status {
                color: #ff6b6b;
                font-size: 10px;
                margin-top: 3px;
            }
            QLineEdit {
                background: #141b2d;
                border: 1px solid #00a8cc;
                border-radius: 4px;
                color: #ffffff;
                font-size: 12px;
                padding: 6px 8px;
                font-family: "Microsoft YaHei", "微软雅黑";
                selection-background-color: #00f0ff;
            }
            QLineEdit:focus {
                border-color: #00f0ff;
                outline: none;
                box-shadow: 0 0 5px #00f0ff;
            }
            QPushButton {
                background: linear-gradient(to right, #008fb3, #00f0ff);
                border: none;
                border-radius: 4px;
                color: #0a0e17;
                font-size: 12px;
                font-weight: bold;
                padding: 6px 0;
                font-family: "Microsoft YaHei", "微软雅黑";
                letter-spacing: 0.5px;
            }
            QPushButton:hover {
                background: linear-gradient(to right, #00a8cc, #00ffff);
            }
            QPushButton:disabled {
                background: #2c3e50;
                color: #7f8c8d;
            }
        """)

        # 统一字体（解决显示异常）
        font = QFont("Microsoft YaHei", 9)
        self.setFont(font)

        # 布局（极简紧凑）
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)

        # 标题
        title = QLabel("卡密验证系统")
        title.setObjectName("title")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # 卡密输入框（13位）
        self.card_input = QLineEdit()
        self.card_input.setPlaceholderText("请输入13位卡密")
        self.card_input.setMaxLength(13)  # 13位卡密
        self.card_input.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.card_input)

        # 验证按钮
        self.login_btn = QPushButton("验证并进入")
        self.login_btn.clicked.connect(self._verify_card)
        layout.addWidget(self.login_btn)

        # 剩余时间显示（初始隐藏）
        self.time_label = QLabel("剩余时长：-- --:--:--")
        self.time_label.setObjectName("time")
        self.time_label.setAlignment(Qt.AlignCenter)
        self.time_label.setVisible(False)
        layout.addWidget(self.time_label)

        # 到期时间显示（初始隐藏）
        self.expire_label = QLabel("到期时间：---- --:--:--")
        self.expire_label.setObjectName("time")
        self.expire_label.setAlignment(Qt.AlignCenter)
        self.expire_label.setVisible(False)
        layout.addWidget(self.expire_label)

        # 状态提示
        self.status_label = QLabel("")
        self.status_label.setObjectName("status")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)

        self.setLayout(layout)

        # 初始化缓存
        self.cache = CardCache()
        self.api = PigeonCloudAPI()
        self.time_thread = None
        self._load_cache()

    def _load_cache(self):
        """加载缓存卡密"""
        card_key, card_info = self.cache.load()
        if card_key and card_info:
            self.card_input.setText(card_key)
            self._verify_card(use_cache=True)

    def _verify_card(self, use_cache: bool = False):
        """卡密验证核心逻辑"""
        card_key = self.card_input.text().strip() if not use_cache else self.cache.load()[0]

        # 13位校验
        if len(card_key) != 13:
            self.status_label.setText("卡密必须为13位！")
            return

        self.login_btn.setEnabled(False)
        self.status_label.setText("验证中...")

        # 调用鸽子云接口
        success, result, error = self.api.kmlogon(card_key)
        if not success:
            self.status_label.setText(f"验证失败：{error}")
            self.login_btn.setEnabled(True)
            self.cache.clear()
            return

        # 解析过期时间
        try:
            expire_ts = int(result.get("msg", {}).get("vip", 0))
            if expire_ts == 0:
                raise ValueError("无有效时长")

            # 启动心跳
            self.api.heartbeat(card_key)
            # 保存缓存
            self.cache.save(card_key, result)

            # 显示时间
            self.time_label.setVisible(True)
            self.expire_label.setVisible(True)
            self.status_label.setText("验证成功！即将进入主程序")

            # 启动倒计时
            self._start_time_thread(card_key, expire_ts)

            # 通知主程序
            self.login_success.emit(card_key, expire_ts)
            QTimer.singleShot(1500, self.close)

        except Exception as e:
            self.status_label.setText(f"解析失败：{str(e)[:15]}")
            self.login_btn.setEnabled(True)
            self.cache.clear()

    def _start_time_thread(self, card_key: str, expire_ts: int):
        """启动倒计时"""
        if self.time_thread:
            self.time_thread.stop()
        self.time_thread = TimeThread(card_key, expire_ts)
        self.time_thread.time_update.connect(self._update_time)
        self.time_thread.time_expire.connect(self._on_expire)
        self.time_thread.start()

    def _update_time(self, remain: str, expire: str):
        """更新时间显示"""
        self.time_label.setText(f"剩余时长：{remain}")
        self.expire_label.setText(f"到期时间：{expire}")

    def _on_expire(self, card_key: str):
        """卡密过期"""
        self.status_label.setText("卡密已过期！")
        self.card_input.clear()
        self.cache.clear()
        self.login_btn.setEnabled(True)

    def closeEvent(self, event):
        """关闭时停止线程"""
        if self.time_thread:
            self.time_thread.stop()
        event.accept()


# ===================== 5. 科技风主窗口 =====================
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("主程序")
        self.setFixedSize(400, 300)  # 小巧主窗口
        self.setStyleSheet("""
            QWidget {
                background: #0a0e17;
                border: 1px solid #00f0ff;
                border-radius: 8px;
            }
            QLabel {
                color: #ffffff;
                font-family: "Microsoft YaHei", "微软雅黑";
            }
            QLabel#title {
                color: #00f0ff;
                font-size: 16px;
                font-weight: bold;
            }
            QLabel#time {
                color: #70e0ff;
                font-size: 12px;
                background: #141b2d;
                padding: 4px;
                border-radius: 4px;
            }
            QPushButton {
                background: linear-gradient(to right, #008fb3, #00f0ff);
                border: none;
                border-radius: 4px;
                color: #0a0e17;
                font-size: 11px;
                padding: 5px 15px;
                font-family: "Microsoft YaHei", "微软雅黑";
            }
            QPushButton:hover {
                background: linear-gradient(to right, #00a8cc, #00ffff);
            }
        """)

        # 统一字体
        font = QFont("Microsoft YaHei", 9)
        self.setFont(font)

        # 布局
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        # 标题
        title = QLabel("主程序功能界面")
        title.setObjectName("title")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # 时间显示栏
        self.time_label = QLabel("剩余时长：--天--时--分--秒")
        self.time_label.setObjectName("time")
        self.time_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.time_label)

        # 功能按钮区
        btn_layout = QVBoxLayout()
        btn_layout.setSpacing(8)

        btn1 = QPushButton("功能按钮1")
        btn2 = QPushButton("功能按钮2")
        btn3 = QPushButton("更换卡密")
        btn3.clicked.connect(self._open_login)

        btn_layout.addWidget(btn1)
        btn_layout.addWidget(btn2)
        btn_layout.addWidget(btn3)
        layout.addLayout(btn_layout)

        # 占位（可替换为实际功能）
        layout.addStretch(1)

        self.setLayout(layout)

        # 初始化
        self.cache = CardCache()
        self.api = PigeonCloudAPI()
        self.time_thread = None
        self._init_time()

    def _init_time(self):
        """初始化倒计时"""
        card_key, card_info = self.cache.load()
        if card_key and card_info:
            expire_ts = int(card_info.get("msg", {}).get("vip", 0))
            if expire_ts > 0:
                self._start_time_thread(card_key, expire_ts)

    def _start_time_thread(self, card_key: str, expire_ts: int):
        """启动倒计时"""
        if self.time_thread:
            self.time_thread.stop()
        self.time_thread = TimeThread(card_key, expire_ts)
        self.time_thread.time_update.connect(lambda r, e: self.time_label.setText(f"剩余时长：{r}"))
        self.time_thread.time_expire.connect(self._on_expire)
        self.time_thread.start()

    def _open_login(self):
        """打开登录窗口"""
        login = LoginDialog()
        login.login_success.connect(lambda k, t: self._start_time_thread(k, t))
        login.exec_()

    def _on_expire(self, card_key: str):
        """卡密过期"""
        self.time_label.setText("剩余时长：已过期")
        QMessageBox.warning(self, "提示", "卡密已过期，请重新验证！")
        self._open_login()

    def closeEvent(self, event):
        """关闭时停止线程+发送心跳退出"""
        if self.time_thread:
            self.time_thread.stop()
        card_key = self.cache.load()[0]
        if card_key:
            self.api.heartbeat(card_key, quit=True)
        event.accept()


# ===================== 6. 程序入口 =====================
if __name__ == "__main__":
    # 高DPI适配（解决字体模糊/显示不全）
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps)

    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(True)

    # 检查缓存
    cache = CardCache()
    card_key, card_info = cache.load()

    if card_key and card_info:
        # 直接打开主窗口
        window = MainWindow()
        window.show()
    else:
        # 先打开登录窗口
        login = LoginDialog()
        if login.exec_():
            window = MainWindow()
            window.show()

    sys.exit(app.exec_())