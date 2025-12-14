import sys
import time
import socket
import struct
import sqlite3
import logging
from datetime import datetime, timedelta, timezone
from pathlib import Path
import ctypes
import wmi
from typing import Dict, Optional, Tuple
from cryptography.fernet import Fernet
from PySide2.QtWidgets import (QApplication, QDialog, QLabel, QLineEdit,
                               QPushButton, QVBoxLayout, QFrame, QMessageBox)
from PySide2.QtCore import Qt, Signal, QThread
from PySide2.QtGui import QFont

# ===================== 核心配置（可根据需求调整） =====================
# 加密密钥（正式环境建议替换为固定密钥，避免每次启动变化）
ENCRYPT_KEY = b'r6F0hN0Il7D5Cp07pgkY-vdb-QkpNG_X0iuY9WFZra8='
# 数据库路径（用户指定：C:\ProgramData\Q\S）
CARD_DB_PATH = Path("C:/ProgramData/Q/S/card_activation.db")
KEY_LENGTH = 16  # 16位卡密
TEST_MODE = False  # 测试模式：True=15秒有效期，False=正式模式（30/360天）
TEST_VALID_SECONDS = 15  # 测试模式有效期

# 预定义卡密（A=360天，B=30天）
PREDEFINED_CARDS: Dict[str, int] = {
    # A类（360天，5个）
    "HN7HC3MEZH9Y7AZ8": 360,
    "7ZN24RMPJ4KDEAY9": 360,
    "TUPQTDSZM88MW4A7": 360,
    "LK4EHSWNGB8164A2": 360,
    "6HNK4V16AVVFJMA5": 360,
    # B类（30天，示例3个）
    "PUVZBZX5CYMEZUB1": 30,
    "1GM537624S3AUVB7": 30,
    "KMU8CZW7DZB71TB9": 30
}

# 日志配置（调试用）
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("C:/ProgramData/Q/S/card_verify.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


# ===================== 1. NTP UTC时间获取（替换原HTTP时间） =====================
def get_utc_time() -> datetime:
    """
    通过NTP协议获取精准UTC网络时间（优先），失败则返回本地UTC时间
    返回：带UTC时区的datetime对象
    """
    ntp_servers = ["pool.ntp.org", "time.nist.gov", "ntp.aliyun.com", "time.cloudflare.com"]
    ntp_packet = bytearray([0x1B]) + bytearray(47)  # NTP请求包
    NTP_EPOCH_DIFF = 2208988800  # NTP时间戳与Unix时间戳差值

    for server in ntp_servers:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
                sock.settimeout(3)
                sock.sendto(ntp_packet, (server, 123))
                data, _ = sock.recvfrom(1024)

                # 解析NTP响应
                seconds = struct.unpack('!I', data[40:44])[0]
                fraction = struct.unpack('!I', data[44:48])[0]
                timestamp = seconds - NTP_EPOCH_DIFF + (fraction / 2 ** 32)

                utc_time = datetime.fromtimestamp(timestamp, tz=timezone.utc)
                logger.info(f"成功从NTP服务器 {server} 获取UTC时间: {utc_time}")
                return utc_time
        except (socket.timeout, socket.error, struct.error, IndexError):
            logger.warning(f"NTP服务器 {server} 访问失败，尝试下一个")
            continue

    # 所有NTP服务器失败，返回本地UTC时间
    local_utc = datetime.now(timezone.utc)
    logger.warning(f"所有NTP服务器失败，使用本地UTC时间: {local_utc}")
    QMessageBox.warning(None, "时间警告", "网络时间获取失败，使用本地时间（可能有误差）！")
    return local_utc


def get_local_time() -> datetime:
    """将UTC时间转换为本地时间（用于显示）"""
    utc_time = get_utc_time()
    return utc_time.astimezone()


# ===================== 2. 加密/设备工具类 =====================
class CryptoTool:
    """数据加密工具（防数据库手动修改）"""

    def __init__(self):
        try:
            self.cipher = Fernet(ENCRYPT_KEY)
        except Exception as e:
            logger.error(f"加密工具初始化失败: {e}")
            raise

    def encrypt(self, data: str) -> str:
        try:
            return self.cipher.encrypt(data.encode()).decode()
        except Exception as e:
            logger.error(f"加密失败 {data}: {e}")
            return ""

    def decrypt(self, encrypted_data: str) -> str:
        try:
            if not encrypted_data:
                return ""
            return self.cipher.decrypt(encrypted_data.encode()).decode()
        except Exception as e:
            logger.error(f"解密失败: {e}")
            return ""


class DeviceTool:
    """设备唯一ID获取（绑定卡密）"""

    @staticmethod
    def get_device_id() -> str:
        try:
            c = wmi.WMI()
            # 主板+CPU组合ID（更唯一）
            board = next(c.Win32_BaseBoard(), None)
            cpu = next(c.Win32_Processor(), None)
            board_sn = board.SerialNumber.strip() if (board and board.SerialNumber) else ""
            cpu_id = cpu.ProcessorId.strip() if (cpu and cpu.ProcessorId) else ""
            device_id = f"{board_sn}_{cpu_id}"
            if device_id.strip():
                return device_id
            # 兜底生成临时ID
            temp_id = f"DEV_{int(time.time() * 1000000)}"
            logger.warning(f"未获取硬件ID，使用临时ID: {temp_id}")
            return temp_id
        except Exception as e:
            temp_id = f"DEV_{int(time.time() * 1000000)}"
            logger.error(f"获取设备ID失败: {e}，使用临时ID: {temp_id}")
            return temp_id


# ===================== 3. 数据库操作类（修复所有语法错误） =====================
class ActivationDatabase:
    def __init__(self):
        self.crypto = CryptoTool()
        self.device_id = DeviceTool.get_device_id()
        self._ensure_dir()  # 确保目录存在
        self._init_table()  # 初始化表（修复ALTER TABLE占位符问题）
        logger.info("数据库初始化完成")

    def _ensure_dir(self):
        """创建数据库目录"""
        try:
            CARD_DB_PATH.parent.mkdir(parents=True, exist_ok=True)
            logger.info(f"数据库目录已创建: {CARD_DB_PATH.parent}")
        except Exception as e:
            logger.error(f"创建目录失败: {e}")
            QMessageBox.critical(None, "目录错误", f"无法创建目录 {CARD_DB_PATH.parent}\n{e}")
            raise

    def _init_table(self):
        """初始化表结构（修复ALTER TABLE语法错误）"""
        try:
            conn = sqlite3.connect(CARD_DB_PATH)
            cursor = conn.cursor()

            # 1. 创建激活记录表（基础结构）
            cursor.execute('''
                           CREATE TABLE IF NOT EXISTS activation_records
                           (
                               encrypted_key
                               TEXT
                               PRIMARY
                               KEY,
                               encrypted_device_id
                               TEXT
                               NOT
                               NULL,
                               encrypted_first_activate_time
                               TEXT
                               NOT
                               NULL,
                               encrypted_initial_remaining_seconds
                               TEXT
                               NOT
                               NULL,
                               last_check_time
                               TEXT
                           )
                           ''')

            # 2. 检查并新增encrypted_is_expired列（修复占位符问题）
            cursor.execute("PRAGMA table_info(activation_records)")
            columns = [col[1] for col in cursor.fetchall()]
            if "encrypted_is_expired" not in columns:
                # 直接拼接加密后的默认值（无占位符，避免语法错误）
                default_val = self.crypto.encrypt("False")
                cursor.execute(f'''
                    ALTER TABLE activation_records 
                    ADD COLUMN encrypted_is_expired TEXT NOT NULL DEFAULT '{default_val}'
                ''')
                logger.info("已新增encrypted_is_expired列")

            # 3. 创建上次输入记录表
            cursor.execute('''
                           CREATE TABLE IF NOT EXISTS last_input
                           (
                               id
                               INTEGER
                               PRIMARY
                               KEY
                               DEFAULT
                               1,
                               encrypted_key
                               TEXT
                           )
                           ''')
            conn.commit()
        except sqlite3.Error as e:
            logger.error(f"数据库初始化失败: {e}")
            QMessageBox.critical(None, "数据库错误", f"表初始化失败: {e}")
            raise
        finally:
            if 'conn' in locals():
                conn.close()

    def get_last_input_key(self) -> str:
        """获取上次输入的卡密"""
        try:
            conn = sqlite3.connect(CARD_DB_PATH)
            cursor = conn.cursor()
            cursor.execute("SELECT encrypted_key FROM last_input WHERE id=1")
            result = cursor.fetchone()
            return self.crypto.decrypt(result[0]) if (result and result[0]) else ""
        except Exception as e:
            logger.error(f"获取上次卡密失败: {e}")
            return ""
        finally:
            conn.close() if 'conn' in locals() else None

    def save_last_input_key(self, key: str):
        """保存上次输入的卡密"""
        try:
            conn = sqlite3.connect(CARD_DB_PATH)
            cursor = conn.cursor()
            encrypted_key = self.crypto.encrypt(key)
            cursor.execute('''
                INSERT OR REPLACE INTO last_input (id, encrypted_key) 
                VALUES (1, ?)
            ''', (encrypted_key,))
            conn.commit()
            logger.info(f"保存上次卡密: {key[:4]}****")
        except Exception as e:
            logger.error(f"保存上次卡密失败: {e}")
        finally:
            conn.close() if 'conn' in locals() else None

    def clear_last_input_key(self):
        """清除过期卡密的记录"""
        try:
            conn = sqlite3.connect(CARD_DB_PATH)
            cursor = conn.cursor()
            cursor.execute("UPDATE last_input SET encrypted_key = NULL WHERE id=1")
            conn.commit()
            logger.info("已清除上次卡密记录")
        except Exception as e:
            logger.error(f"清除上次卡密失败: {e}")
        finally:
            conn.close() if 'conn' in locals() else None

    def check_card_activated(self, key: str) -> bool:
        """检查卡密是否已激活"""
        try:
            conn = sqlite3.connect(CARD_DB_PATH)
            cursor = conn.cursor()
            encrypted_key = self.crypto.encrypt(key)
            cursor.execute("SELECT 1 FROM activation_records WHERE encrypted_key=?", (encrypted_key,))
            return cursor.fetchone() is not None
        except Exception as e:
            logger.error(f"检查卡密激活状态失败: {e}")
            return False
        finally:
            conn.close() if 'conn' in locals() else None

    def get_card_status(self, key: str) -> Tuple[Optional[dict], bool]:
        """获取卡密状态：(激活信息, 是否过期)"""
        try:
            conn = sqlite3.connect(CARD_DB_PATH)
            cursor = conn.cursor()
            encrypted_key = self.crypto.encrypt(key)
            cursor.execute('''
                           SELECT encrypted_first_activate_time,
                                  encrypted_initial_remaining_seconds,
                                  encrypted_is_expired
                           FROM activation_records
                           WHERE encrypted_key = ?
                           ''', (encrypted_key,))
            result = cursor.fetchone()
            if not result:
                return None, False

            # 解密数据
            first_time_str = self.crypto.decrypt(result[0])
            initial_seconds = int(self.crypto.decrypt(result[1]))
            is_expired = self.crypto.decrypt(result[2]).lower() == "true"
            first_time = datetime.fromisoformat(first_time_str).astimezone()

            # 二次校验过期状态（防数据库篡改）
            now = get_local_time()
            elapsed = (now - first_time).total_seconds()
            actual_expired = elapsed >= initial_seconds
            if actual_expired != is_expired:
                self._update_expired_status(key, actual_expired)
                is_expired = actual_expired
                logger.warning(f"卡密{key[:4]}****状态修正为: {'过期' if is_expired else '未过期'}")

            activation_info = {
                "first_time": first_time,
                "initial_seconds": initial_seconds
            }
            return activation_info, is_expired
        except Exception as e:
            logger.error(f"获取卡密状态失败: {e}")
            return None, False
        finally:
            conn.close() if 'conn' in locals() else None

    def _update_expired_status(self, key: str, is_expired: bool):
        """更新卡密过期状态"""
        try:
            conn = sqlite3.connect(CARD_DB_PATH)
            cursor = conn.cursor()
            encrypted_key = self.crypto.encrypt(key)
            encrypted_is_expired = self.crypto.encrypt(str(is_expired))
            cursor.execute('''
                           UPDATE activation_records
                           SET encrypted_is_expired = ?,
                               last_check_time      = ?
                           WHERE encrypted_key = ?
                           ''', (encrypted_is_expired, get_local_time().isoformat(), encrypted_key))
            conn.commit()
        except Exception as e:
            logger.error(f"更新过期状态失败: {e}")
        finally:
            conn.close() if 'conn' in locals() else None

    def activate_card(self, key: str, activate_time: datetime) -> bool:
        """激活卡密（仅首次有效）"""
        if self.check_card_activated(key):
            logger.error(f"卡密{key[:4]}****已激活，无法重复激活")
            return False

        try:
            # 计算有效期
            if TEST_MODE:
                valid_seconds = TEST_VALID_SECONDS
            else:
                valid_seconds = PREDEFINED_CARDS[key] * 86400

            conn = sqlite3.connect(CARD_DB_PATH)
            cursor = conn.cursor()
            # 加密所有字段（防手动修改）
            encrypted_key = self.crypto.encrypt(key)
            encrypted_device = self.crypto.encrypt(self.device_id)
            encrypted_time = self.crypto.encrypt(activate_time.isoformat())
            encrypted_seconds = self.crypto.encrypt(str(valid_seconds))
            encrypted_expired = self.crypto.encrypt("False")

            cursor.execute('''
                           INSERT INTO activation_records (encrypted_key, encrypted_device_id,
                                                           encrypted_first_activate_time,
                                                           encrypted_initial_remaining_seconds, encrypted_is_expired,
                                                           last_check_time)
                           VALUES (?, ?, ?, ?, ?, ?)
                           ''', (encrypted_key, encrypted_device, encrypted_time, encrypted_seconds,
                                 encrypted_expired, get_local_time().isoformat()))
            conn.commit()
            logger.info(f"卡密{key[:4]}****激活成功，有效期{valid_seconds}秒")
            return True
        except sqlite3.IntegrityError:
            logger.error(f"卡密{key[:4]}****重复激活（数据库约束）")
            return False
        except Exception as e:
            logger.error(f"激活卡密失败: {e}")
            return False
        finally:
            conn.close() if 'conn' in locals() else None


# ===================== 4. 倒计时线程 =====================
class TimeCalcThread(QThread):
    time_updated = Signal(str, str)
    time_expired = Signal(str)

    def __init__(self, key: str, first_time: datetime, initial_seconds: int):
        super().__init__()
        self.key = key
        self.first_time = first_time
        self.initial_seconds = initial_seconds
        self.running = True

    def run(self):
        while self.running:
            now = get_local_time()
            # 计算剩余时间
            elapsed = (now - self.first_time).total_seconds()
            remaining = self.initial_seconds - elapsed

            if remaining <= 0:
                remaining = 0
                self.running = False
                self.time_expired.emit(self.key)
                # 更新数据库过期状态
                db = ActivationDatabase()
                db._update_expired_status(self.key, True)
                db.clear_last_input_key()

            # 格式化剩余时间
            days = int(remaining // 86400)
            hours = int((remaining % 86400) // 3600)
            minutes = int((remaining % 3600) // 60)
            seconds = int(remaining % 60)
            remaining_str = f"{days:02d}天{hours:02d}时{minutes:02d}分{seconds:02d}秒"
            end_time = self.first_time + timedelta(seconds=self.initial_seconds)
            end_time_str = end_time.strftime("%Y-%m-%d %H:%M:%S")

            self.time_updated.emit(remaining_str, end_time_str)
            time.sleep(1)

    def stop(self):
        self.running = False


# ===================== 5. 验证窗口 =====================
class CardVerifyDialog(QDialog):
    verify_success = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.db = ActivationDatabase()
        self.last_key = self.db.get_last_input_key()
        self.current_key = ""
        self.active_info = None
        self.time_thread = None
        self._init_ui()

    def _init_ui(self):
        """初始化界面"""
        self.setWindowTitle("卡密验证")
        self.setFixedSize(320, 280)
        self.setModal(True)
        # 窗口居中
        screen = QApplication.desktop().availableGeometry()
        self.move(screen.center() - self.frameGeometry().center())

        # 样式美化
        self.setStyleSheet('''
            QDialog {
                background-color: #000000;
                border: 2px solid #00ffcc;
                border-radius: 8px;
                background-image: 
                    linear-gradient(rgba(0,255,204,0.05) 1px, transparent 1px),
                    linear-gradient(90deg, rgba(0,255,204,0.05) 1px, transparent 1px);
                background-size: 20px 20px;
            }
            QLabel#title {
                color: #00ffcc;
                font: bold 18px Microsoft YaHei;
                text-align: center;
                text-shadow: 0 0 6px #00ffcc;
            }
            QLabel#info {
                color: #00ffcc;
                font: 12px Microsoft YaHei;
                text-align: center;
                margin: 2px 0;
            }
            QLineEdit {
                background-color: #121218;
                border: 1px solid #00ffcc;
                color: #00ffcc;
                padding: 8px 10px;
                border-radius: 4px;
                font: 12px Microsoft YaHei;
                text-align: center;
            }
            QLineEdit:focus {
                border-color: #00ffff;
                box-shadow: 0 0 8px rgba(0,255,255,0.6);
                outline: none;
            }
            QPushButton {
                background-color: #00ffcc;
                color: #000000;
                font: bold 13px Microsoft YaHei;
                padding: 8px 0;
                border-radius: 4px;
                box-shadow: 0 0 8px rgba(0,255,204,0.4);
            }
            QPushButton:hover {
                background-color: #00ffff;
                box-shadow: 0 0 12px rgba(0,255,255,0.6);
            }
            QPushButton:pressed {
                background-color: #00e0c0;
            }
            QFrame {
                border-top: 1px dashed #00ffcc30;
                margin: 8px 0;
            }
        ''')

        # 布局
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        # 标题
        title_label = QLabel("卡密验证")
        title_label.setObjectName("title")
        layout.addWidget(title_label)

        # 卡密输入框
        self.key_input = QLineEdit()
        self.key_input.setPlaceholderText("输入16位卡密（区分大小写）")
        self.key_input.setMaxLength(KEY_LENGTH)
        if self.last_key:
            self.key_input.setText(self.last_key)
        layout.addWidget(self.key_input)

        # 验证按钮
        self.verify_btn = QPushButton("验证激活")
        self.verify_btn.clicked.connect(self._verify_card)
        layout.addWidget(self.verify_btn)

        # 分隔线
        layout.addWidget(QFrame())

        # 剩余时间显示
        self.remaining_label = QLabel("剩余时间：--")
        self.remaining_label.setObjectName("info")
        layout.addWidget(self.remaining_label)

        # 结束时间显示
        self.end_label = QLabel("结束时间：--")
        self.end_label.setObjectName("info")
        layout.addWidget(self.end_label)

        self.setLayout(layout)

        # 自动验证上次输入的卡密
        if self.last_key:
            self._auto_verify_last_key()

    def _auto_verify_last_key(self):
        """自动验证上次输入的卡密"""
        key = self.last_key.strip()
        if len(key) != KEY_LENGTH:
            self.key_input.clear()
            self.db.clear_last_input_key()
            return
        self._check_card_status(key)

    def _verify_card(self):
        """手动验证卡密"""
        key = self.key_input.text().strip()
        # 校验长度
        if len(key) != KEY_LENGTH:
            QMessageBox.warning(self, "错误", f"卡密必须为{KEY_LENGTH}位！")
            self.key_input.selectAll()
            return
        # 校验卡密有效性
        if key not in PREDEFINED_CARDS:
            QMessageBox.critical(self, "错误", "卡密无效！")
            self.key_input.selectAll()
            return

        # 检查是否已激活
        if self.db.check_card_activated(key):
            self._check_card_status(key)
            return

        # 首次激活
        activate_time = get_local_time()
        if self.db.activate_card(key, activate_time):
            self.db.save_last_input_key(key)
            self.current_key = key
            self.active_info = {
                "first_time": activate_time,
                "initial_seconds": TEST_VALID_SECONDS if TEST_MODE else PREDEFINED_CARDS[key] * 86400
            }
            self._start_timer()
            QMessageBox.information(self, "成功",
                                    f"卡密激活成功！\n有效期：{PREDEFINED_CARDS[key]}天" if not TEST_MODE else f"测试模式，有效期{TEST_VALID_SECONDS}秒")
            self.verify_success.emit()
        else:
            QMessageBox.critical(self, "失败", "该卡密已被激活（仅可激活一次）！")

    def _check_card_status(self, key):
        """检查已激活卡密的状态"""
        self.current_key = key
        active_info, is_expired = self.db.get_card_status(key)
        if not active_info:
            QMessageBox.warning(self, "错误", "卡密未激活！")
            return
        if is_expired:
            QMessageBox.critical(self, "失效", "该卡密已过期失效！")
            self.key_input.clear()
            self.db.clear_last_input_key()
            return
        # 未过期，启动倒计时
        self.active_info = active_info
        self._start_timer()
        QMessageBox.information(self, "已激活", "该卡密已激活，剩余时间如下：")
        self.verify_success.emit()

    def _start_timer(self):
        """启动倒计时线程"""
        if self.time_thread and self.time_thread.isRunning():
            self.time_thread.stop()
            self.time_thread.wait()
        self.time_thread = TimeCalcThread(
            self.current_key,
            self.active_info["first_time"],
            self.active_info["initial_seconds"]
        )
        self.time_thread.time_updated.connect(self._update_time)
        self.time_thread.time_expired.connect(self._on_expired)
        self.time_thread.start()

    def _update_time(self, remaining_str, end_time_str):
        """更新剩余时间显示"""
        self.remaining_label.setText(f"剩余时间：{remaining_str}")
        self.end_label.setText(f"结束时间：{end_time_str}")

    def _on_expired(self, key):
        """卡密过期处理"""
        QMessageBox.critical(self, "失效", f"卡密{key[:4]}****已过期失效！")
        self.key_input.clear()
        self.db.clear_last_input_key()
        QApplication.instance().quit()

    def closeEvent(self, event):
        """窗口关闭时停止线程"""
        if self.time_thread and self.time_thread.isRunning():
            self.time_thread.stop()
            self.time_thread.wait()
        event.accept()


# ===================== 6. 启动函数（管理员权限检查） =====================
def is_admin():
    """检查是否拥有管理员权限"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def start_card_verification():
    """启动卡密验证流程"""
    # 检查管理员权限（确保数据库可写入）
    if not is_admin():
        logger.info("无管理员权限，尝试以管理员身份重启")
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
        sys.exit(0)

    # 启动Qt应用
    app = QApplication(sys.argv)
    dialog = CardVerifyDialog()
    verify_success = [False]

    def on_success():
        verify_success[0] = True
        dialog.close()

    dialog.verify_success.connect(on_success)
    dialog.exec_()
    return verify_success[0]


# ===================== 主程序入口 =====================
if __name__ == "__main__":
    # 创建日志目录
    Path("C:/ProgramData/Q/S").mkdir(parents=True, exist_ok=True)
    # 启动验证
    if start_card_verification():
        logger.info("卡密验证成功，启动主程序")
        print("卡密验证成功！")
        sys.exit(0)
    else:
        logger.info("卡密验证失败/取消")
        print("卡密验证失败或用户取消！")
        sys.exit(1)