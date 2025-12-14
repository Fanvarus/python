import requests
from bs4 import BeautifulSoup
import json
import time
import re
from datetime import datetime
import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter.scrolledtext import ScrolledText

# 全局配置
LATEST_ISSUES = 20  # 查询最近期数
TEMPLATE_FILENAME = "双色球投注模板.txt"  # TXT模板文件名
DEFAULT_FONT = ("微软雅黑", 10)
TITLE_FONT = ("微软雅黑", 14, "bold")
COLORS = {
    "primary": "#2E86AB",  # 主色调（蓝）
    "secondary": "#A23B72",  # 辅助色（紫）
    "success": "#F18F01",  # 成功色（橙）
    "warning": "#C73E1D",  # 警告色（红）
    "background": "#F8F9FA",  # 背景色（浅灰）
    "text": "#2D3436",  # 文本色（深灰）
    "select": "#D1E7DD"  # 选择背景色（浅绿）
}


class LotteryApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("双色球开奖查询工具")
        self.geometry("1200x800")
        self.minsize(1000, 700)
        self.configure(bg=COLORS["background"])

        # 全局变量
        self.user_bets = []  # 加载的投注方案
        self.lottery_results = []  # 获取的开奖数据
        self.winning_records = []  # 中奖记录
        self.total_prizes = []  # 总奖金

        # 初始化界面
        self._setup_style()
        self._create_widgets()
        self._layout_widgets()

        # 禁用初始状态下不可用的按钮
        self.btn_query.config(state=tk.DISABLED)
        self.btn_save.config(state=tk.DISABLED)

    def _setup_style(self):
        """设置界面样式"""
        self.style = ttk.Style()
        self.style.theme_use("clam")  # 基础主题

        # 按钮样式
        self.style.configure(
            "Primary.TButton",
            font=DEFAULT_FONT,
            background=COLORS["primary"],
            foreground="white",
            padding=(10, 5),
            borderwidth=0,
            relief=tk.FLAT
        )
        self.style.map(
            "Primary.TButton",
            background=[("active", COLORS["primary"] + "99")],
            foreground=[("active", "white")]
        )

        self.style.configure(
            "Secondary.TButton",
            font=DEFAULT_FONT,
            background=COLORS["secondary"],
            foreground="white",
            padding=(10, 5),
            borderwidth=0,
            relief=tk.FLAT
        )
        self.style.map(
            "Secondary.TButton",
            background=[("active", COLORS["secondary"] + "99")],
            foreground=[("active", "white")]
        )

        # Treeview样式（表格）
        self.style.configure(
            "Lottery.Treeview",
            font=DEFAULT_FONT,
            rowheight=25,
            fieldbackground=COLORS["background"],
            background=COLORS["background"],
            foreground=COLORS["text"]
        )
        # 设置选中行的背景色
        self.style.map(
            "Lottery.Treeview",
            background=[("selected", COLORS["select"])],
            foreground=[("selected", COLORS["text"])]
        )
        self.style.configure(
            "Lottery.Treeview.Heading",
            font=("微软雅黑", 10, "bold"),
            background=COLORS["primary"],
            foreground="white",
            padding=(5, 0)
        )
        self.style.map(
            "Lottery.Treeview.Heading",
            background=[("active", COLORS["primary"] + "99")]
        )

        # 标签样式
        self.style.configure(
            "Title.TLabel",
            font=TITLE_FONT,
            foreground=COLORS["primary"],
            background=COLORS["background"]
        )
        self.style.configure(
            "Info.TLabel",
            font=DEFAULT_FONT,
            foreground=COLORS["text"],
            background=COLORS["background"]
        )
        self.style.configure(
            "Warning.TLabel",
            font=DEFAULT_FONT,
            foreground=COLORS["warning"],
            background=COLORS["background"]
        )

    def _create_widgets(self):
        """创建所有界面控件"""
        # 1. 标题区域
        self.title_frame = ttk.Frame(self, style="Info.TLabel")
        self.lbl_main_title = ttk.Label(
            self.title_frame,
            text="双色球开奖查询工具",
            style="Title.TLabel"
        )
        self.lbl_sub_title = ttk.Label(
            self.title_frame,
            text="支持模板生成、多源查询、中奖分析",
            style="Info.TLabel"
        )

        # 2. 功能按钮区域
        self.btn_frame = ttk.Frame(self, style="Info.TLabel")
        self.btn_generate = ttk.Button(
            self.btn_frame,
            text="生成投注模板",
            style="Primary.TButton",
            command=self.generate_bet_template
        )
        self.btn_load = ttk.Button(
            self.btn_frame,
            text="加载投注模板",
            style="Primary.TButton",
            command=self.load_bet_template
        )
        self.btn_query = ttk.Button(
            self.btn_frame,
            text="查询开奖结果",
            style="Secondary.TButton",
            command=self.start_query_thread
        )
        self.btn_save = ttk.Button(
            self.btn_frame,
            text="保存查询结果",
            style="Secondary.TButton",
            command=self.save_winning_details
        )

        # 3. 投注方案展示区域
        self.bet_frame = ttk.LabelFrame(self, text="我的投注方案", style="Info.TLabel")
        # 移除selectbackground参数，通过样式设置选择背景
        self.tree_bets = ttk.Treeview(
            self.bet_frame,
            style="Lottery.Treeview",
            columns=("name", "red", "blue", "multiple"),
            show="headings"
        )
        # 设置投注方案表格列
        self.tree_bets.heading("name", text="方案名称", anchor=tk.CENTER)
        self.tree_bets.heading("red", text="红球", anchor=tk.CENTER)
        self.tree_bets.heading("blue", text="蓝球", anchor=tk.CENTER)
        self.tree_bets.heading("multiple", text="投注倍数", anchor=tk.CENTER)
        self.tree_bets.column("name", width=200, anchor=tk.CENTER)
        self.tree_bets.column("red", width=300, anchor=tk.CENTER)
        self.tree_bets.column("blue", width=80, anchor=tk.CENTER)
        self.tree_bets.column("multiple", width=100, anchor=tk.CENTER)
        # 投注方案滚动条
        self.scroll_bets = ttk.Scrollbar(
            self.bet_frame,
            orient=tk.VERTICAL,
            command=self.tree_bets.yview
        )
        self.tree_bets.configure(yscrollcommand=self.scroll_bets.set)

        # 4. 开奖结果展示区域
        self.result_frame = ttk.LabelFrame(self, text="开奖结果与中奖情况", style="Info.TLabel")
        # 移除selectbackground参数，通过样式设置选择背景
        self.tree_results = ttk.Treeview(
            self.result_frame,
            style="Lottery.Treeview",
            columns=("issue", "date", "time", "numbers", "prize"),
            show="headings"
        )
        # 设置开奖结果表格列
        self.tree_results.heading("issue", text="期号", anchor=tk.CENTER)
        self.tree_results.heading("date", text="开奖日期", anchor=tk.CENTER)
        self.tree_results.heading("time", text="时间", anchor=tk.CENTER)
        self.tree_results.heading("numbers", text="开奖号码", anchor=tk.CENTER)
        self.tree_results.heading("prize", text="中奖情况", anchor=tk.CENTER)
        self.tree_results.column("issue", width=120, anchor=tk.CENTER)
        self.tree_results.column("date", width=120, anchor=tk.CENTER)
        self.tree_results.column("time", width=80, anchor=tk.CENTER)
        self.tree_results.column("numbers", width=300, anchor=tk.CENTER)
        self.tree_results.column("prize", width=300, anchor=tk.CENTER)
        # 开奖结果滚动条
        self.scroll_results = ttk.Scrollbar(
            self.result_frame,
            orient=tk.VERTICAL,
            command=self.tree_results.yview
        )
        self.tree_results.configure(yscrollcommand=self.scroll_results.set)

        # 5. 中奖汇总区域
        self.summary_frame = ttk.LabelFrame(self, text="中奖汇总", style="Info.TLabel")
        self.txt_summary = ScrolledText(
            self.summary_frame,
            font=DEFAULT_FONT,
            wrap=tk.WORD,
            state=tk.DISABLED,
            background="white",
            foreground=COLORS["text"],
            relief=tk.FLAT,
            borderwidth=1
        )

        # 6. 状态提示区域
        self.status_frame = ttk.Frame(self, style="Info.TLabel")
        self.lbl_status = ttk.Label(
            self.status_frame,
            text="就绪：请先生成或加载投注模板",
            style="Info.TLabel"
        )

    def _layout_widgets(self):
        """布局所有控件（使用grid实现灵活排版）"""
        # 标题区域
        self.title_frame.grid(row=0, column=0, columnspan=4, padx=20, pady=(20, 10), sticky="w")
        self.lbl_main_title.grid(row=0, column=0, sticky="w")
        self.lbl_sub_title.grid(row=1, column=0, sticky="w")

        # 功能按钮区域
        self.btn_frame.grid(row=1, column=0, columnspan=4, padx=20, pady=(10, 20), sticky="we")
        self.btn_generate.grid(row=0, column=0, padx=(0, 10), sticky="w")
        self.btn_load.grid(row=0, column=1, padx=(0, 10), sticky="w")
        self.btn_query.grid(row=0, column=2, padx=(0, 10), sticky="w")
        self.btn_save.grid(row=0, column=3, padx=(0, 10), sticky="w")
        # 按钮区域右对齐填充
        self.btn_frame.grid_columnconfigure(4, weight=1)

        # 投注方案区域
        self.bet_frame.grid(row=2, column=0, columnspan=4, padx=20, pady=(0, 10), sticky="nsew")
        self.tree_bets.grid(row=0, column=0, sticky="nsew")
        self.scroll_bets.grid(row=0, column=1, sticky="ns")
        self.bet_frame.grid_rowconfigure(0, weight=1)
        self.bet_frame.grid_columnconfigure(0, weight=1)

        # 开奖结果区域
        self.result_frame.grid(row=3, column=0, columnspan=4, padx=20, pady=(0, 10), sticky="nsew")
        self.tree_results.grid(row=0, column=0, sticky="nsew")
        self.scroll_results.grid(row=0, column=1, sticky="ns")
        self.result_frame.grid_rowconfigure(0, weight=1)
        self.result_frame.grid_columnconfigure(0, weight=1)

        # 中奖汇总区域
        self.summary_frame.grid(row=4, column=0, columnspan=4, padx=20, pady=(0, 10), sticky="nsew")
        self.txt_summary.grid(row=0, column=0, sticky="nsew")
        self.summary_frame.grid_rowconfigure(0, weight=1)
        self.summary_frame.grid_columnconfigure(0, weight=1)

        # 状态提示区域
        self.status_frame.grid(row=5, column=0, columnspan=4, padx=20, pady=(10, 20), sticky="we")
        self.lbl_status.grid(row=0, column=0, sticky="w")

        # 全局行权重（实现自适应高度）
        self.grid_rowconfigure(2, weight=1)
        self.grid_rowconfigure(3, weight=2)
        self.grid_rowconfigure(4, weight=1)
        self.grid_columnconfigure(0, weight=1)

    # ------------------------------ 核心功能函数 ------------------------------
    def generate_bet_template(self):
        """生成投注模板到桌面"""
        try:
            # 获取桌面路径
            desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
            template_path = os.path.join(desktop_path, TEMPLATE_FILENAME)

            # 模板内容
            template_content = """
# 双色球投注模板（TXT版）
# 编辑说明：
# 1. 每行代表1个投注方案，空行和以"#"开头的行会被忽略
# 2. 方案格式：方案名称,红球1,红球2,红球3,红球4,红球5,红球6,蓝球,投注倍数
# 3. 格式要求：
#    - 红球：6个1-33的不重复整数（用英文逗号分隔）
#    - 蓝球：1个1-16的整数
#    - 倍数：正整数（≥1，代表投注倍数）
#    - 名称：可自定义（不包含英文逗号）
# 4. 示例如下（可直接修改或复制新增方案）

# 方案示例1
我的守号方案,1,5,10,15,20,25,8,1

# 方案示例2
随机选号方案,2,6,11,16,21,26,12,2

# 新增方案请按照上述格式添加（示例：）
# 幸运方案,3,7,12,17,22,27,5,3
"""

            # 写入文件
            with open(template_path, 'w', encoding='utf-8') as f:
                f.write(template_content.lstrip())

            # 提示用户
            self.update_status(f"✅ 模板已生成至桌面：{template_path}", "info")
            messagebox.showinfo(
                "生成成功",
                f"投注模板已保存到桌面\n路径：{template_path}\n\n请编辑模板后重新加载！"
            )
        except Exception as e:
            self.update_status(f"❌ 生成模板失败：{str(e)}", "warning")
            messagebox.showerror("生成失败", f"模板生成出错：{str(e)}")

    def load_bet_template(self):
        """加载投注模板（支持手动选择文件）"""
        # 打开文件选择对话框
        file_path = filedialog.askopenfilename(
            title="选择投注模板",
            filetypes=[("TXT文件", "*.txt"), ("所有文件", "*.*")],
            initialdir=os.path.join(os.path.expanduser('~'), 'Desktop'),
            initialfile=TEMPLATE_FILENAME
        )

        if not file_path:
            return  # 用户取消选择

        try:
            valid_bets = []
            line_num = 0

            # 读取并解析模板
            with open(file_path, 'r', encoding='utf-8') as f:
                for line in f:
                    line_num += 1
                    stripped_line = line.strip()
                    if not stripped_line or stripped_line.startswith('#'):
                        continue

                    # 验证格式
                    parts = stripped_line.split(',')
                    if len(parts) != 9:
                        raise ValueError(f"第{line_num}行：需包含9个部分（当前{len(parts)}个）")

                    # 提取字段
                    scheme_name = parts[0].strip()
                    if not scheme_name:
                        raise ValueError(f"第{line_num}行：方案名称不能为空")

                    # 验证红球
                    red_balls = []
                    for i, part in enumerate(parts[1:7], 1):
                        try:
                            red_num = int(part.strip())
                        except ValueError:
                            raise ValueError(f"第{line_num}行：第{i}个红球不是整数")
                        if red_num < 1 or red_num > 33:
                            raise ValueError(f"第{line_num}行：第{i}个红球超出1-33范围")
                        red_balls.append(red_num)
                    if len(set(red_balls)) != 6:
                        raise ValueError(f"第{line_num}行：红球存在重复数字")

                    # 验证蓝球
                    try:
                        blue_ball = int(parts[7].strip())
                    except ValueError:
                        raise ValueError(f"第{line_num}行：蓝球不是整数")
                    if blue_ball < 1 or blue_ball > 16:
                        raise ValueError(f"第{line_num}行：蓝球超出1-16范围")

                    # 验证倍数
                    try:
                        multiple = int(parts[8].strip())
                    except ValueError:
                        raise ValueError(f"第{line_num}行：倍数不是整数")
                    if multiple < 1:
                        raise ValueError(f"第{line_num}行：倍数需≥1")

                    valid_bets.append({
                        'name': scheme_name,
                        'red': red_balls,
                        'blue': blue_ball,
                        'multiple': multiple
                    })

            if not valid_bets:
                raise ValueError("模板中无有效投注方案")

            # 更新全局变量和界面
            self.user_bets = valid_bets
            self.update_bet_tree()
            self.update_status(f"✅ 成功加载{len(valid_bets)}个投注方案", "info")
            messagebox.showinfo("加载成功", f"共加载{len(valid_bets)}个投注方案")

            # 启用查询按钮
            self.btn_query.config(state=tk.NORMAL)

        except Exception as e:
            self.update_status(f"❌ 加载模板失败：{str(e)}", "warning")
            messagebox.showerror("加载失败", f"模板解析出错：{str(e)}")

    def start_query_thread(self):
        """启动查询线程（避免界面卡住）"""
        # 禁用按钮防止重复查询
        self.btn_query.config(state=tk.DISABLED)
        self.update_status("🔍 正在获取开奖数据...（请稍候）", "info")

        # 启动子线程执行查询
        query_thread = threading.Thread(target=self.fetch_and_analyze, daemon=True)
        query_thread.start()

    def fetch_and_analyze(self):
        """获取开奖数据并分析中奖情况（子线程执行）"""
        try:
            # 1. 获取开奖数据
            self.lottery_results = self.fetch_lottery_results()

            # 2. 分析中奖情况
            self.total_prizes, self.winning_records, _ = self.analyze_winning()

            # 3. 更新界面（需回到主线程）
            self.after(0, self.update_result_interface)

        except Exception as e:
            # 异常处理（回到主线程更新界面）
            self.after(0, lambda: self.handle_query_error(str(e)))

    def fetch_lottery_results(self):
        """多源获取彩票结果（复用原有逻辑）"""
        sources = [
            ("中国福彩网API", self.fetch_cwl_gov_results),
            ("500彩票网爬虫", self.fetch_500_data),
            ("网易彩票API", self.fetch_netease_data),
            ("千彩网API", self.fetch_296o_data)
        ]

        for source_name, fetch_func in sources:
            try:
                self.after(0, lambda s=source_name: self.update_status(f"🔍 尝试从{s}获取数据...", "info"))
                results = fetch_func()
                if results:
                    sorted_results = sorted(results, key=lambda x: x["issue"], reverse=True)
                    return sorted_results
            except Exception as e:
                self.after(0, lambda s=source_name, err=str(e): self.update_status(f"❌ {s}获取失败：{err[:20]}...",
                                                                                   "warning"))

        raise Exception("所有数据源均不可用，请检查网络或稍后重试")

    def analyze_winning(self):
        """分析中奖情况（复用原有逻辑）"""
        valid_bets = self.user_bets
        results = self.lottery_results

        total_prizes = []
        winning_records = []

        for bet in valid_bets:
            bet_total = 0
            for res in results:
                level, prize_str, prize_val = self.check_prize(
                    bet["red"], bet["blue"], res["red"], res["blue"]
                )
                total = prize_val * bet["multiple"] if level != "未中奖" else 0
                bet_total += total

                if level != "未中奖":
                    red_str = " ".join(f"{n:02d}" for n in res["red"])
                    numbers_str = f"{red_str} + {res['blue']:02d}"
                    winning_records.append({
                        "issue": res["issue"],
                        "date": res["date"],
                        "time": res["time"],
                        "scheme": bet["name"],
                        "red": bet["red"],
                        "blue": bet["blue"],
                        "multiple": bet["multiple"],
                        "prize": total,
                        "level": level,
                        "winning_numbers": numbers_str
                    })
            total_prizes.append(bet_total)

        return total_prizes, winning_records, results

    # ------------------------------ 数据获取函数（复用原有逻辑） ------------------------------
    def fetch_cwl_gov_results(self):
        try:
            url = "http://www.cwl.gov.cn/cwl_admin/front/cwlkj/searchKjxx/findDrawNotice"
            params = {"name": "ssq", "issueCount": LATEST_ISSUES, "issueStart": "", "issueEnd": "", "dayStart": "",
                      "dayEnd": ""}
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
                "Referer": "http://www.cwl.gov.cn/kjxx/ssq/"}
            response = requests.get(url, params=params, headers=headers, timeout=10)
            data = response.json()

            if data.get("state") == 0:
                results = []
                for item in data["result"]:
                    red_balls = list(map(int, item["red"].split(",")))
                    if len(red_balls) != 6 or any(n < 1 or n > 33 for n in red_balls):
                        continue
                    blue_ball = int(item["blue"])
                    if blue_ball < 1 or blue_ball > 16:
                        continue
                    open_time = item["date"][11:16] if " " in item["date"] else "21:15"
                    if open_time == "00:00":
                        open_time = "21:15"
                    results.append({
                        "issue": item["code"],
                        "date": item["date"][:10],
                        "time": open_time,
                        "red": red_balls,
                        "blue": blue_ball
                    })
                return results[:LATEST_ISSUES]
        except Exception:
            pass
        return None

    def fetch_500_data(self):
        try:
            url = "https://datachart.500.com/ssq/history/newinc/history.php?start=00001&end=99999"
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"}
            response = requests.get(url, headers=headers, timeout=10)
            response.encoding = 'gbk'
            soup = BeautifulSoup(response.text, 'html.parser')
            table = soup.find('tbody', {'id': 'tdata'})
            if not table:
                return None

            results = []
            for row in table.find_all('tr')[:LATEST_ISSUES]:
                cols = row.find_all('td')
                if len(cols) < 16:
                    continue
                issue = cols[0].text.strip()
                date = cols[15].text.strip()
                red_balls = [int(cols[i].text.strip()) for i in range(1, 7)]
                blue_ball = int(cols[7].text.strip())
                if all(1 <= n <= 33 for n in red_balls) and 1 <= blue_ball <= 16:
                    dt = datetime.strptime(date, "%Y-%m-%d")
                    if dt.weekday() in [1, 3, 6]:
                        results.append({
                            "issue": issue,
                            "date": date,
                            "time": "21:15",
                            "red": red_balls,
                            "blue": blue_ball
                        })
            return results
        except Exception:
            return None

    def fetch_netease_data(self):
        try:
            url = "https://cailele.tech/lottery/ssq"
            params = {"limit": LATEST_ISSUES}
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"}
            response = requests.get(url, params=params, headers=headers, timeout=8)
            data = response.json()

            if "result" in data:
                results = []
                for item in data["result"]:
                    numbers = item["lottery_res"].split("|")
                    if len(numbers) != 2:
                        continue
                    red_balls = list(map(int, numbers[0].split(",")))
                    blue_ball = int(numbers[1])
                    if len(red_balls) == 6 and all(1 <= n <= 33 for n in red_balls) and 1 <= blue_ball <= 16:
                        dt = datetime.strptime(item["lottery_date"], "%Y-%m-%d")
                        if dt.weekday() in [1, 3, 6]:
                            results.append({
                                "issue": item["lottery_no"],
                                "date": item["lottery_date"],
                                "time": "21:15",
                                "red": red_balls,
                                "blue": blue_ball
                            })
                return results[:LATEST_ISSUES]
        except Exception:
            return None

    def fetch_296o_data(self):
        try:
            url = "https://api.296o.com/api"
            params = {"code": "ssq", "rows": LATEST_ISSUES, "format": "json"}
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"}
            response = requests.get(url, params=params, headers=headers, timeout=8)
            data = response.json()

            if "data" in data:
                results = []
                for item in data["data"]:
                    numbers = item["opencode"].split("|")
                    if len(numbers) != 2:
                        continue
                    red_balls = list(map(int, numbers[0].split(",")))
                    blue_ball = int(numbers[1])
                    if len(red_balls) == 6 and all(1 <= n <= 33 for n in red_balls) and 1 <= blue_ball <= 16:
                        dt = datetime.strptime(item["opentime"][:10], "%Y-%m-%d")
                        if dt.weekday() in [1, 3, 6]:
                            results.append({
                                "issue": item["expect"].replace("-", ""),
                                "date": item["opentime"][:10],
                                "time": "21:15",
                                "red": red_balls,
                                "blue": blue_ball
                            })
                return results[:LATEST_ISSUES]
        except Exception:
            return None

    def check_prize(self, user_red, user_blue, prize_red, prize_blue):
        """判断中奖等级（复用原有逻辑）"""
        red_match = len(set(user_red) & set(prize_red))
        blue_match = user_blue == prize_blue

        if red_match == 6 and blue_match:
            return "一等奖", "浮动(最高1000万)", 0
        elif red_match == 6:
            return "二等奖", "浮动", 0
        elif red_match == 5 and blue_match:
            return "三等奖", "3000元", 3000
        elif red_match == 5 or (red_match == 4 and blue_match):
            return "四等奖", "200元", 200
        elif red_match == 4 or (red_match == 3 and blue_match):
            return "五等奖", "10元", 10
        elif blue_match:
            return "六等奖", "5元", 5
        return "未中奖", "0元", 0

    # ------------------------------ 界面更新函数 ------------------------------
    def update_bet_tree(self):
        """更新投注方案表格"""
        # 清空现有数据
        for item in self.tree_bets.get_children():
            self.tree_bets.delete(item)

        # 添加新数据
        for bet in self.user_bets:
            red_str = "、".join(map(str, bet["red"]))
            self.tree_bets.insert(
                "", tk.END,
                values=(bet["name"], red_str, bet["blue"], bet["multiple"])
            )

    def update_result_interface(self):
        """更新开奖结果和中奖汇总界面"""
        # 1. 更新开奖结果表格
        self.update_result_tree()

        # 2. 更新中奖汇总文本
        self.update_summary_text()

        # 3. 更新状态和按钮
        self.update_status(f"✅ 查询完成！共获取{len(self.lottery_results)}期数据", "info")
        self.btn_query.config(state=tk.NORMAL)
        self.btn_save.config(state=tk.NORMAL)

        # 4. 提示中奖情况
        total_all = sum(self.total_prizes)
        if total_all > 0:
            messagebox.showinfo(
                "查询完成",
                f"恭喜！您的投注方案共中奖{total_all}元\n\n详细情况请查看中奖汇总"
            )
        else:
            messagebox.showinfo("查询完成", "未查询到中奖记录，继续加油！")

    def update_result_tree(self):
        """更新开奖结果表格（只显示有中奖的期数）"""
        # 清空现有数据
        for item in self.tree_results.get_children():
            self.tree_results.delete(item)

        # 筛选有中奖的期数
        winning_issues = set(record["issue"] for record in self.winning_records)
        result_data = []

        for res in self.lottery_results:
            if res["issue"] not in winning_issues:
                continue  # 跳过无中奖的期数

            # 格式化开奖号码
            red_str = " ".join(f"{n:02d}" for n in res["red"])
            numbers_str = f"红球[{red_str}] + 蓝球{res['blue']:02d}"

            # 汇总该期所有方案的中奖情况
            prize_info = []
            for i, bet in enumerate(self.user_bets):
                level, _, _ = self.check_prize(bet["red"], bet["blue"], res["red"], res["blue"])
                if level != "未中奖":
                    prize_info.append(f"{bet['name']}：{level}")

            result_data.append({
                "issue": res["issue"],
                "date": res["date"],
                "time": res["time"],
                "numbers": numbers_str,
                "prize": " | ".join(prize_info) if prize_info else "未中奖"
            })

        # 添加数据到表格
        for data in result_data:
            self.tree_results.insert(
                "", tk.END,
                values=(data["issue"], data["date"], data["time"], data["numbers"], data["prize"])
            )

    def update_summary_text(self):
        """更新中奖汇总文本"""
        # 清空现有内容
        self.txt_summary.config(state=tk.NORMAL)
        self.txt_summary.delete(1.0, tk.END)

        # 构建汇总内容
        total_all = sum(self.total_prizes)
        content = f"=== 双色球中奖汇总报告 ===\n"
        content += f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        content += f"查询期数：{len(self.lottery_results)} 期\n"
        content += f"参与方案：{len(self.user_bets)} 个\n"
        content += f"总中奖金额：{total_all} 元\n"
        content += f"平均每期奖金：{total_all / len(self.lottery_results):.2f} 元\n\n"

        # 方案详情
        content += "=== 各方案中奖详情 ===\n"
        for i, (bet, total) in enumerate(zip(self.user_bets, self.total_prizes), 1):
            red_str = "、".join(map(str, bet["red"]))
            content += f"{i}. {bet['name']}\n"
            content += f"   投注：红球[{red_str}] + 蓝球{bet['blue']}（{bet['multiple']}倍）\n"
            content += f"   奖金：{total} 元\n\n"

        # 中奖记录（如有）
        if self.winning_records:
            content += "=== 详细中奖记录 ===\n"
            # 按期号分组
            issue_groups = {}
            for record in self.winning_records:
                if record["issue"] not in issue_groups:
                    issue_groups[record["issue"]] = []
                issue_groups[record["issue"]].append(record)

            for issue, records in sorted(issue_groups.items(), reverse=True):
                first = records[0]
                content += f"第{issue}期（{first['date']} {first['time']}）\n"
                content += f"   开奖号码：{first['winning_numbers']}\n"
                for idx, record in enumerate(records, 1):
                    content += f"   {idx}. {record['scheme']}：{record['level']}（{record['prize']}元）\n"
                content += "\n"
        else:
            content += "=== 详细中奖记录 ===\n"
            content += "   暂无中奖记录，继续加油！\n"

        # 兑奖须知
        content += "=== 兑奖须知 ===\n"
        content += "1. 中奖后需在开奖日起60天内到当地福彩站点兑奖\n"
        content += "2. 单注奖金1万元及以上需缴纳20%个人偶然所得税\n"
        content += "3. 兑奖唯一凭证为官方纸质彩票，本报告仅作参考\n"
        content += "4. 理性购彩，量力而行，享受娱乐属性\n"

        # 插入内容
        self.txt_summary.insert(1.0, content)
        self.txt_summary.config(state=tk.DISABLED)

    def update_status(self, text, status_type="info"):
        """更新状态提示"""
        if status_type == "warning":
            self.lbl_status.config(text=text, style="Warning.TLabel")
        else:
            self.lbl_status.config(text=text, style="Info.TLabel")
        self.update()  # 强制刷新界面

    def handle_query_error(self, error_msg):
        """处理查询错误"""
        self.update_status(f"❌ 查询失败：{error_msg}", "warning")
        self.btn_query.config(state=tk.NORMAL)
        messagebox.showerror("查询失败", f"获取开奖数据出错：{error_msg}")

    def save_winning_details(self):
        """保存查询结果到文件"""
        if not self.lottery_results or not self.user_bets:
            messagebox.showwarning("保存失败", "暂无查询结果可保存")
            return

        # 选择保存路径
        save_path = filedialog.asksaveasfilename(
            title="保存查询结果",
            filetypes=[("TXT文件", "*.txt"), ("所有文件", "*.*")],
            initialdir=os.path.join(os.path.expanduser('~'), 'Desktop'),
            initialfile=f"双色球开奖详情_{datetime.now().strftime('%Y%m%d%H%M%S')}.txt"
        )

        if not save_path:
            return

        try:
            # 构建保存内容
            content = "=" * 80 + "\n"
            content += "双色球开奖详情报告".center(80) + "\n"
            content += f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            content += "=" * 80 + "\n\n"

            # 最新开奖结果
            if self.lottery_results:
                latest = self.lottery_results[0]
                red_str = " ".join(f"{n:02d}" for n in latest["red"])
                content += "【最新一期开奖结果】\n"
                content += f"期号:    第{latest['issue']}期\n"
                content += f"日期:    {latest['date']}\n"
                content += f"时间:    {latest['time']}\n"
                content += f"开奖号码: 红球[{red_str}] + 蓝球{latest['blue']:02d}\n\n"
                content += "-" * 80 + "\n\n"

            # 投注方案
            content += "【您的投注方案】\n"
            for i, bet in enumerate(self.user_bets, 1):
                red_str = "、".join(map(str, bet["red"]))
                content += f"方案{i}：{bet['name']}\n"
                content += f"  红球：{red_str}\n"
                content += f"  蓝球：{bet['blue']}\n"
                content += f"  倍数：{bet['multiple']}倍\n"
                content += f"  总奖金：{self.total_prizes[i - 1]}元\n\n"
            content += "-" * 80 + "\n\n"

            # 中奖统计
            total_all = sum(self.total_prizes)
            content += "【中奖统计汇总】\n"
            content += f"参与方案数：{len(self.user_bets)} 个\n"
            content += f"查询期数：{len(self.lottery_results)} 期\n"
            content += f"总中奖金额：{total_all} 元\n"
            content += f"平均每期奖金：{total_all / len(self.lottery_results):.2f} 元\n\n"
            content += "-" * 80 + "\n\n"

            # 完整开奖记录
            content += "【最近20期开奖记录】\n"
            content += f"{'期号':<10} {'日期':<12} {'时间':<6} {'开奖号码':<25} {'各方案中奖情况'}\n"
            content += "-" * 100 + "\n"
            for res in self.lottery_results:
                red_str = " ".join(f"{n:02d}" for n in res["red"])
                numbers_str = f"红球[{red_str}] + 蓝球{res['blue']:02d}"
                scheme_results = []
                for bet in self.user_bets:
                    level, _, _ = self.check_prize(bet["red"], bet["blue"], res["red"], res["blue"])
                    scheme_results.append(f"{bet['name']}:{level}")
                content += f"{res['issue']:<10} {res['date']:<12} {res['time']:<6} {numbers_str:<25} {', '.join(scheme_results)}\n"
            content += "\n" + "-" * 80 + "\n\n"

            # 详细中奖记录
            content += "【详细中奖记录】\n"
            if self.winning_records:
                issue_groups = {}
                for record in self.winning_records:
                    if record["issue"] not in issue_groups:
                        issue_groups[record["issue"]] = []
                    issue_groups[record["issue"]].append(record)

                for issue, records in issue_groups.items():
                    first_record = records[0]
                    content += f"\n► 第{issue}期（{first_record['date']} {first_record['time']}）\n"
                    content += f"  开奖号码：{first_record['winning_numbers']}\n"
                    for idx, record in enumerate(records, 1):
                        content += f"  {idx}. {record['scheme']}\n"
                        content += f"     投注：红球{record['red']} + 蓝球{record['blue']}（{record['multiple']}倍）\n"
                        content += f"     奖项：{record['level']}，奖金{record['prize']}元\n"
            else:
                content += "  ⚠️  暂无中奖记录，继续加油！\n"

            # 兑奖须知
            content += "\n" + "=" * 80 + "\n"
            content += "【兑奖须知】\n"
            content += "  1. 中奖后需在开奖日起60天内到当地福利彩票销售站点或中心兑奖\n"
            content += "  2. 单注奖金1万元及以上需缴纳20%个人偶然所得税（由兑奖机构代扣）\n"
            content += "  3. 兑奖唯一凭证为官方纸质彩票，本电子报告仅作查询参考，不具备兑奖效力\n"
            content += "  4. 官方查询渠道：中国福利彩票网（www.cwl.gov.cn）、福彩官方APP\n"
            content += "  5. 理性购彩，量力而行，享受彩票的娱乐属性\n"
            content += "=" * 80

            # 保存文件
            with open(save_path, 'w', encoding='utf-8') as f:
                f.write(content)

            self.update_status(f"✅ 结果已保存至：{save_path}", "info")
            messagebox.showinfo("保存成功", f"查询结果已保存到：\n{save_path}")

        except Exception as e:
            self.update_status(f"❌ 保存失败：{str(e)}", "warning")
            messagebox.showerror("保存失败", f"文件保存出错：{str(e)}")


# ------------------------------ 程序入口 ------------------------------
def main():
    # 自动安装依赖（首次运行）
    try:
        import requests
        from bs4 import BeautifulSoup
    except ImportError:
        print("⚠️  检测到缺失依赖库，正在自动安装...")
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "requests", "beautifulsoup4"])
        print("✅ 依赖库安装完成，启动程序...")

    # 启动TK应用
    app = LotteryApp()
    app.mainloop()


if __name__ == "__main__":
    main()
