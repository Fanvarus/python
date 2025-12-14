import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import configparser
from datetime import datetime
import glob
from tkinter.scrolledtext import ScrolledText


class PhoneNumberImageChecker:
    def __init__(self, root):
        self.root = root
        self.root.title("手机号图片匹配检查工具（增强版）")
        self.root.geometry("800x550")
        self.root.resizable(True, True)  # 允许窗口缩放

        # 配置文件路径
        self.config_path = os.path.join(os.path.expanduser("~"), ".phone_image_checker_plus.ini")

        # 初始化变量
        self.single_txt_path = tk.StringVar()  # 单个TXT文件路径
        self.multi_txt_paths = []  # 多个TXT文件路径列表
        self.folder_path = tk.StringVar()  # 图片文件夹路径
        self.output_folder_path = tk.StringVar()  # 输出文件夹路径

        # 加载上次保存的路径
        self.load_last_paths()

        # 创建界面
        self.create_ui()

    def create_ui(self):
        # 主框架（带内边距）
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. 单个TXT文件选择区域
        single_txt_frame = ttk.LabelFrame(main_frame, text="单个TXT文件", padding="10")
        single_txt_frame.pack(fill=tk.X, pady=6, ipady=5)

        # 单个文件输入框和按钮
        single_entry = ttk.Entry(single_txt_frame, textvariable=self.single_txt_path)
        single_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        single_btn = ttk.Button(single_txt_frame, text="选择", command=self.select_single_txt, width=10)
        single_btn.pack(side=tk.RIGHT, padx=5)

        # 2. 多个TXT文件选择区域
        multi_txt_frame = ttk.LabelFrame(main_frame, text="多个TXT文件（可选）", padding="10")
        multi_txt_frame.pack(fill=tk.X, pady=6, ipady=5)

        # 滚动文本框（调整高度）
        self.multi_txt_text = ScrolledText(multi_txt_frame, height=5, wrap=tk.WORD, font=("Arial", 9))
        self.multi_txt_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        # 多个文件操作按钮（垂直排列）
        multi_btn_frame = ttk.Frame(multi_txt_frame)
        multi_btn_frame.pack(side=tk.RIGHT, padx=5, fill=tk.Y)

        add_btn = ttk.Button(multi_btn_frame, text="添加文件", command=self.add_multi_txt, width=10)
        add_btn.pack(fill=tk.X, pady=3)
        clear_btn = ttk.Button(multi_btn_frame, text="清空列表", command=self.clear_multi_txt, width=10)
        clear_btn.pack(fill=tk.X, pady=3)

        # 3. 图片文件夹选择区域
        folder_frame = ttk.LabelFrame(main_frame, text="图片文件夹（含子文件夹）", padding="10")
        folder_frame.pack(fill=tk.X, pady=6, ipady=5)

        folder_entry = ttk.Entry(folder_frame, textvariable=self.folder_path)
        folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        folder_btn = ttk.Button(folder_frame, text="选择", command=self.select_folder, width=10)
        folder_btn.pack(side=tk.RIGHT, padx=5)

        # 4. 输出文件夹选择区域
        output_frame = ttk.LabelFrame(main_frame, text="结果输出文件夹", padding="10")
        output_frame.pack(fill=tk.X, pady=6, ipady=5)

        output_entry = ttk.Entry(output_frame, textvariable=self.output_folder_path)
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        output_btn = ttk.Button(output_frame, text="选择", command=self.select_output_folder, width=10)
        output_btn.pack(side=tk.RIGHT, padx=5)

        # 操作按钮区域（居中显示）
        button_frame = ttk.Frame(main_frame, padding="10")
        button_frame.pack(fill=tk.X, pady=8)

        # 居中容器
        btn_center_frame = ttk.Frame(button_frame)
        btn_center_frame.pack(anchor=tk.CENTER)

        self.check_button = ttk.Button(btn_center_frame, text="开始检查", command=self.start_check, width=15)
        self.check_button.pack(side=tk.LEFT, padx=8)
        clear_all_btn = ttk.Button(btn_center_frame, text="清空所有选择", command=self.clear_all, width=15)
        clear_all_btn.pack(side=tk.LEFT, padx=8)

        # 状态标签（居中显示，换行）
        self.status_label = ttk.Label(main_frame,
                                      text="就绪 - 请选择TXT文件、图片文件夹和输出文件夹（支持单个/多个TXT文件）",
                                      wraplength=700)
        self.status_label.pack(pady=5, anchor=tk.CENTER)

        # 配置样式（统一风格）
        self.root.style = ttk.Style()
        self.root.style.configure("TLabelFrame", font=("Arial", 10, "bold"))
        self.root.style.configure("TButton", font=("Arial", 9))
        self.root.style.configure("TEntry", font=("Arial", 9))
        self.root.style.configure("TLabel", font=("Arial", 9))

    # 以下方法保持不变，省略重复代码...
    def load_last_paths(self):
        config = configparser.ConfigParser()
        if os.path.exists(self.config_path):
            config.read(self.config_path)
            if "Paths" in config.sections():
                self.single_txt_path.set(config.get("Paths", "single_txt", fallback=""))
                self.folder_path.set(config.get("Paths", "image_folder", fallback=""))
                default_output = os.path.join(os.path.expanduser("~"), "Desktop")
                self.output_folder_path.set(config.get("Paths", "output_folder", fallback=default_output))

    def save_last_paths(self):
        config = configparser.ConfigParser()
        config["Paths"] = {
            "single_txt": self.single_txt_path.get(),
            "image_folder": self.folder_path.get(),
            "output_folder": self.output_folder_path.get()
        }
        with open(self.config_path, "w") as configfile:
            config.write(configfile)

    def select_single_txt(self):
        file_path = filedialog.askopenfilename(
            title="选择单个手机号TXT文件",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        if file_path:
            self.single_txt_path.set(file_path)
            self.save_last_paths()

    def add_multi_txt(self):
        file_paths = filedialog.askopenfilenames(
            title="选择多个手机号TXT文件",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        if file_paths:
            new_files = [path for path in file_paths if path not in self.multi_txt_paths]
            self.multi_txt_paths.extend(new_files)
            self.update_multi_txt_display()

    def clear_multi_txt(self):
        self.multi_txt_paths.clear()
        self.multi_txt_text.delete(1.0, tk.END)
        self.status_label.config(text="已清空多个TXT文件选择")

    def update_multi_txt_display(self):
        self.multi_txt_text.delete(1.0, tk.END)
        for i, path in enumerate(self.multi_txt_paths, 1):
            filename = os.path.basename(path)
            self.multi_txt_text.insert(tk.END, f"{i}. {filename}\n")

    def select_folder(self):
        folder_path = filedialog.askdirectory(title="选择图片文件夹（将递归检查子文件夹）")
        if folder_path:
            self.folder_path.set(folder_path)
            self.save_last_paths()

    def select_output_folder(self):
        folder_path = filedialog.askdirectory(title="选择结果输出文件夹")
        if folder_path:
            self.output_folder_path.set(folder_path)
            self.save_last_paths()

    def clear_all(self):
        self.single_txt_path.set("")
        self.multi_txt_paths.clear()
        self.multi_txt_text.delete(1.0, tk.END)
        self.folder_path.set("")
        self.save_last_paths()
        self.status_label.config(text="已清空所有选择 - 请重新选择")

    def read_single_txt(self):
        txt_path = self.single_txt_path.get()
        if not txt_path or not os.path.exists(txt_path):
            return []
        return self.get_phone_numbers_from_txt(txt_path)

    def read_multi_txt(self):
        all_phones = []
        for txt_path in self.multi_txt_paths:
            if os.path.exists(txt_path):
                phones = self.get_phone_numbers_from_txt(txt_path)
                if phones:
                    all_phones.extend(phones)
        return all_phones

    def get_phone_numbers_from_txt(self, txt_path):
        phone_numbers = []
        try:
            with open(txt_path, "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    phone = line.strip()
                    if phone:
                        phone_numbers.append(phone)
            return phone_numbers
        except Exception as e:
            messagebox.showerror("错误", f"读取文件 {os.path.basename(txt_path)} 失败：{str(e)}")
            return None

    def get_png_filenames(self, folder_path):
        png_filenames = set()
        try:
            png_files = glob.glob(os.path.join(folder_path, "**", "*.png"), recursive=True)
            png_files += glob.glob(os.path.join(folder_path, "**", "*.PNG"), recursive=True)
            for file in png_files:
                filename = os.path.splitext(os.path.basename(file))[0]
                png_filenames.add(filename.strip())
            return png_filenames
        except Exception as e:
            messagebox.showerror("错误", f"读取图片文件失败：{str(e)}")
            return None

    def validate_input(self):
        has_single_txt = bool(self.single_txt_path.get() and os.path.exists(self.single_txt_path.get()))
        has_multi_txt = bool(self.multi_txt_paths)
        if not has_single_txt and not has_multi_txt:
            messagebox.showwarning("警告", "请至少选择一个TXT文件（单个或多个）！")
            return False
        if not self.folder_path.get() or not os.path.isdir(self.folder_path.get()):
            messagebox.showwarning("警告", "请选择有效的图片文件夹！")
            return False
        if not self.output_folder_path.get() or not os.path.isdir(self.output_folder_path.get()):
            messagebox.showwarning("警告", "请选择有效的输出文件夹！")
            return False
        return True

    def start_check(self):
        if not self.validate_input():
            return
        self.check_button.config(state=tk.DISABLED)
        self.status_label.config(text="正在初始化检查...（递归扫描子文件夹）")
        self.root.update_idletasks()
        try:
            self.status_label.config(text="正在读取手机号...")
            self.root.update_idletasks()
            all_phones = []
            single_phones = self.read_single_txt()
            if single_phones is None:
                return
            all_phones.extend(single_phones)
            multi_phones = self.read_multi_txt()
            if multi_phones is None:
                return
            all_phones.extend(multi_phones)
            unique_phones = list(dict.fromkeys(all_phones))
            if not unique_phones:
                messagebox.showwarning("警告", "所有选中的TXT文件中未找到有效手机号！")
                return
            self.status_label.config(text="正在递归扫描图片文件...")
            self.root.update_idletasks()
            png_filenames = self.get_png_filenames(self.folder_path.get())
            if not png_filenames:
                messagebox.showwarning("警告", "文件夹及其子文件夹中未找到PNG图片！")
                return
            self.status_label.config(text="正在对比匹配...")
            self.root.update_idletasks()
            missing_phones = [phone for phone in unique_phones if phone not in png_filenames]
            self.status_label.config(text="正在生成结果文件...")
            self.root.update_idletasks()
            result_path = self.generate_result_file(missing_phones, unique_phones, png_filenames)
            result_msg = f"检查完成！\n"
            result_msg += f"总TXT文件数：{1 + len(self.multi_txt_paths)}\n"
            result_msg += f"去重后总手机号数：{len(unique_phones)}\n"
            result_msg += f"找到的PNG图片数：{len(png_filenames)}\n"
            result_msg += f"缺少对应图片的手机号数：{len(missing_phones)}\n"
            result_msg += f"结果文件已保存至：\n{result_path}"
            messagebox.showinfo("检查完成", result_msg)
            self.status_label.config(text=f"检查完成 - 缺少{len(missing_phones)}个手机号的图片（已扫描子文件夹）")
        except Exception as e:
            messagebox.showerror("错误", f"检查过程中出现错误：{str(e)}")
            self.status_label.config(text="检查失败 - 请重试")
        finally:
            self.check_button.config(state=tk.NORMAL)

    def generate_result_file(self, missing_phones, all_phones, png_filenames):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"缺少对应图片的手机号_{timestamp}.txt"
        result_path = os.path.join(self.output_folder_path.get(), filename)
        with open(result_path, "w", encoding="utf-8") as f:
            f.write(f"=" * 80 + "\n")
            f.write(f"手机号图片匹配检查结果\n")
            f.write(f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"搜索图片路径：{self.folder_path.get()}（含所有子文件夹）\n")
            f.write(f"选择的TXT文件数：{1 + len(self.multi_txt_paths)}\n")
            f.write(f"去重后总手机号数：{len(all_phones)}\n")
            f.write(f"找到的PNG图片数：{len(png_filenames)}\n")
            f.write(f"缺少对应图片的手机号数：{len(missing_phones)}\n")
            f.write(f"=" * 80 + "\n\n")
            if missing_phones:
                f.write("缺少对应图片的手机号列表：\n")
                f.write("-" * 50 + "\n")
                for i, phone in enumerate(missing_phones, 1):
                    f.write(f"{i:4d}. {phone}\n")
            else:
                f.write("恭喜！所有手机号都有对应的PNG图片文件！\n")
        return result_path


if __name__ == "__main__":
    root = tk.Tk()
    app = PhoneNumberImageChecker(root)
    root.mainloop()