import os
import zipfile
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter.ttk import Progressbar
import threading
import logging
import patoolib
import shutil
import requests
import subprocess
import sys
import locale
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor
from logging.handlers import RotatingFileHandler
import json

# 设置日志配置
log_file = 'zip_extractor.log'
log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
log_handler = RotatingFileHandler(log_file, maxBytes=1 * 1024 * 1024, backupCount=2)  # 每个日志文件最大1MB，保留2个备份
log_handler.setFormatter(log_formatter)
logging.basicConfig(level=logging.INFO, handlers=[log_handler])

# 常量
SUPPORTED_FORMATS = ('.zip', '.rar', '.7z', '.tar', '.gz', '.bz2')
CURRENT_VERSION = "1.0"
UPDATE_URL = "https://example.com/latest_version.txt"
DOWNLOAD_URL = "https://example.com/downloads/尘飞批量解压_v{version}.exe"

# 保存和读取上次选择的目录
CONFIG_FILE = 'config.json'

def load_last_directory():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            config = json.load(f)
            return config.get('last_directory', '')
    return ''

def save_last_directory(directory):
    config = {'last_directory': directory}
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f)

class ZipExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("尘飞批量解压 V1.0")

        # 设置多语言支持，默认为中文
        self.language = "zh_CN"
        self.translations = {}  # 初始化为空字典，稍后设置语言时填充

        # 设置窗口初始大小和可调整大小
        self.root.geometry("600x500")  # 设置窗口初始大小为600x500像素
        self.root.resizable(True, True)  # 允许窗口调整大小

        # 创建并放置标签和按钮
        self.create_widgets()

        # 设置语言（确保界面元素已创建）
        self.set_language(self.language)

        # 加载上次选择的目录
        self.directory_path = load_last_directory()
        if self.directory_path and os.path.isdir(self.directory_path):
            self.status_label.config(text=f"{self.translations['choose_folder']} {self.directory_path}")
            self.extract_button.config(state=tk.NORMAL)

        # 启动时静默检查更新
        self.check_for_updates_on_startup(silent=True)

    def create_widgets(self):
        """创建所有界面元素"""
        self.label = tk.Label(self.root, text="", font=("Arial", 12))
        self.label.pack(pady=20)

        self.select_button = tk.Button(self.root, text="", command=self.select_directory, font=("Arial", 12))
        self.select_button.pack(pady=10)

        self.extract_button = tk.Button(self.root, text="", command=self.start_extraction, state=tk.DISABLED,
                                        font=("Arial", 12))
        self.extract_button.pack(pady=10)

        self.cancel_button = tk.Button(self.root, text="", command=self.cancel_extraction, state=tk.DISABLED,
                                       font=("Arial", 12))
        self.cancel_button.pack(pady=10)

        self.help_button = tk.Button(self.root, text="", command=self.show_help, font=("Arial", 12))
        self.help_button.pack(pady=10)

        self.update_button = tk.Button(self.root, text="", command=self.check_for_updates, font=("Arial", 12))
        self.update_button.pack(pady=10)

        self.status_label = tk.Label(self.root, text="", wraplength=500, justify="left", font=("Arial", 12))
        self.status_label.pack(pady=10)

        self.progress_bar = Progressbar(self.root, orient='horizontal', mode='determinate', length=500)
        self.progress_bar.pack(pady=10)

        self.time_label = tk.Label(self.root, text="", font=("Arial", 12))
        self.time_label.pack(pady=10)

        self.log_text = tk.Text(self.root, height=10, width=80, wrap='word')
        self.log_text.pack(pady=10)
        self.log_text.config(state=tk.DISABLED)

    def set_language(self, lang_code):
        """设置语言"""
        self.language = lang_code
        if self.language == "zh_CN":
            self.translations = {
                "choose_folder": "请选择包含压缩文件的文件夹：",
                "select_button": "选择文件夹",
                "extract_button": "开始解压",
                "cancel_button": "取消",
                "help_button": "使用说明",
                "update_button": "检查更新",
                "help_message": (
                    "使用说明:\n"
                    "1. 点击“选择文件夹”按钮，选择包含压缩文件的文件夹。\n"
                    "2. 点击“开始解压”按钮，程序将自动解压该文件夹及其所有子文件夹中的所有支持格式的压缩文件（如 .zip, .rar, .7z 等）。\n"
                    "3. 解压完成后，程序会弹出一个统计报告，显示解压文件的数量、总大小以及删除原文件后释放的空间。\n"
                    "4. 统计报告将保存在解压目录中，并自动打开解压后的文件夹供您查看。\n"
                    "5. 如果需要取消解压操作，可以点击“取消”按钮。\n"
                    "6. 您可以选择删除原始压缩文件以释放空间。\n"
                )
            }
        elif self.language == "en_US":
            self.translations = {
                "choose_folder": "Please select a folder containing compressed files:",
                "select_button": "Select Folder",
                "extract_button": "Start Extraction",
                "cancel_button": "Cancel",
                "help_button": "Help",
                "update_button": "Check for Updates",
                "help_message": (
                    "Usage Instructions:\n"
                    "1. Click the “Select Folder” button to choose a folder containing compressed files.\n"
                    "2. Click the “Start Extraction” button to automatically extract all supported formats of compressed files (.zip, .rar, .7z, etc.) in the folder and its subfolders.\n"
                    "3. After extraction is complete, a report will be generated showing the number of extracted files, total size, and space released by deleting original files.\n"
                    "4. The report will be saved in the extraction directory, and the extracted folder will be opened for you to view.\n"
                    "5. If you need to cancel the extraction process, click the “Cancel” button.\n"
                    "6. You can choose to delete the original compressed files to free up space.\n"
                )
            }
        self.update_ui_text()

    def update_ui_text(self):
        """更新界面文本"""
        self.label.config(text=self.translations["choose_folder"])
        self.select_button.config(text=self.translations["select_button"])
        self.extract_button.config(text=self.translations["extract_button"])
        self.cancel_button.config(text=self.translations["cancel_button"])
        self.help_button.config(text=self.translations["help_button"])
        self.update_button.config(text=self.translations["update_button"])

    def select_directory(self):
        """选择文件夹"""
        initial_dir = self.directory_path or os.getcwd()  # 默认为程序目录或上次选择的目录
        self.directory_path = filedialog.askdirectory(initialdir=initial_dir)
        if self.directory_path and os.path.isdir(self.directory_path):
            self.extract_button.config(state=tk.NORMAL)
            self.status_label.config(text=f"{self.translations['choose_folder']} {self.directory_path}")
            save_last_directory(self.directory_path)  # 保存当前选择的目录
        else:
            self.extract_button.config(state=tk.DISABLED)
            self.status_label.config(text=self.translations["choose_folder"])

    def start_extraction(self):
        """开始解压所有压缩文件"""
        if not self.directory_path:
            messagebox.showerror("错误", self.translations["choose_folder"])
            return

        try:
            self.compressed_files = self.find_compressed_files(self.directory_path)
            if not self.compressed_files:
                messagebox.showinfo("提示", "所选文件夹及其子文件夹中没有找到支持的压缩文件。")
                return

            self.total_files = len(self.compressed_files)
            self.current_file_index = 0
            self.stop_flag = False
            self.extracted_files_size = 0
            self.deleted_files_size = 0
            self.progress_bar['value'] = 0
            self.status_label.config(text="正在准备解压...")
            self.start_time = datetime.now()  # 记录开始时间
            self.update_time_label()

            self.extract_button.config(state=tk.DISABLED)
            self.cancel_button.config(state=tk.NORMAL)

            self.extraction_thread = threading.Thread(target=self.extract_files)
            self.extraction_thread.start()

        except Exception as e:
            logging.error(f"准备解压时出现错误: {e}")
            messagebox.showerror("错误", f"准备解压时出现错误: {e}")

    def find_compressed_files(self, directory):
        """递归查找目录及其子目录中的压缩文件"""
        compressed_files = []
        for root, dirs, files in os.walk(directory):
            for file in files:
                if file.lower().endswith(SUPPORTED_FORMATS):
                    compressed_files.append(os.path.join(root, file))
        return compressed_files

    def extract_files(self):
        """解压所有压缩文件（后台线程）"""
        success_count = 0
        failed_files = []

        with ThreadPoolExecutor() as executor:
            futures = []
            for compressed_file in self.compressed_files:
                if self.stop_flag:
                    break

                future = executor.submit(self.extract_single_file, compressed_file)
                futures.append(future)

            for i, future in enumerate(futures):
                if self.stop_flag:
                    break

                if future.result():
                    success_count += 1
                else:
                    failed_files.append(future.exception())

                # 更新进度条
                self.current_file_index = i + 1
                progress = (self.current_file_index / self.total_files) * 100
                self.progress_bar['value'] = progress
                self.status_label.config(text=f"正在解压: {os.path.basename(compressed_file)} ({self.current_file_index}/{self.total_files})")
                self.update_time_label()
                self.root.update_idletasks()  # 强制刷新UI

        if not self.stop_flag:
            if failed_files:
                messagebox.showwarning("部分完成",
                                       f"成功解压{success_count}个压缩文件，但有{len(failed_files)}个文件解压失败。\n失败文件: {', '.join(failed_files)}")
            else:
                messagebox.showinfo("完成", f"成功解压{success_count}个压缩文件！")
            self.confirm_delete_compressed_files()
        else:
            messagebox.showwarning("取消", "解压操作已被取消。")

        self.reset_ui()

    def extract_single_file(self, file_path):
        """解压单个文件（由线程池调用）"""
        try:
            compressed_file = os.path.basename(file_path)
            extract_dir = os.path.dirname(file_path)
            if compressed_file.lower().endswith('.zip'):
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    zip_ref.extractall(extract_dir)
                    self.extracted_files_size += sum(
                        os.path.getsize(os.path.join(extract_dir, f)) for f in os.listdir(extract_dir) if os.path.isfile(os.path.join(extract_dir, f)))
                    self.log(f"成功解压: {compressed_file}")
            else:
                patoolib.extract_archive(file_path, outdir=extract_dir)
                self.extracted_files_size += sum(
                    os.path.getsize(os.path.join(extract_dir, f)) for f in os.listdir(extract_dir) if os.path.isfile(os.path.join(extract_dir, f)))
                self.log(f"成功解压: {compressed_file}")
            return True
        except Exception as e:
            self.log(f"解压文件 {compressed_file} 时出现错误: {e}")
            logging.error(f"解压文件 {compressed_file} 时出现错误: {e}")
            return False

    def confirm_delete_compressed_files(self):
        """确认是否删除原压缩文件"""
        response = messagebox.askyesno("确认删除", "解压已完成。您是否要删除原始的压缩文件？")
        if response:
            deleted_files = []
            failed_deletions = []
            for compressed_file in self.compressed_files:
                try:
                    file_size = os.path.getsize(compressed_file)
                    os.remove(compressed_file)
                    deleted_files.append(compressed_file)
                    self.deleted_files_size += file_size
                    self.log(f"已删除: {compressed_file}")
                except Exception as e:
                    failed_deletions.append(compressed_file)
                    self.log(f"删除文件 {compressed_file} 时出现错误: {e}")
                    logging.error(f"删除文件 {compressed_file} 时出现错误: {e}")

            if failed_deletions:
                messagebox.showwarning("部分完成",
                                       f"成功删除{len(deleted_files)}个压缩文件，但有{len(failed_deletions)}个文件删除失败。\n失败文件: {', '.join(failed_deletions)}")
            else:
                messagebox.showinfo("完成", "原始压缩文件已全部删除。")
            self.generate_report()
        else:
            messagebox.showinfo("提醒", "原始压缩文件将保留。")
            self.generate_report()

    def generate_report(self):
        """生成并保存统计报告文件"""
        extracted_files_count = 0
        extracted_files_size = 0

        for root, dirs, files in os.walk(self.directory_path):
            for file in files:
                file_path = os.path.join(root, file)
                if not file_path.lower().endswith(SUPPORTED_FORMATS):  # 排除压缩文件
                    extracted_files_count += 1
                    extracted_files_size += os.path.getsize(file_path)

        extracted_files_size_gb = extracted_files_size / (1024 ** 3)
        deleted_files_size_gb = self.deleted_files_size / (1024 ** 3)

        report_message = (
            f"解压报告:\n"
            f"1. 成功解压文件数量: {extracted_files_count}\n"
            f"2. 解压文件总大小: {extracted_files_size_gb:.2f} GB\n"
            f"3. 删除原文件后释放空间: {deleted_files_size_gb:.2f} GB\n"
        )

        # 保存报告文件到解压目录
        report_file_path = os.path.join(self.directory_path, "解压报告.txt")
        with open(report_file_path, "w", encoding="utf-8") as report_file:
            report_file.write(report_message)

        messagebox.showinfo("解压报告", f"解压报告已保存到: {report_file_path}")
        self.open_extracted_directory()

    def open_extracted_directory(self):
        """打开解压后的文件夹"""
        if self.directory_path:
            try:
                os.startfile(self.directory_path)
            except AttributeError:
                # macOS 和 Linux 的处理方式
                if sys.platform == "darwin":
                    subprocess.Popen(["open", self.directory_path])
                else:
                    subprocess.Popen(["xdg-open", self.directory_path])

    def cancel_extraction(self):
        """取消解压操作"""
        self.stop_flag = True
        self.status_label.config(text="正在取消解压操作...")

    def reset_ui(self):
        """重置UI状态"""
        self.extract_button.config(state=tk.NORMAL)
        self.cancel_button.config(state=tk.DISABLED)
        self.status_label.config(text="")
        self.progress_bar['value'] = 0
        self.time_label.config(text="")

    def log(self, message):
        """更新日志文本框"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        logging.info(message)

    def show_help(self):
        """显示使用说明"""
        messagebox.showinfo("使用说明", self.translations["help_message"])

    def check_for_updates(self, silent=False):
        """检查更新"""
        latest_version = self.get_latest_version()
        if latest_version is None:
            if not silent:
                messagebox.showerror("错误", "无法连接到更新服务器，请检查网络连接。")
            return

        if latest_version > CURRENT_VERSION:
            if not silent:
                response = messagebox.askyesno("更新可用",
                                               f"发现新版本 {latest_version}，当前版本是 {CURRENT_VERSION}。是否立即更新？")
                if response:
                    self.download_and_install_update(latest_version)
            else:
                self.notify_user_of_update(latest_version)
        elif not silent:
            messagebox.showinfo("最新版本", "您正在使用最新版本。")

    def notify_user_of_update(self, latest_version):
        """通知用户有新版本可用"""
        response = messagebox.askyesno("更新可用", f"发现新版本 {latest_version}，当前版本是 {CURRENT_VERSION}。是否立即更新？")
        if response:
            self.download_and_install_update(latest_version)

    def get_latest_version(self):
        """获取最新版本号"""
        try:
            response = requests.get(UPDATE_URL)
            if response.status_code == 200:
                return response.text.strip()
            else:
                return None
        except requests.RequestException:
            return None

    def download_and_install_update(self, version):
        """下载并安装更新"""
        url = DOWNLOAD_URL.format(version=version)
        update_path = os.path.join(os.getcwd(), f"尘飞批量解压_v{version}.exe")

        try:
            response = requests.get(url, stream=True)
            if response.status_code == 200:
                with open(update_path, 'wb') as file:
                    for chunk in response.iter_content(chunk_size=8192):
                        file.write(chunk)
                messagebox.showinfo("更新完成", "更新文件已下载，即将启动更新程序。")
                self.launch_updater(update_path)
            else:
                messagebox.showerror("下载失败", "无法下载更新文件，请稍后再试。")
        except requests.RequestException:
            messagebox.showerror("下载失败", "无法连接到更新服务器，请检查网络连接。")

    def launch_updater(self, update_path):
        """启动更新程序"""
        try:
            subprocess.Popen([update_path])
            self.root.destroy()  # 关闭当前程序
        except Exception as e:
            messagebox.showerror("启动更新失败", f"无法启动更新程序: {e}")

    def check_for_updates_on_startup(self, silent=False):
        """启动时静默检查更新"""
        threading.Thread(target=lambda: self.check_for_updates(silent=silent)).start()

    def update_time_label(self):
        """更新时间标签"""
        elapsed_time = datetime.now() - self.start_time
        self.time_label.config(text=f"已用时间: {str(timedelta(seconds=int(elapsed_time.total_seconds())))}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ZipExtractorApp(root)
    root.mainloop()
