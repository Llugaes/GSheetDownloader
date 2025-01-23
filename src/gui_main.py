import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import asyncio
from config_manager import ConfigManager
from gsheet_to_excel_async import download_multi_google_sheet_async, get_sheets_service_v4
import threading
import os

class GSheetDownloaderGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Google Sheet 下载器")
        self.config_manager = ConfigManager()
        self.download_buttons = []  # 用于存储下载相关的按钮
        self.setup_ui()
        self.load_recent_sheets()
        self.update_auth_status()  # 添加认证状态检查

    def setup_ui(self):
        # URL输入框
        url_frame = ttk.Frame(self.root)
        url_frame.pack(padx=10, pady=5, fill=tk.X)
        ttk.Label(url_frame, text="Sheet URL:").pack(side=tk.LEFT)
        self.url_entry = ttk.Entry(url_frame)
        self.url_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(url_frame, text="添加", command=self.add_sheet).pack(side=tk.RIGHT)

        # 已添加的Sheets列表
        list_frame = ttk.LabelFrame(self.root, text="已添加的Sheets")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.sheet_list = ttk.Treeview(list_frame, columns=("url",), show="headings", height=10)
        self.sheet_list.heading("url", text="Sheet URL")
        self.sheet_list.column("url", width=400)
        self.sheet_list.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.sheet_list.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.sheet_list.configure(yscrollcommand=scrollbar.set)
        
        # 创建右键菜单
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="编辑", command=self.edit_sheet)
        self.context_menu.add_command(label="删除", command=self.delete_sheet)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="下载选中", command=self.download_selected)
        
        # 绑定右键菜单
        self.sheet_list.bind("<Button-3>", self.show_context_menu)
        self.sheet_list.bind("<Double-1>", lambda e: self.edit_sheet())

        # 输出目录选择
        dir_frame = ttk.Frame(self.root)
        dir_frame.pack(padx=10, pady=5, fill=tk.X)
        ttk.Label(dir_frame, text="输出目录:").pack(side=tk.LEFT)
        self.dir_entry = ttk.Entry(dir_frame)
        self.dir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.dir_entry.insert(0, self.config_manager.config.get('output_dir', ''))
        ttk.Button(dir_frame, text="选择", command=self.select_output_dir).pack(side=tk.RIGHT)

        # 下载按钮框架
        download_frame = ttk.Frame(self.root)
        download_frame.pack(pady=10)
        
        download_all_btn = ttk.Button(download_frame, text="下载全部", command=self.start_download)
        download_all_btn.pack(side=tk.LEFT, padx=5)
        self.download_buttons.append(download_all_btn)
        
        download_selected_btn = ttk.Button(download_frame, text="下载选中", command=self.download_selected)
        download_selected_btn.pack(side=tk.LEFT, padx=5)
        self.download_buttons.append(download_selected_btn)
        
        settings_button = ttk.Button(download_frame, text="认证设置", command=self.show_auth_settings)
        settings_button.pack(side=tk.LEFT, padx=5)

    def add_sheet(self):
        url = self.url_entry.get().strip()
        if url:
            self.config_manager.add_sheet(url)
            self.load_recent_sheets()
            self.url_entry.delete(0, tk.END)

    def load_recent_sheets(self):
        # 清空现有列表
        for item in self.sheet_list.get_children():
            self.sheet_list.delete(item)
        # 加载配置中的sheets
        for sheet in self.config_manager.config['recent_sheets']:
            self.sheet_list.insert('', tk.END, values=(sheet['url'],))

    def select_output_dir(self):
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.dir_entry.delete(0, tk.END)
            self.dir_entry.insert(0, dir_path)
            self.config_manager.config['output_dir'] = dir_path
            self.config_manager.save_config()

    async def download_with_progress(self, sheet_ids, output_dir):
        try:
            await download_multi_google_sheet_async(sheet_ids, output_dir)
            self.root.after(0, lambda: messagebox.showinfo("完成", "下载完成！"))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"下载出错: {str(e)}"))

    def start_download(self):
        print("[GUI] 开始下载全部")
        output_dir = self.dir_entry.get().strip()
        print(f"[GUI] 输出目录: {output_dir}")
        
        # 确保输出目录是有效的
        if not output_dir:
            output_dir = os.path.join(os.path.expanduser("~"), "Downloads")
            print(f"[GUI] 使用默认下载目录: {output_dir}")
            self.dir_entry.delete(0, tk.END)
            self.dir_entry.insert(0, output_dir)
        
        # 转换为绝对路径
        output_dir = os.path.abspath(output_dir)
        print(f"[GUI] 使用绝对路径: {output_dir}")
        
        try:
            os.makedirs(output_dir, exist_ok=True)
            print(f"[GUI] 输出目录已确认: {output_dir}")
        except Exception as e:
            print(f"[GUI] 创建目录失败: {str(e)}")
            messagebox.showerror("错误", f"创建输出目录失败: {str(e)}")
            return

        # 从配置中获取sheet_ids
        sheet_ids = []
        print(f"[GUI] 当前配置中的sheets: {self.config_manager.config['recent_sheets']}")
        
        for sheet in self.config_manager.config['recent_sheets']:
            if sheet.get('id'):
                sheet_ids.append(sheet['id'])
                print(f"[GUI] 添加sheet_id: {sheet['id']}")
        
        print(f"[GUI] 最终的sheet_ids列表: {sheet_ids}")
        
        if not sheet_ids:
            print("[GUI] 没有找到有效的sheet_ids")
            messagebox.showerror("错误", "请至少添加一个有效的Sheet")
            return

        # 创建并显示进度窗口
        progress_window = tk.Toplevel(self.root)
        progress_window.title("下载进度")
        progress_window.geometry("300x150")
        progress_window.transient(self.root)
        progress_window.grab_set()
        
        info_label = ttk.Label(progress_window, text="正在准备下载...")
        info_label.pack(pady=10)
        
        progress_bar = ttk.Progressbar(progress_window, mode='indeterminate')
        progress_bar.pack(fill=tk.X, padx=20, pady=10)
        progress_bar.start()

        def run_download():
            try:
                print(f"[GUI] 开始异步下载，参数：sheet_ids={sheet_ids}, output_dir={output_dir}")
                asyncio.run(download_multi_google_sheet_async(sheet_ids, output_dir))
                print("[GUI] 下载完成")
                progress_window.after(0, progress_window.destroy)
                messagebox.showinfo("完成", "下载完成！")
            except Exception as e:
                print(f"[GUI] 下载出错: {str(e)}")
                progress_window.after(0, progress_window.destroy)
                messagebox.showerror("错误", f"下载出错: {str(e)}")

        thread = threading.Thread(target=run_download, daemon=True)
        thread.start()

    def show_context_menu(self, event):
        item = self.sheet_list.identify_row(event.y)
        if item:
            self.sheet_list.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)
            
    def edit_sheet(self):
        selected = self.sheet_list.selection()
        if not selected:
            return
            
        item = selected[0]
        url = self.sheet_list.item(item)['values'][0]
        
        # 创建编辑对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("编辑 Sheet")
        dialog.geometry("400x100")
        
        ttk.Label(dialog, text="Sheet URL:").pack(padx=5, pady=5)
        entry = ttk.Entry(dialog, width=50)
        entry.insert(0, url)
        entry.pack(padx=5, pady=5)
        
        def save_changes():
            new_url = entry.get().strip()
            if new_url:
                # 更新配置
                for sheet in self.config_manager.config['recent_sheets']:
                    if sheet['url'] == url:
                        sheet['url'] = new_url
                        sheet['id'] = self.config_manager.extract_sheet_id(new_url)
                        break
                self.config_manager.save_config()
                self.load_recent_sheets()
                dialog.destroy()
        
        ttk.Button(dialog, text="保存", command=save_changes).pack(pady=5)
        
    def delete_sheet(self):
        selected = self.sheet_list.selection()
        if not selected:
            return
            
        if messagebox.askyesno("确认", "确定要删除选中的 Sheet 吗？"):
            item = selected[0]
            url = self.sheet_list.item(item)['values'][0]
            
            # 从配置中删除
            self.config_manager.config['recent_sheets'] = [
                sheet for sheet in self.config_manager.config['recent_sheets']
                if sheet['url'] != url
            ]
            self.config_manager.save_config()
            self.load_recent_sheets()

    def download_selected(self):
        selected = self.sheet_list.selection()
        if not selected:
            messagebox.showwarning("提示", "请先选择要下载的Sheet")
            return

        output_dir = self.dir_entry.get().strip()
        if not output_dir:
            output_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                messagebox.showerror("错误", f"创建输出目录失败: {str(e)}")
                return

        # 从配置中获取选中项的sheet_ids
        sheet_ids = []
        recent_sheets = {sheet['url']: sheet['id'] for sheet in self.config_manager.config['recent_sheets']}
        
        for item in selected:
            url = self.sheet_list.item(item)['values'][0]
            if url in recent_sheets:
                sheet_ids.append(recent_sheets[url])
        
        if not sheet_ids:
            messagebox.showerror("错误", "无法获取选中Sheet的ID")
            return

        # 创建进度窗口
        progress_window = tk.Toplevel(self.root)
        progress_window.title("下载进度")
        progress_window.geometry("300x150")
        progress_window.transient(self.root)
        progress_window.grab_set()
        
        info_label = ttk.Label(progress_window, text="正在准备下载...")
        info_label.pack(pady=10)
        
        progress_bar = ttk.Progressbar(progress_window, mode='indeterminate')
        progress_bar.pack(fill=tk.X, padx=20, pady=10)
        progress_bar.start()

        def run_download():
            try:
                asyncio.run(download_multi_google_sheet_async(sheet_ids, output_dir))
                progress_window.after(0, progress_window.destroy)
                messagebox.showinfo("完成", "下载完成！")
            except Exception as e:
                progress_window.after(0, progress_window.destroy)
                messagebox.showerror("错误", f"下载出错: {str(e)}")

        thread = threading.Thread(target=run_download, daemon=True)
        thread.start()

    def update_auth_status(self):
        """更新认证状态和按钮状态"""
        app_data_dir = os.path.join(os.path.expanduser("~"), ".gsheet_downloader")
        creds_path = os.path.join(app_data_dir, "credentials.json")
        token_path = os.path.join(app_data_dir, "token.json")
        
        is_authenticated = os.path.exists(creds_path) and os.path.exists(token_path)
        
        # 更新下载按钮状态
        for btn in self.download_buttons:
            btn.configure(state='normal' if is_authenticated else 'disabled')

    def show_auth_settings(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("认证设置")
        dialog.geometry("500x250")
        dialog.transient(self.root)
        
        # 当前认证状态
        status_frame = ttk.LabelFrame(dialog, text="认证状态")
        status_frame.pack(fill=tk.X, padx=10, pady=5)
        
        app_data_dir = os.path.join(os.path.expanduser("~"), ".gsheet_downloader")
        creds_path = os.path.join(app_data_dir, "credentials.json")
        token_path = os.path.join(app_data_dir, "token.json")
        
        is_authenticated = os.path.exists(creds_path) and os.path.exists(token_path)
        status_label = ttk.Label(status_frame, 
            text=f"状态：{'已认证' if is_authenticated else '未认证'}\n"
                 f"认证文件位置：{app_data_dir}")
        status_label.pack(pady=5)
        
        # 认证文件管理按钮框架
        btn_frame = ttk.Frame(status_frame)
        btn_frame.pack(pady=5)
        
        def remove_auth():
            if messagebox.askyesno("确认", "确定要删除认证信息吗？\n删除后需要重新认证才能使用。"):
                try:
                    if os.path.exists(creds_path):
                        os.remove(creds_path)
                    if os.path.exists(token_path):
                        os.remove(token_path)
                    messagebox.showinfo("成功", "认证信息已删除")
                    self.update_auth_status()
                    dialog.destroy()
                except Exception as e:
                    messagebox.showerror("错误", f"删除认证文件失败: {str(e)}")
        
        def open_auth_folder():
            if os.path.exists(app_data_dir):
                os.startfile(app_data_dir) if os.name == 'nt' else os.system(f'open "{app_data_dir}"')
            else:
                messagebox.showinfo("提示", "认证文件夹不存在")
        
        if is_authenticated:
            ttk.Button(btn_frame, text="删除认证", command=remove_auth).pack(side=tk.LEFT, padx=5)
            ttk.Button(btn_frame, text="打开文件夹", command=open_auth_folder).pack(side=tk.LEFT, padx=5)
        
        # 认证设置
        creds_frame = ttk.LabelFrame(dialog, text="认证设置")
        creds_frame.pack(fill=tk.X, padx=10, pady=5)
        
        def select_credentials():
            file_path = filedialog.askopenfilename(
                title="选择认证文件",
                filetypes=[("JSON文件", "*.json")]
            )
            if file_path:
                try:
                    app_data_dir = os.path.join(os.path.expanduser("~"), ".gsheet_downloader")
                    os.makedirs(app_data_dir, exist_ok=True)
                    
                    creds_path = os.path.join(app_data_dir, "credentials.json")
                    import shutil
                    shutil.copy2(file_path, creds_path)
                    
                    # 设置环境变量
                    os.environ['GCP_CREDENTIALS_JSON'] = creds_path
                    os.environ['GCP_TOKEN_JSON'] = os.path.join(app_data_dir, "token.json")
                    
                    # 关闭当前窗口，打开认证窗口
                    dialog.destroy()
                    self.show_auth_dialog()
                except Exception as e:
                    messagebox.showerror("错误", f"设置认证文件失败: {str(e)}")
        
        ttk.Button(creds_frame, text="选择认证文件", command=select_credentials).pack(pady=10)
        ttk.Label(creds_frame, text="提示：请从Google Cloud Console下载认证文件").pack(pady=5)

    def show_auth_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Google认证")
        dialog.geometry("500x250")
        dialog.transient(self.root)
        
        ttk.Label(dialog, text="请在浏览器中完成Google账号授权", font=('', 12, 'bold')).pack(pady=10)
        ttk.Label(dialog, text="系统将自动打开浏览器进行认证").pack(pady=5)
        
        # 创建进度条
        progress = ttk.Progressbar(dialog, mode='indeterminate')
        progress.pack(fill=tk.X, padx=20, pady=10)
        progress.start()
        
        def start_auth():
            try:
                service = get_sheets_service_v4()
                dialog.after(0, lambda: self.auth_success(dialog))
            except Exception as e:
                dialog.after(0, lambda: messagebox.showerror("错误", f"认证失败: {str(e)}"))
                dialog.after(0, dialog.destroy)
        
        threading.Thread(target=start_auth, daemon=True).start()

    def auth_success(self, dialog):
        dialog.destroy()
        self.update_auth_status()
        messagebox.showinfo("成功", "认证成功！")

def main():
    app = GSheetDownloaderGUI()
    app.root.mainloop()

if __name__ == "__main__":
    main() 