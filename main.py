import tkinter as tk
from tkinter import filedialog
from tkinter import scrolledtext
from pathlib import Path
import re
import threading
import openpyxl
from openpyxl.styles import Alignment, Font
from tkinter import ttk
import ttkbootstrap as tb
from ttkbootstrap.constants import *

class TextExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("书名提取工具-for Dingla")
        self.root.geometry("1100x700")
        self.setup_gui()
        
    def setup_gui(self):
        # 创建主框架
        main_frame = tb.Frame(self.root, bootstyle="light")
        main_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        # 标题标签
        title_label = tb.Label(
            main_frame, 
            text="书名提取工具", 
            font=("微软雅黑", 23, "bold"),
            bootstyle=SUCCESS
        )
        title_label.pack(pady=(0,15))
        
        # 目录选择区域
        dir_frame = tb.Frame(main_frame)
        dir_frame.pack(fill=X, pady=(0,10))
        
        tb.Label(dir_frame, text="请选择目录:", bootstyle=INFO).pack(anchor=W, pady=(0,5))
        
        # 使用ttkbootstrap的Entry支持拖放
        self.dir_entry = tb.Entry(dir_frame)
        self.dir_entry.pack(side=LEFT, fill=X, expand=True)
        self.dir_entry.bind("<Button-1>", self.browse_directory)  # 点击打开目录选择
        self.dir_entry.bind("<B1-Motion>", self.on_drag)  # 模拟拖放效果
        
        # 浏览按钮
        browse_btn = tb.Button(
            dir_frame, 
            width=16,
            text="浏览目录", 
            bootstyle=(OUTLINE,INFO),
            command=self.browse_directory
        )
        browse_btn.pack(side=RIGHT, padx=(5,0))
        
        # 处理按钮
        self.process_btn = tb.Button(
            main_frame, 
            width=16,
            text="开始提取内容", 
            bootstyle=PRIMARY,
            command=self.process_directory
        )
        self.process_btn.pack(pady=10)
        
        
        # 日志区域
        log_frame = tb.Frame(main_frame)
        log_frame.pack(fill=BOTH, expand=True, pady=(10,0))
        
        tb.Label(log_frame, text="处理日志:", bootstyle=INFO).pack(anchor=W, pady=(0,5))
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame, 
            wrap=tk.WORD,
            height=6,
            state=tk.DISABLED
        )
        self.log_text.pack(fill=BOTH, expand=True)

        # 进度条
        self.progress = tb.Progressbar(
            main_frame, 
            bootstyle=STRIPED,
            orient=HORIZONTAL,
            mode='determinate'
        )
        self.progress.pack(fill=X, pady=5)
        
        # 版权信息
        tb.Label(
            main_frame, 
            text="© 2025 文本处理工具 by Liug", 
            bootstyle=SECONDARY, 
            font=("Arial", 10)
        ).pack(side=BOTTOM, pady=5)
    
    def log_message(self, message):
        """将消息添加到日志区域"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)  # 自动滚动到底部
        self.log_text.config(state=tk.DISABLED)
        self.root.update_idletasks()  # 立即更新界面
    
    def on_drag(self, event):
        """模拟拖放操作（效果反馈）"""
        self.dir_entry.config(bootstyle="warning")
        
    def browse_directory(self, event=None):
        """打开文件选择对话框"""
        path = filedialog.askdirectory(title="请选择目录")
        if path:
            self.dir_entry.delete(0, tk.END)
            self.dir_entry.insert(0, path)
            self.dir_entry.config(bootstyle="success")
    
    def process_directory(self):
        """处理目录提取任务"""
        input_path = Path(self.dir_entry.get().strip())
        
        if not input_path.exists() or not input_path.is_dir():
            self.log_message("❌ 错误: 请选择有效的目录")
            self.dir_entry.config(bootstyle="danger")
            return
            
        # 禁用按钮防止多重处理
        self.process_btn.config(state=tk.DISABLED)
        self.progress["value"] = 0
        self.root.update_idletasks()
        
        # 在新线程中执行处理
        threading.Thread(
            target=self.execute_extraction, 
            args=(input_path,),
            daemon=True
        ).start()
    
    def execute_extraction(self, input_path: Path):
        """执行实际的文件处理"""
        try:
            self.log_message(f"🔍 开始扫描目录: {input_path}")
            
            # 获取所有txt文件
            txt_files = list(input_path.glob("*.txt"))
            if not txt_files:
                self.log_message("⚠️ 未找到任何.txt文件")
                self.finalize(False)
                return
                
            total_files = len(txt_files)
            self.log_message(f"📁 找到 {total_files} 个txt文件")
            
            # 准备Excel文件
            excel_path = input_path / "提取结果.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "提取内容"
            
            # 设置表头样式
            header_font = Font(bold=True)
            header_alignment = Alignment(horizontal="center")
            ws.column_dimensions["A"].width = 60
            ws.column_dimensions["B"].width = 25
            
            # 写入表头
            ws.append(["提取内容", "文件名"])
            for cell in ws[1]:
                cell.font = header_font
                cell.alignment = header_alignment
            
            # 处理所有txt文件
            processed_count = 0
            total_extracted = 0
            
            for i, txt_file in enumerate(txt_files):
                try:
                    with open(txt_file, "r", encoding="utf-8", errors="ignore") as f:
                        content = f.read()
                    
                    # 提取《》中的内容
                    matches = re.findall(r"《(.+?)》", content)
                    
                    if matches:
                        for match in matches:
                            ws.append([match, txt_file.name])
                        self.log_message(f"✓ {txt_file.name} 中提取到 {len(matches)} 条内容")
                        processed_count += 1
                        total_extracted += len(matches)
                    else:
                        self.log_message(f"➖ {txt_file.name} 中未找到《》内容")
                        
                    # 更新进度
                    progress = (i + 1) / total_files * 100
                    self.progress["value"] = progress
                    self.root.update_idletasks()
                
                except Exception as e:
                    self.log_message(f"⚠️ 无法读取文件 {txt_file.name}: {str(e)}")
            
            # 保存Excel文件
            wb.save(excel_path)
            self.log_message(f"✅ 处理完成! 共提取 {total_extracted} 条内容")
            self.log_message(f"💾 结果已保存到: {excel_path}")
            self.finalize(True)
            
        except Exception as e:
            self.log_message(f"❌ 处理过程中出错: {str(e)}")
            self.finalize(False)
    
    def finalize(self, success):
        """处理后清理工作"""
        self.progress["value"] = 100
        style = SUCCESS if success else DANGER
        self.progress.config(bootstyle=(style, STRIPED))
        
        # 重新启用处理按钮
        self.process_btn.config(state=tk.NORMAL)
        self.root.update_idletasks()

if __name__ == "__main__":
    root = tb.Window(themename="cosmo")
    icon_path = Path(__file__).parent / "docsicon.ico"
    if icon_path.exists():
        root.iconbitmap(str(icon_path))
    else:
        root.iconbitmap(None)
    
    app = TextExtractor(root)
    root.mainloop()
