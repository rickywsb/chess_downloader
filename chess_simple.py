#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
国际象棋教练工具 - 标签页版本
分页布局：班级管理 | 对阵表处理
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog
import json
import os
import re
import subprocess
import platform
import requests
from datetime import datetime, timedelta
import time

# 尝试导入dateutil，如果没有则使用替代方案
try:
    from dateutil.relativedelta import relativedelta
    HAS_DATEUTIL = True
except ImportError:
    HAS_DATEUTIL = False
    def relativedelta(months=0):
        """简单的月份偏移替代方案"""
        return timedelta(days=months * 30)

class ChessToolTabs:
    def __init__(self, root):
        self.root = root
        self.setup_window()
        self.students = {}
        self.classes_data = {}
        self.current_class = ""
        self.parsed_pairings = []
        self.create_widgets()
        self.load_data()
        
    def setup_window(self):
        """设置窗口"""
        self.root.title("国际象棋教练工具")
        
        # 获取屏幕尺寸
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # 设置窗口尺寸
        window_width = min(1000, int(screen_width * 0.8))
        window_height = min(750, int(screen_height * 0.85))
        
        # 居中
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.minsize(800, 600)
        
        # 字体设置
        if platform.system() == "Windows":
            self.font_normal = ("Microsoft YaHei", 10)
            self.font_button = ("Microsoft YaHei", 11, "bold")
            self.font_title = ("Microsoft YaHei", 12, "bold")
        else:
            self.font_normal = ("Arial", 10)
            self.font_button = ("Arial", 11, "bold")
            self.font_title = ("Arial", 12, "bold")
    
    def create_widgets(self):
        """创建界面 - 标签页布局"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建标签页控件
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # 创建两个标签页
        self.create_class_management_tab()
        self.create_pairing_tab()
        
        # 状态栏
        self.create_status_bar(main_frame)
        
    def create_class_management_tab(self):
        """创建班级管理标签页"""
        # 班级管理页面
        class_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(class_frame, text="班级管理")
        
        # 左侧：班级列表和操作
        left_panel = ttk.LabelFrame(class_frame, text="班级管理", padding="10")
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # 班级操作区
        class_ops_frame = ttk.Frame(left_panel)
        class_ops_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(class_ops_frame, text="当前班级:", font=self.font_normal).pack(side=tk.LEFT)
        self.class_var = tk.StringVar()
        self.class_combo = ttk.Combobox(class_ops_frame, textvariable=self.class_var, 
                                       width=20, font=self.font_normal, state="readonly")
        self.class_combo.pack(side=tk.LEFT, padx=5)
        self.class_combo.bind('<<ComboboxSelected>>', self.on_class_selected)
        
        # 班级操作按钮
        class_btn_frame = ttk.Frame(left_panel)
        class_btn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(class_btn_frame, text="新建班级", command=self.create_class, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(class_btn_frame, text="删除班级", command=self.delete_class, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(class_btn_frame, text="班级设置", command=self.edit_class, width=12).pack(side=tk.LEFT, padx=2)
        
        # 班级信息显示
        info_frame = ttk.LabelFrame(left_panel, text="班级信息", padding="5")
        info_frame.pack(fill=tk.X, pady=10)
        
        self.class_info_text = tk.Text(info_frame, height=4, font=self.font_normal, wrap=tk.WORD)
        self.class_info_text.pack(fill=tk.X)
        
        # 右侧：学员管理
        right_panel = ttk.LabelFrame(class_frame, text="学员管理", padding="10")
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # 学员操作按钮
        student_btn_frame = ttk.Frame(right_panel)
        student_btn_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(student_btn_frame, text="手动添加", command=self.add_student_manual, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(student_btn_frame, text="粘贴导入", command=self.import_students_paste, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(student_btn_frame, text="文件导入", command=self.import_students_file, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(student_btn_frame, text="删除学员", command=self.delete_student, width=12).pack(side=tk.LEFT, padx=2)
        
        # 学员列表
        students_frame = ttk.Frame(right_panel)
        students_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建表格
        columns = ("姓名", "用户名")
        self.students_tree = ttk.Treeview(students_frame, columns=columns, show="headings", height=15)
        
        # 设置列标题
        self.students_tree.heading("姓名", text="姓名")
        self.students_tree.heading("用户名", text="Chess.com用户名")
        
        # 设置列宽
        self.students_tree.column("姓名", width=150)
        self.students_tree.column("用户名", width=200)
        
        # 添加滚动条
        students_scrollbar = ttk.Scrollbar(students_frame, orient="vertical", command=self.students_tree.yview)
        self.students_tree.configure(yscrollcommand=students_scrollbar.set)
        
        self.students_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        students_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 双击编辑学员
        self.students_tree.bind("<Double-1>", self.edit_student)
    
    def create_pairing_tab(self):
        """创建对阵表处理标签页"""
        # 对阵表处理页面
        pairing_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(pairing_frame, text="对阵表处理")
        
        # 顶部：班级选择和操作按钮
        top_frame = ttk.Frame(pairing_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 班级选择
        ttk.Label(top_frame, text="处理班级:", font=self.font_normal).pack(side=tk.LEFT)
        self.pairing_class_var = tk.StringVar()
        self.pairing_class_combo = ttk.Combobox(top_frame, textvariable=self.pairing_class_var, 
                                               width=20, font=self.font_normal, state="readonly")
        self.pairing_class_combo.pack(side=tk.LEFT, padx=5)
        self.pairing_class_combo.bind('<<ComboboxSelected>>', self.on_pairing_class_selected)
        
        # 操作按钮
        btn_frame = ttk.Frame(top_frame)
        btn_frame.pack(side=tk.RIGHT)
        
        ttk.Button(btn_frame, text="粘贴对阵表", command=self.paste_pairings, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="解析对阵", command=self.parse_pairings, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="下载棋谱", command=self.download_games, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="清空", command=self.clear_pairings, width=8).pack(side=tk.LEFT, padx=2)
        
        # 中间：对阵表输入和解析结果（左右分栏）
        middle_frame = ttk.Frame(pairing_frame)
        middle_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 左侧：对阵表输入
        left_pairing = ttk.LabelFrame(middle_frame, text="对阵表输入", padding="5")
        left_pairing.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        self.pairings_text = scrolledtext.ScrolledText(
            left_pairing, 
            font=self.font_normal,
            wrap=tk.WORD
        )
        self.pairings_text.pack(fill=tk.BOTH, expand=True)
        
        # 右侧：解析结果
        right_pairing = ttk.LabelFrame(middle_frame, text="解析结果", padding="5")
        right_pairing.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        self.result_text = scrolledtext.ScrolledText(
            right_pairing,
            font=self.font_normal,
            wrap=tk.WORD
        )
        self.result_text.pack(fill=tk.BOTH, expand=True)
    
    def create_status_bar(self, parent):
        """创建状态栏"""
        status_frame = ttk.Frame(parent)
        status_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(status_frame, textvariable=self.status_var, font=self.font_normal).pack(side=tk.LEFT)
        
        # 进度条
        self.progress_bar = ttk.Progressbar(status_frame, mode='determinate', length=300)
        self.progress_bar.pack(side=tk.RIGHT, padx=5)
    
    
    def log(self, message):
        """更新状态"""
        if hasattr(self, 'status_var'):
            self.status_var.set(message)
            self.root.update()
    
    def load_data(self):
        """加载数据 - 兼容新旧格式"""
        try:
            if os.path.exists("classes.json"):
                with open("classes.json", "r", encoding="utf-8") as f:
                    data = json.load(f)
                
                # 检查数据格式
                if "classes" in data:
                    # 新格式
                    self.classes_data = data["classes"]
                    self.current_class = data.get("current_class", "")
                else:
                    # 旧格式：直接是班级字典
                    self.classes_data = data
                    if self.classes_data:
                        self.current_class = list(self.classes_data.keys())[0]
                
                self.log("数据加载成功")
            else:
                self.classes_data = {}
                self.current_class = ""
                self.log("创建新数据文件")
                
        except Exception as e:
            self.log(f"加载数据出错: {str(e)}")
            self.classes_data = {}
            self.current_class = ""
        
        # 更新界面
        self.update_class_lists()
        if self.current_class:
            self.class_var.set(self.current_class)
            self.pairing_class_var.set(self.current_class)
            self.on_class_selected()
    
    def save_data(self):
        """保存数据"""
        try:
            data = {
                "classes": self.classes_data,
                "current_class": self.current_class,
                "settings": {
                    "round_folder_format": "{class_name}-round{round_number}"
                }
            }
            with open("classes.json", "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self.log("数据保存成功")
        except Exception as e:
            self.log(f"保存数据出错: {str(e)}")
    
    def update_class_lists(self):
        """更新所有班级列表"""
        class_names = list(self.classes_data.keys())
        
        # 更新班级管理页的下拉框
        self.class_combo['values'] = class_names
        
        # 更新对阵表页的下拉框
        self.pairing_class_combo['values'] = class_names
    
    def on_class_selected(self, event=None):
        """班级选择事件 - 班级管理页"""
        class_name = self.class_var.get()
        if class_name and class_name in self.classes_data:
            self.current_class = class_name
            class_info = self.classes_data[class_name]
            
            # 获取学员数据
            if isinstance(class_info, dict):
                if "students" in class_info:
                    self.students = class_info["students"]
                    # 显示班级信息
                    info_text = f"班级: {class_name}\n"
                    info_text += f"创建日期: {class_info.get('created_date', '未知')}\n"
                    info_text += f"描述: {class_info.get('description', '无')}\n"
                    info_text += f"学员数量: {len(self.students)}"
                else:
                    # 旧格式：直接是学员字典
                    self.students = class_info
                    info_text = f"班级: {class_name}\n学员数量: {len(self.students)}"
            else:
                self.students = {}
                info_text = f"班级: {class_name}\n数据格式错误"
            
            # 更新班级信息显示
            self.class_info_text.delete(1.0, tk.END)
            self.class_info_text.insert(1.0, info_text)
            
            # 更新学员列表
            self.update_students_tree()
            
            # 同步对阵表页的班级选择
            self.pairing_class_var.set(class_name)
            
            self.log(f"已选择班级: {class_name}")
    
    def on_pairing_class_selected(self, event=None):
        """班级选择事件 - 对阵表页"""
        class_name = self.pairing_class_var.get()
        if class_name and class_name in self.classes_data:
            self.current_class = class_name
            class_info = self.classes_data[class_name]
            
            # 获取学员数据
            if isinstance(class_info, dict):
                if "students" in class_info:
                    self.students = class_info["students"]
                else:
                    self.students = class_info
            else:
                self.students = {}
            
            # 同步班级管理页的选择
            self.class_var.set(class_name)
            
            self.log(f"对阵表处理班级: {class_name}")
    
    def update_students_tree(self):
        """更新学员表格"""
        # 清空现有数据
        for item in self.students_tree.get_children():
            self.students_tree.delete(item)
        
        # 添加学员数据
        for real_name, username in self.students.items():
            self.students_tree.insert("", tk.END, values=(real_name, username))
    
    def create_class(self):
        """创建新班级"""
        # 创建对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("新建班级")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 50, self.root.winfo_rooty() + 50))
        
        # 班级名称
        ttk.Label(dialog, text="班级名称:", font=self.font_normal).pack(pady=5)
        name_var = tk.StringVar()
        name_entry = ttk.Entry(dialog, textvariable=name_var, width=30, font=self.font_normal)
        name_entry.pack(pady=5)
        name_entry.focus()
        
        # 班级描述
        ttk.Label(dialog, text="班级描述:", font=self.font_normal).pack(pady=5)
        desc_text = tk.Text(dialog, height=5, width=40, font=self.font_normal)
        desc_text.pack(pady=5)
        
        # 按钮
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        
        def create_action():
            class_name = name_var.get().strip()
            if not class_name:
                messagebox.showwarning("警告", "请输入班级名称！", parent=dialog)
                return
            
            if class_name in self.classes_data:
                messagebox.showwarning("警告", "班级名称已存在！", parent=dialog)
                return
            
            desc = desc_text.get(1.0, tk.END).strip()
            
            # 创建班级
            self.classes_data[class_name] = {
                "created_date": datetime.now().strftime("%Y-%m-%d"),
                "description": desc,
                "students": {}
            }
            
            self.current_class = class_name
            self.students = {}
            
            self.save_data()
            self.update_class_lists()
            self.class_var.set(class_name)
            self.pairing_class_var.set(class_name)
            self.on_class_selected()
            
            self.log(f"创建班级: {class_name}")
            dialog.destroy()
        
        ttk.Button(btn_frame, text="创建", command=create_action, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy, width=10).pack(side=tk.LEFT, padx=5)
        
        # 回车创建
        dialog.bind('<Return>', lambda e: create_action())
    
    def delete_class(self):
        """删除班级"""
        if not self.current_class:
            messagebox.showwarning("警告", "请先选择班级！")
            return
        
        result = messagebox.askyesno("确认删除", f"确定要删除班级 '{self.current_class}' 吗？\n此操作不可恢复！")
        if result:
            del self.classes_data[self.current_class]
            
            # 选择新的当前班级
            if self.classes_data:
                self.current_class = list(self.classes_data.keys())[0]
                self.class_var.set(self.current_class)
                self.pairing_class_var.set(self.current_class)
                self.on_class_selected()
            else:
                self.current_class = ""
                self.students = {}
                self.class_var.set("")
                self.pairing_class_var.set("")
                self.class_info_text.delete(1.0, tk.END)
                self.update_students_tree()
            
            self.save_data()
            self.update_class_lists()
            self.log("班级已删除")
    
    def edit_class(self):
        """编辑班级设置"""
        if not self.current_class:
            messagebox.showwarning("警告", "请先选择班级！")
            return
        
        # 编辑对话框
        dialog = tk.Toplevel(self.root)
        dialog.title(f"编辑班级 - {self.current_class}")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 50, self.root.winfo_rooty() + 50))
        
        class_info = self.classes_data[self.current_class]
        current_desc = class_info.get('description', '') if isinstance(class_info, dict) else ''
        
        # 班级描述
        ttk.Label(dialog, text="班级描述:", font=self.font_normal).pack(pady=5)
        desc_text = tk.Text(dialog, height=8, width=40, font=self.font_normal)
        desc_text.pack(pady=5)
        desc_text.insert(1.0, current_desc)
        
        # 按钮
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        
        def save_action():
            desc = desc_text.get(1.0, tk.END).strip()
            
            if isinstance(self.classes_data[self.current_class], dict):
                self.classes_data[self.current_class]['description'] = desc
            else:
                # 升级旧格式
                old_students = self.classes_data[self.current_class]
                self.classes_data[self.current_class] = {
                    "created_date": datetime.now().strftime("%Y-%m-%d"),
                    "description": desc,
                    "students": old_students
                }
            
            self.save_data()
            self.on_class_selected()  # 刷新显示
            self.log("班级信息已更新")
            dialog.destroy()
        
        ttk.Button(btn_frame, text="保存", command=save_action, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy, width=10).pack(side=tk.LEFT, padx=5)
    
    
    def add_student_manual(self):
        """手动添加学员"""
        if not self.current_class:
            messagebox.showwarning("警告", "请先选择班级！")
            return
        
        # 添加学员对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("添加学员")
        dialog.geometry("350x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 100, self.root.winfo_rooty() + 100))
        
        # 真实姓名
        ttk.Label(dialog, text="真实姓名:", font=self.font_normal).pack(pady=5)
        name_var = tk.StringVar()
        name_entry = ttk.Entry(dialog, textvariable=name_var, width=25, font=self.font_normal)
        name_entry.pack(pady=5)
        name_entry.focus()
        
        # Chess.com用户名
        ttk.Label(dialog, text="Chess.com用户名:", font=self.font_normal).pack(pady=5)
        username_var = tk.StringVar()
        username_entry = ttk.Entry(dialog, textvariable=username_var, width=25, font=self.font_normal)
        username_entry.pack(pady=5)
        
        # 按钮
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=20)
        
        def add_action():
            real_name = name_var.get().strip()
            username = username_var.get().strip()
            
            if not real_name or not username:
                messagebox.showwarning("警告", "请填写完整信息！", parent=dialog)
                return
            
            # 添加到当前班级
            class_info = self.classes_data[self.current_class]
            if isinstance(class_info, dict) and "students" in class_info:
                class_info["students"][real_name] = username
            else:
                # 兼容旧格式
                class_info[real_name] = username
            
            self.save_data()
            self.on_class_selected()  # 刷新显示
            self.log(f"添加学员: {real_name}")
            dialog.destroy()
        
        ttk.Button(btn_frame, text="添加", command=add_action, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy, width=10).pack(side=tk.LEFT, padx=5)
        
        # 回车添加
        dialog.bind('<Return>', lambda e: add_action())
    
    def import_students_paste(self):
        """粘贴导入学员"""
        if not self.current_class:
            messagebox.showwarning("警告", "请先选择班级！")
            return
        
        # 粘贴导入对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("粘贴导入学员")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 50, self.root.winfo_rooty() + 50))
        
        # 说明
        help_text = """支持的格式：
1. 张三 zhangsan
2. 李四 -> lisi  
3. 王五 wangwu
每行一个学员，程序会自动识别格式"""
        
        ttk.Label(dialog, text=help_text, font=self.font_normal, justify=tk.LEFT).pack(pady=5)
        
        # 文本输入区
        text_frame = ttk.Frame(dialog)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        paste_text = scrolledtext.ScrolledText(text_frame, font=self.font_normal, wrap=tk.WORD)
        paste_text.pack(fill=tk.BOTH, expand=True)
        paste_text.focus()
        
        # 尝试粘贴剪贴板内容
        try:
            clipboard_content = self.root.clipboard_get()
            paste_text.insert(1.0, clipboard_content)
        except:
            pass
        
        # 按钮
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        
        def import_action():
            content = paste_text.get(1.0, tk.END).strip()
            if not content:
                messagebox.showwarning("警告", "请输入学员数据！", parent=dialog)
                return
            
            try:
                students_data = self.parse_students_list(content)
                
                if students_data:
                    # 添加到当前班级
                    class_info = self.classes_data[self.current_class]
                    if isinstance(class_info, dict) and "students" in class_info:
                        class_info["students"].update(students_data)
                    else:
                        # 兼容旧格式
                        class_info.update(students_data)
                    
                    self.save_data()
                    self.on_class_selected()  # 刷新显示
                    self.log(f"导入 {len(students_data)} 名学员")
                    messagebox.showinfo("成功", f"成功导入 {len(students_data)} 名学员！", parent=dialog)
                    dialog.destroy()
                else:
                    messagebox.showwarning("警告", "未能解析到学员信息！请检查格式。", parent=dialog)
                    
            except Exception as e:
                messagebox.showerror("错误", f"导入失败: {str(e)}", parent=dialog)
        
        ttk.Button(btn_frame, text="导入", command=import_action, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy, width=10).pack(side=tk.LEFT, padx=5)
    
    def import_students_file(self):
        """文件导入学员"""
        if not self.current_class:
            messagebox.showwarning("警告", "请先选择班级！")
            return
        
        file_path = filedialog.askopenfilename(
            title="选择学员名单文件",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    content = f.read().strip()
                
                students_data = self.parse_students_list(content)
                
                if students_data:
                    # 添加到当前班级
                    class_info = self.classes_data[self.current_class]
                    if isinstance(class_info, dict) and "students" in class_info:
                        class_info["students"].update(students_data)
                    else:
                        # 兼容旧格式
                        class_info.update(students_data)
                    
                    self.save_data()
                    self.on_class_selected()  # 刷新显示
                    self.log(f"从文件导入 {len(students_data)} 名学员")
                    messagebox.showinfo("成功", f"成功导入 {len(students_data)} 名学员！")
                else:
                    messagebox.showwarning("警告", "未能解析到学员信息！")
                    
            except Exception as e:
                messagebox.showerror("错误", f"导入失败: {str(e)}")
    
    def delete_student(self):
        """删除选中的学员"""
        if not self.current_class:
            messagebox.showwarning("警告", "请先选择班级！")
            return
        
        selected = self.students_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要删除的学员！")
            return
        
        item = selected[0]
        values = self.students_tree.item(item, 'values')
        real_name = values[0]
        
        result = messagebox.askyesno("确认删除", f"确定要删除学员 '{real_name}' 吗？")
        if result:
            # 从当前班级删除
            class_info = self.classes_data[self.current_class]
            if isinstance(class_info, dict) and "students" in class_info:
                if real_name in class_info["students"]:
                    del class_info["students"][real_name]
            else:
                # 兼容旧格式
                if real_name in class_info:
                    del class_info[real_name]
            
            self.save_data()
            self.on_class_selected()  # 刷新显示
            self.log(f"删除学员: {real_name}")
    
    def edit_student(self, event):
        """双击编辑学员"""
        if not self.current_class:
            return
        
        selected = self.students_tree.selection()
        if not selected:
            return
        
        item = selected[0]
        values = self.students_tree.item(item, 'values')
        old_name = values[0]
        old_username = values[1]
        
        # 编辑对话框
        dialog = tk.Toplevel(self.root)
        dialog.title(f"编辑学员 - {old_name}")
        dialog.geometry("350x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 100, self.root.winfo_rooty() + 100))
        
        # 真实姓名
        ttk.Label(dialog, text="真实姓名:", font=self.font_normal).pack(pady=5)
        name_var = tk.StringVar(value=old_name)
        name_entry = ttk.Entry(dialog, textvariable=name_var, width=25, font=self.font_normal)
        name_entry.pack(pady=5)
        name_entry.select_range(0, tk.END)
        name_entry.focus()
        
        # Chess.com用户名
        ttk.Label(dialog, text="Chess.com用户名:", font=self.font_normal).pack(pady=5)
        username_var = tk.StringVar(value=old_username)
        username_entry = ttk.Entry(dialog, textvariable=username_var, width=25, font=self.font_normal)
        username_entry.pack(pady=5)
        
        # 按钮
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=20)
        
        def save_action():
            new_name = name_var.get().strip()
            new_username = username_var.get().strip()
            
            if not new_name or not new_username:
                messagebox.showwarning("警告", "请填写完整信息！", parent=dialog)
                return
            
            # 更新学员信息
            class_info = self.classes_data[self.current_class]
            if isinstance(class_info, dict) and "students" in class_info:
                # 删除旧记录
                if old_name in class_info["students"]:
                    del class_info["students"][old_name]
                # 添加新记录
                class_info["students"][new_name] = new_username
            else:
                # 兼容旧格式
                if old_name in class_info:
                    del class_info[old_name]
                class_info[new_name] = new_username
            
            self.save_data()
            self.on_class_selected()  # 刷新显示
            self.log(f"更新学员: {new_name}")
            dialog.destroy()
        
        ttk.Button(btn_frame, text="保存", command=save_action, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy, width=10).pack(side=tk.LEFT, padx=5)
        
        # 回车保存
        dialog.bind('<Return>', lambda e: save_action())
    
    def parse_students_list(self, content):
        """解析学员名单"""
        students = {}
        lines = content.strip().split('\n')
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # 移除序号
            line = re.sub(r'^\d+[\.\)]\s*', '', line)
            
            # 解析格式
            if '->' in line:
                parts = line.split('->')
                if len(parts) == 2:
                    real_name = parts[0].strip()
                    username = parts[1].strip()
                    students[real_name] = username
            else:
                parts = line.split()
                if len(parts) >= 2:
                    real_name = ' '.join(parts[:-1])
                    username = parts[-1]
                    students[real_name] = username
                elif len(parts) == 1:
                    students[parts[0]] = parts[0]
        
        return students
    
    # 对阵表处理方法
    def paste_pairings(self):
        """粘贴对阵表"""
        try:
            clipboard_content = self.root.clipboard_get()
            self.pairings_text.delete(1.0, tk.END)
            self.pairings_text.insert(1.0, clipboard_content)
            self.log("已粘贴对阵表")
        except:
            messagebox.showwarning("警告", "剪贴板无内容！")
    
    def clear_pairings(self):
        """清空对阵表"""
        self.pairings_text.delete(1.0, tk.END)
        self.result_text.delete(1.0, tk.END)
        self.parsed_pairings = []
        self.progress_bar['value'] = 0
        self.log("已清空对阵表")
    
    def parse_pairings(self):
        """解析对阵表"""
        if not self.students:
            messagebox.showwarning("警告", "请先选择有学员的班级！")
            return
        
        content = self.pairings_text.get(1.0, tk.END).strip()
        if not content:
            messagebox.showwarning("警告", "请输入对阵表！")
            return
        
        try:
            self.parsed_pairings = self.parse_pairings_content(content)
            
            # 显示结果
            self.result_text.delete(1.0, tk.END)
            result = "解析结果:\n" + "="*50 + "\n\n"
            
            for i, pairing in enumerate(self.parsed_pairings, 1):
                white = pairing.get('white', '未知')
                black = pairing.get('black', '未知')
                white_user = pairing.get('white_username', '未找到')
                black_user = pairing.get('black_username', '未找到')
                
                status = "✓" if white_user != "未找到" and black_user != "未找到" else "✗"
                
                result += f"{status} 第{i}局:\n"
                result += f"   白方: {white} → {white_user}\n"
                result += f"   黑方: {black} → {black_user}\n"
                result += "-" * 40 + "\n"
            
            matched = sum(1 for p in self.parsed_pairings 
                         if p.get('white_username') != '未找到' and p.get('black_username') != '未找到')
            
            result += f"\n统计: 总共{len(self.parsed_pairings)}局，成功匹配{matched}局"
            
            self.result_text.insert(1.0, result)
            self.log(f"解析完成: {len(self.parsed_pairings)}局，匹配{matched}局")
            
        except Exception as e:
            messagebox.showerror("错误", f"解析失败: {str(e)}")
            self.log(f"解析失败: {str(e)}")
    
    def parse_pairings_content(self, content):
        """解析对阵内容"""
        pairings = []
        lines = content.strip().split('\n')
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # 更强力的序号移除 - 处理各种序号格式
            # 匹配: 数字 + 可选符号 + 空格
            line = re.sub(r'^\s*\d+[\.\):\-\s]*\s*', '', line)
            line = line.strip()
            
            if not line:
                continue
            
            pairing = self.extract_pairing(line)
            if pairing:
                pairings.append(pairing)
        
        return pairings
    
    def extract_pairing(self, line):
        """提取对阵 - 增强版"""
        # 更多对阵格式的正则表达式
        patterns = [
            # 标准格式
            r'(.+?)\s+[vs对战VS]\s+(.+?)(?:\s|$)',
            r'(.+?)\s+-\s+(.+?)(?:\s|$)', 
            r'(.+?)\s+对\s+(.+?)(?:\s|$)',
            r'(.+?)\s*[:：]\s*(.+?)(?:\s|$)',
            # 简单空格分隔
            r'^(.+?)\s+(.+?)$',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, line, re.IGNORECASE)
            if match:
                white_name = match.group(1).strip()
                black_name = match.group(2).strip()
                
                # 验证名字不为空且不是纯数字
                if white_name and black_name and not white_name.isdigit() and not black_name.isdigit():
                    return {
                        'white': white_name,
                        'black': black_name,
                        'white_username': self.find_username(white_name),
                        'black_username': self.find_username(black_name)
                    }
        
        return None
    
    def find_username(self, name):
        """查找用户名 - 增强匹配算法"""
        if not name or not self.students:
            return "未找到"
        
        name = name.strip()
        
        # 1. 精确匹配
        if name in self.students:
            return self.students[name]
        
        # 2. 忽略大小写精确匹配
        for real_name, username in self.students.items():
            if real_name.lower() == name.lower():
                return username
        
        # 3. 部分匹配 - 输入名字包含在真实姓名中
        for real_name, username in self.students.items():
            if name.lower() in real_name.lower():
                return username
        
        # 4. 部分匹配 - 真实姓名包含在输入名字中  
        for real_name, username in self.students.items():
            if real_name.lower() in name.lower():
                return username
        
        # 5. 首字匹配 - 处理简称（如"Li"匹配"Li Xing"）
        name_first = name.split()[0].lower() if name.split() else name.lower()
        for real_name, username in self.students.items():
            real_first = real_name.split()[0].lower() if real_name.split() else real_name.lower()
            if name_first == real_first and len(name_first) > 1:
                return username
        
        # 6. 姓氏匹配 - 检查所有单词
        name_parts = [part.lower() for part in name.split() if len(part) > 1]
        for real_name, username in self.students.items():
            real_parts = [part.lower() for part in real_name.split() if len(part) > 1]
            for name_part in name_parts:
                for real_part in real_parts:
                    if name_part == real_part:
                        return username
        
        # 7. 模糊匹配 - 去除空格和标点
        clean_name = re.sub(r'[^\w]', '', name.lower())
        for real_name, username in self.students.items():
            clean_real = re.sub(r'[^\w]', '', real_name.lower())
            if clean_name == clean_real:
                return username
            # 包含匹配
            if clean_name in clean_real or clean_real in clean_name:
                if abs(len(clean_name) - len(clean_real)) <= 2:  # 长度差异不大
                    return username
        
        return "未找到"
    
    def download_games(self):
        """下载棋谱"""
        if not self.parsed_pairings:
            messagebox.showwarning("警告", "请先解析对阵表！")
            return
        
        # 创建下载文件夹
        folder = f"downloads_{self.current_class}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        os.makedirs(folder, exist_ok=True)
        
        total = len(self.parsed_pairings)
        self.progress_bar['maximum'] = total
        success = 0
        failed_pairings = []
        
        for i, pairing in enumerate(self.parsed_pairings):
            white_user = pairing['white_username']
            black_user = pairing['black_username']
            white_name = pairing['white']
            black_name = pairing['black']
            
            if white_user == "未找到" or black_user == "未找到":
                failed_pairings.append(f"第{i+1}局: {white_name} vs {black_name} (用户名未找到)")
                continue
            
            self.progress_bar['value'] = i + 1
            self.log(f"下载第{i+1}/{total}局: {white_name} vs {black_name}")
            self.root.update()
            
            try:
                # 下载两个玩家的所有棋谱
                white_success = self.download_player_games(white_user, folder, f"{white_name}({white_user})")
                black_success = self.download_player_games(black_user, folder, f"{black_name}({black_user})")
                
                if white_success or black_success:
                    success += 1
                else:
                    failed_pairings.append(f"第{i+1}局: {white_name} vs {black_name} (下载失败)")
                
            except Exception as e:
                self.log(f"下载第{i+1}局出错: {str(e)}")
                failed_pairings.append(f"第{i+1}局: {white_name} vs {black_name} (错误: {str(e)})")
        
        self.progress_bar['value'] = total
        
        # 创建下载报告
        report_file = os.path.join(folder, "下载报告.txt")
        with open(report_file, "w", encoding="utf-8") as f:
            f.write(f"下载报告 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 50 + "\n\n")
            f.write(f"班级: {self.current_class}\n")
            f.write(f"总对阵数: {total}\n")
            f.write(f"成功下载: {success}\n")
            f.write(f"失败数: {len(failed_pairings)}\n\n")
            
            if failed_pairings:
                f.write("失败详情:\n")
                for failure in failed_pairings:
                    f.write(f"- {failure}\n")
        
        self.log(f"下载完成: 成功{success}/{total}局")
        
        result_message = (f"下载完成!\n"
                         f"总计: {total} 局\n"
                         f"成功: {success} 局\n"
                         f"失败: {len(failed_pairings)} 局\n"
                         f"保存到: {folder}")
        
        messagebox.showinfo("下载完成", result_message)
        
        # 打开文件夹
        try:
            if platform.system() == "Windows":
                os.startfile(folder)
            elif platform.system() == "Darwin":
                subprocess.run(["open", folder])
            else:
                subprocess.run(["xdg-open", folder])
        except:
            pass
    
    def download_player_games(self, username, folder, display_name):
        """下载单个玩家的棋谱"""
        try:
            self.log(f"开始下载 {display_name} 的棋谱...")
            
            # 验证用户名
            if not username or username == "未找到":
                self.log(f"✗ {display_name} 用户名无效")
                return False
            
            # Chess.com API 获取玩家信息（先验证用户是否存在）
            player_url = f"https://api.chess.com/pub/player/{username}"
            try:
                player_response = requests.get(player_url, timeout=10)
                if player_response.status_code != 200:
                    self.log(f"✗ 用户 {username} 不存在或无法访问")
                    return False
            except requests.RequestException as e:
                self.log(f"✗ 验证用户 {username} 时网络错误: {str(e)}")
                return False
            
            current_date = datetime.now()
            downloaded_any = False
            
            # 尝试下载最近几个月的游戏
            for months_back in range(6):  # 增加到6个月
                # 计算日期
                if HAS_DATEUTIL:
                    archive_date = current_date - relativedelta(months=months_back)
                else:
                    # 简单的月份计算
                    year = current_date.year
                    month = current_date.month - months_back
                    while month <= 0:
                        month += 12
                        year -= 1
                    archive_date = datetime(year, month, 1)
                
                year = archive_date.year
                month = archive_date.month
                
                archive_url = f"https://api.chess.com/pub/player/{username}/games/{year}/{month:02d}"
                
                try:
                    self.log(f"  正在检查 {year}-{month:02d}...")
                    response = requests.get(archive_url, timeout=15)
                    
                    if response.status_code == 200:
                        games_data = response.json()
                        games = games_data.get('games', [])
                        
                        if games:
                            # 保存为PGN文件
                            filename = f"{display_name}_{year}_{month:02d}.pgn"
                            # 清理文件名中的特殊字符
                            filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
                            filepath = os.path.join(folder, filename)
                            
                            with open(filepath, "w", encoding="utf-8") as f:
                                f.write(f"# {display_name} 的棋谱 - {year}年{month}月\n")
                                f.write(f"# 下载时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                                f.write(f"# 总共 {len(games)} 局游戏\n")
                                f.write(f"# Chess.com用户名: {username}\n\n")
                                
                                game_count = 0
                                for idx, game in enumerate(games, 1):
                                    pgn = game.get('pgn', '')
                                    if pgn:
                                        f.write(f"[Event \"Chess.com - Game #{idx}\"]\n")
                                        f.write(pgn)
                                        if not pgn.endswith('\n'):
                                            f.write('\n')
                                        f.write('\n')
                                        game_count += 1
                            
                            self.log(f"  ✓ 已保存 {game_count} 局棋谱 ({year}-{month:02d})")
                            downloaded_any = True
                    
                    elif response.status_code == 404:
                        # 该月份没有游戏记录，这是正常的
                        pass
                    else:
                        self.log(f"  ! HTTP错误 {response.status_code} - {year}-{month:02d}")
                    
                    # 避免请求过于频繁
                    time.sleep(0.3)
                    
                except requests.Timeout:
                    self.log(f"  ! 请求超时 - {year}-{month:02d}")
                    continue
                except requests.RequestException as e:
                    self.log(f"  ! 网络错误 - {year}-{month:02d}: {str(e)}")
                    continue
                except Exception as e:
                    self.log(f"  ! 处理错误 - {year}-{month:02d}: {str(e)}")
                    continue
            
            if downloaded_any:
                self.log(f"✓ {display_name} 下载完成")
            else:
                self.log(f"✗ {display_name} 未找到任何游戏记录")
            
            return downloaded_any
            
        except Exception as e:
            self.log(f"✗ 下载 {display_name} 失败: {str(e)}")
            return False

def main():
    root = tk.Tk()
    app = ChessToolTabs(root)
    
    def on_closing():
        try:
            app.save_data()
        except:
            pass
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()
