#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""国际象棋教练前端 GUI。

提供班级/学员管理、对阵解析以及 Chess.com 棋谱下载功能。
"""

from __future__ import annotations

import json
import os
import platform
import re
import subprocess
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional

import requests
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog, ttk
from requests.adapters import HTTPAdapter

try:
    from dateutil.relativedelta import relativedelta
    HAS_DATEUTIL = True
except ImportError:  # 本地无依赖时退化处理
    HAS_DATEUTIL = False

    def relativedelta(months: int = 0):  # type: ignore
        return timedelta(days=months * 30)


class TeacherChessApp:
    DATA_FILE = "classes.json"

    def __init__(self, root: tk.Tk):
        self.root = root
        self.base_dir = Path(__file__).resolve().parent
        self.data_path = self.base_dir / self.DATA_FILE

        self.classes_data: Dict[str, Dict] = {}
        self.current_class: str = ""
        self.students: Dict[str, str] = {}
        self.parsed_pairings: List[Dict[str, str]] = []
        self.archive_cache: Dict[str, List[Dict[str, Any]]] = {}
        self.archive_month_limit = 18
        self.recent_days = 14
        self.http = requests.Session()
        self.http.headers.update(
            {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/120.0 Safari/537.36"
                ),
                "Accept": "application/json, text/plain, */*",
            }
        )
        self.http.mount("https://", HTTPAdapter(max_retries=2))

        self._setup_window()
        self._build_ui()
        self.load_data()

    @staticmethod
    def sanitize_username(username: str) -> str:
        """移除用户名中的空格和不可见字符。"""
        sanitized = re.sub(r"\s+", "", username.strip())
        return sanitized

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def _setup_window(self) -> None:
        self.root.title("国际象棋教练工具（GUI）")
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        win_w = min(1080, int(screen_w * 0.82))
        win_h = min(780, int(screen_h * 0.85))
        x = (screen_w - win_w) // 2
        y = (screen_h - win_h) // 3
        self.root.geometry(f"{win_w}x{win_h}+{x}+{y}")
        self.root.minsize(860, 620)

        # 字体设置
        if platform.system() == "Windows":
            self.font_normal = ("Microsoft YaHei", 10)
            self.font_button = ("Microsoft YaHei", 11, "bold")
            self.font_title = ("Microsoft YaHei", 12, "bold")
        else:
            self.font_normal = ("Arial", 10)
            self.font_button = ("Arial", 11, "bold")
            self.font_title = ("Arial", 12, "bold")

    def _build_ui(self) -> None:
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)

        self.class_tab = ttk.Frame(notebook, padding=10)
        self.pairing_tab = ttk.Frame(notebook, padding=10)

        notebook.add(self.class_tab, text="班级管理")
        notebook.add(self.pairing_tab, text="对阵下载")

        self._build_class_tab()
        self._build_pairing_tab()
        self._build_status_bar(main_frame)

    def _build_class_tab(self) -> None:
        top_frame = ttk.Frame(self.class_tab)
        top_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(top_frame, text="当前班级:", font=self.font_normal).pack(side=tk.LEFT)
        self.class_var = tk.StringVar()
        self.class_combo = ttk.Combobox(
            top_frame,
            textvariable=self.class_var,
            state="readonly",
            width=30,
            font=self.font_normal,
        )
        self.class_combo.pack(side=tk.LEFT, padx=5)
        self.class_combo.bind("<<ComboboxSelected>>", self.on_class_selected)

        ttk.Button(top_frame, text="新建班级", command=self.create_class, width=12).pack(side=tk.LEFT, padx=4)
        ttk.Button(top_frame, text="删除班级", command=self.delete_class, width=12).pack(side=tk.LEFT, padx=4)
        ttk.Button(top_frame, text="班级信息", command=self.edit_class, width=12).pack(side=tk.LEFT, padx=4)

        # 分栏区
        content = ttk.Frame(self.class_tab)
        content.pack(fill=tk.BOTH, expand=True)

        left = ttk.LabelFrame(content, text="班级信息", padding=10)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))

        self.class_info = tk.Text(left, height=8, font=self.font_normal, wrap=tk.WORD, state=tk.DISABLED)
        self.class_info.pack(fill=tk.X)

        right = ttk.LabelFrame(content, text="学员管理", padding=10)
        right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(6, 0))

        btn_row = ttk.Frame(right)
        btn_row.pack(fill=tk.X, pady=(0, 8))

        ttk.Button(btn_row, text="手动添加", width=12, command=self.add_student_manual).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_row, text="粘贴导入", width=12, command=self.import_students_paste).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_row, text="文件导入", width=12, command=self.import_students_file).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_row, text="删除学员", width=12, command=self.delete_student).pack(side=tk.LEFT, padx=3)

        columns = ("name", "username")
        self.student_tree = ttk.Treeview(right, columns=columns, show="headings")
        self.student_tree.heading("name", text="姓名")
        self.student_tree.heading("username", text="Chess.com 用户名")
        self.student_tree.column("name", width=180)
        self.student_tree.column("username", width=200)
        self.student_tree.pack(fill=tk.BOTH, expand=True)
        self.student_tree.bind("<Double-1>", self.edit_student)

        scrollbar = ttk.Scrollbar(self.student_tree, orient=tk.VERTICAL, command=self.student_tree.yview)
        self.student_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _build_pairing_tab(self) -> None:
        top = ttk.Frame(self.pairing_tab)
        top.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(top, text="处理班级:", font=self.font_normal).pack(side=tk.LEFT)
        self.pairing_class_var = tk.StringVar()
        self.pairing_combo = ttk.Combobox(
            top,
            textvariable=self.pairing_class_var,
            state="readonly",
            width=30,
            font=self.font_normal,
        )
        self.pairing_combo.pack(side=tk.LEFT, padx=5)
        self.pairing_combo.bind("<<ComboboxSelected>>", self.on_pairing_class_selected)

        ttk.Button(top, text="粘贴对阵表", width=12, command=self.paste_pairings).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="解析对阵", width=12, command=self.parse_pairings).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="下载棋谱", width=12, command=self.download_games).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="清空", width=8, command=self.clear_pairings).pack(side=tk.LEFT, padx=4)

        middle = ttk.Frame(self.pairing_tab)
        middle.pack(fill=tk.BOTH, expand=True)

        input_frame = ttk.LabelFrame(middle, text="对阵表输入", padding=8)
        input_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))

        self.pairings_text = scrolledtext.ScrolledText(input_frame, font=self.font_normal, wrap=tk.WORD)
        self.pairings_text.pack(fill=tk.BOTH, expand=True)

        result_frame = ttk.LabelFrame(middle, text="解析结果", padding=8)
        result_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(6, 0))

        self.result_text = scrolledtext.ScrolledText(result_frame, font=self.font_normal, wrap=tk.WORD)
        self.result_text.pack(fill=tk.BOTH, expand=True)

    def _build_status_bar(self, parent: ttk.Frame) -> None:
        status_frame = ttk.Frame(parent)
        status_frame.pack(fill=tk.X, pady=(8, 0))

        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(status_frame, textvariable=self.status_var, font=self.font_normal).pack(side=tk.LEFT)

        self.progress = ttk.Progressbar(status_frame, mode="determinate", length=320)
        self.progress.pack(side=tk.RIGHT)

    def log(self, text: str) -> None:
        self.status_var.set(text)
        self.root.update_idletasks()

    # ------------------------------------------------------------------
    # 数据读写
    # ------------------------------------------------------------------
    def load_data(self) -> None:
        if not self.data_path.exists():
            self.classes_data = {}
            self.current_class = ""
            self.students = {}
            self._refresh_class_widgets()
            return

        try:
            with self.data_path.open("r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as exc:
            messagebox.showerror("读取失败", f"无法读取 {self.data_path.name}: {exc}")
            self.classes_data = {}
            self.current_class = ""
            self.students = {}
            self._refresh_class_widgets()
            return

        if isinstance(data, dict) and "classes" in data:
            self.classes_data = data.get("classes", {})
            self.current_class = data.get("current_class", "")
        elif isinstance(data, dict):  # 兼容旧版
            self.classes_data = data
            self.current_class = next(iter(data), "")
        else:
            self.classes_data = {}
            self.current_class = ""

        if self.current_class not in self.classes_data:
            self.current_class = next(iter(self.classes_data), "")

        self._refresh_students_ref()
        self._refresh_class_widgets()
        self.log("数据加载完成")

    def save_data(self) -> None:
        payload = {
            "classes": self.classes_data,
            "current_class": self.current_class,
            "settings": {"round_folder_format": "{class_name}-round{round_number}"},
        }
        try:
            with self.data_path.open("w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
            self.log("数据已保存")
        except Exception as exc:
            messagebox.showerror("保存失败", f"无法写入 {self.data_path.name}: {exc}")

    def _refresh_students_ref(self) -> None:
        if self.current_class and self.current_class in self.classes_data:
            info = self.classes_data[self.current_class]
            if isinstance(info, dict):
                self.students = info.setdefault("students", {})
            else:
                self.students = {}
        else:
            self.students = {}

    def _refresh_class_widgets(self) -> None:
        class_names = list(self.classes_data)
        self.class_combo["values"] = class_names
        self.pairing_combo["values"] = class_names

        if self.current_class:
            self.class_var.set(self.current_class)
            self.pairing_class_var.set(self.current_class)
        else:
            self.class_var.set("")
            self.pairing_class_var.set("")

        self._update_class_info()
        self._update_student_tree()

    # ------------------------------------------------------------------
    # 班级操作
    # ------------------------------------------------------------------
    def on_class_selected(self, _event=None) -> None:
        name = self.class_var.get()
        if name and name in self.classes_data:
            self.current_class = name
            self._refresh_students_ref()
            self._update_class_info()
            self._update_student_tree()
            self.pairing_class_var.set(name)
            self.save_data()
            self.log(f"已选择班级: {name}")

    def on_pairing_class_selected(self, _event=None) -> None:
        name = self.pairing_class_var.get()
        if name and name in self.classes_data:
            self.current_class = name
            self.class_var.set(name)
            self._refresh_students_ref()
            self._update_class_info()
            self._update_student_tree()
            self.save_data()
            self.log(f"对阵班级: {name}")

    def create_class(self) -> None:
        name = simple_input("新建班级", "请输入班级名称:")
        if not name:
            return
        if name in self.classes_data:
            messagebox.showwarning("提示", "班级名称已存在。")
            return
        desc = simple_input("班级描述", "请输入班级描述（可留空）:") or ""
        self.classes_data[name] = {
            "created_date": datetime.now().strftime("%Y-%m-%d"),
            "description": desc,
            "students": {},
        }
        self.current_class = name
        self._refresh_students_ref()
        self._refresh_class_widgets()
        self.save_data()
        self.log(f"创建班级 {name}")

    def delete_class(self) -> None:
        if not self.current_class:
            messagebox.showwarning("提示", "请先选择班级。")
            return
        if not messagebox.askyesno("确认删除", f"确定删除班级 '{self.current_class}'?\n该操作不可恢复。"):
            return
        self.classes_data.pop(self.current_class, None)
        self.current_class = next(iter(self.classes_data), "") if self.classes_data else ""
        self._refresh_students_ref()
        self._refresh_class_widgets()
        self.save_data()
        self.log("班级已删除")

    def edit_class(self) -> None:
        if not self.current_class:
            messagebox.showwarning("提示", "请先选择班级。")
            return
        info = self.classes_data.get(self.current_class, {})
        desc_old = info.get("description", "") if isinstance(info, dict) else ""
        desc_new = multiline_input("班级信息", "修改班级描述:", desc_old)
        if desc_new is None:
            return
        if isinstance(info, dict):
            info["description"] = desc_new.strip()
            info.setdefault("created_date", datetime.now().strftime("%Y-%m-%d"))
        self._refresh_class_widgets()
        self.save_data()
        self.log("班级描述已更新")

    def _update_class_info(self) -> None:
        self.class_info.configure(state=tk.NORMAL)
        self.class_info.delete("1.0", tk.END)
        if not self.current_class:
            self.class_info.insert("1.0", "尚未选择班级。")
        else:
            info = self.classes_data.get(self.current_class, {})
            created = info.get("created_date", "未知") if isinstance(info, dict) else "未知"
            desc = info.get("description", "") if isinstance(info, dict) else ""
            count = len(self.students)
            text = f"班级: {self.current_class}\n创建日期: {created}\n学员数量: {count}\n描述: {desc or '无'}"
            self.class_info.insert("1.0", text)
        self.class_info.configure(state=tk.DISABLED)

    def _update_student_tree(self) -> None:
        for item in self.student_tree.get_children():
            self.student_tree.delete(item)
        for name, username in sorted(self.students.items()):
            self.student_tree.insert("", tk.END, values=(name, username))

    # ------------------------------------------------------------------
    # 学员操作
    # ------------------------------------------------------------------
    def add_student_manual(self) -> None:
        if not self.current_class:
            messagebox.showwarning("提示", "请先选择班级。")
            return
        dialog = StudentDialog(self.root, title="添加学员")
        self.root.wait_window(dialog.top)
        if dialog.result:
            name, username = dialog.result
            username = self.sanitize_username(username)
            if not username:
                messagebox.showwarning("提示", "用户名不能为空。")
                return
            self.students[name] = username
            self._update_student_tree()
            self.save_data()
            self.log(f"添加学员 {name}")

    def edit_student(self, _event) -> None:
        if not self.current_class:
            return
        sel = self.student_tree.selection()
        if not sel:
            return
        item = sel[0]
        name, username = self.student_tree.item(item, "values")
        dialog = StudentDialog(self.root, title=f"编辑学员 - {name}", initial=(name, username))
        self.root.wait_window(dialog.top)
        if dialog.result:
            new_name, new_username = dialog.result
            new_username = self.sanitize_username(new_username)
            if not new_username:
                messagebox.showwarning("提示", "用户名不能为空。")
                return
            if new_name != name:
                self.students.pop(name, None)
            self.students[new_name] = new_username
            self._update_student_tree()
            self.save_data()
            self.log(f"更新学员 {new_name}")

    def delete_student(self) -> None:
        if not self.current_class:
            messagebox.showwarning("提示", "请先选择班级。")
            return
        sel = self.student_tree.selection()
        if not sel:
            messagebox.showwarning("提示", "请先在列表中选中学员。")
            return
        item = sel[0]
        name, _ = self.student_tree.item(item, "values")
        if not messagebox.askyesno("确认删除", f"确定删除学员 '{name}'?"):
            return
        self.students.pop(name, None)
        self._update_student_tree()
        self.save_data()
        self.log(f"删除学员 {name}")

    def import_students_paste(self) -> None:
        if not self.current_class:
            messagebox.showwarning("提示", "请先选择班级。")
            return
        template = "19 ChengYu Li：ChengyuL16\n20 Kyle Zhang:e7Rook"
        content = multiline_input("粘贴学员名单", "请输入内容 (姓名:用户名):", template)
        if content is None:
            return
        parsed = self.parse_students_list(content)
        if not parsed:
            messagebox.showwarning("提示", "未能解析学员信息，请检查格式。")
            return
        self.students.update(parsed)
        self._update_student_tree()
        self.save_data()
        messagebox.showinfo("导入完成", f"成功导入 {len(parsed)} 名学员。")
        self.log("学员列表已更新")

    def import_students_file(self) -> None:
        if not self.current_class:
            messagebox.showwarning("提示", "请先选择班级。")
            return
        file_path = filedialog.askopenfilename(
            title="选择学员名单",
            filetypes=[("文本文件", "*.txt *.csv"), ("所有文件", "*.*")],
        )
        if not file_path:
            return
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                content = f.read()
        except Exception as exc:
            messagebox.showerror("读取失败", f"无法读取文件: {exc}")
            return
        parsed = self.parse_students_list(content)
        if not parsed:
            messagebox.showwarning("提示", "未能解析学员信息，请检查文本格式。")
            return
        self.students.update(parsed)
        self._update_student_tree()
        self.save_data()
        messagebox.showinfo("导入完成", f"成功导入 {len(parsed)} 名学员。")
        self.log("学员列表已更新")

    def parse_students_list(self, content: str) -> Dict[str, str]:
        students: Dict[str, str] = {}
        for raw in content.splitlines():
            line = raw.strip()
            if not line:
                continue
            line = re.sub(r"^[\s\u3000]*\d+[\.、:：\)\-\s]*", "", line)
            line = line.strip()
            for sep in ("->", "：", ":", "-", "—"):
                if sep in line:
                    left, right = line.rsplit(sep, 1)
                    name = left.strip()
                    username = right.strip()
                    break
            else:
                items = line.split()
                if len(items) >= 2:
                    name = " ".join(items[:-1]).strip()
                    username = items[-1].strip()
                else:
                    continue
            if name and username:
                # 处理姓名或用户名中继续出现冒号的情况
                if any(sep in username for sep in (":", "：")):
                    parts = re.split(r"[:：]", username)
                    username_candidate = parts[-1].strip()
                    name_tail = " ".join(p.strip() for p in parts[:-1] if p.strip())
                    if username_candidate:
                        username = username_candidate
                        if name_tail:
                            name = f"{name} {name_tail}".strip()

                username = self.sanitize_username(username)
                if not username:
                    continue
                students[name] = username
        return students

    # ------------------------------------------------------------------
    # 对阵解析
    # ------------------------------------------------------------------
    def paste_pairings(self) -> None:
        try:
            text = self.root.clipboard_get()
        except tk.TclError:
            messagebox.showwarning("提示", "剪贴板没有内容。")
            return
        self.pairings_text.delete("1.0", tk.END)
        self.pairings_text.insert("1.0", text)
        self.log("对阵表已粘贴")

    def clear_pairings(self) -> None:
        self.pairings_text.delete("1.0", tk.END)
        self.result_text.delete("1.0", tk.END)
        self.parsed_pairings = []
        self.progress["value"] = 0
        self.log("对阵表已清空")

    def parse_pairings(self) -> None:
        if not self.students:
            messagebox.showwarning("提示", "当前班级没有学员信息。")
            return
        text = self.pairings_text.get("1.0", tk.END).strip()
        if not text:
            messagebox.showwarning("提示", "请输入对阵表内容。")
            return

        try:
            self.parsed_pairings = self.parse_pairings_content(text)
        except Exception as exc:
            messagebox.showerror("解析失败", str(exc))
            return

        if not self.parsed_pairings:
            messagebox.showwarning("提示", "未能解析出对阵，请检查格式。")
            return

        output_lines = ["解析结果", "=" * 40, ""]
        matched = 0
        for idx, pairing in enumerate(self.parsed_pairings, 1):
            white = pairing.get("white", "未知")
            black = pairing.get("black", "未知")
            white_user = pairing.get("white_username", "未找到")
            black_user = pairing.get("black_username", "未找到")
            ok = white_user != "未找到" and black_user != "未找到"
            if ok:
                matched += 1
            status = "✓" if ok else "✗"
            output_lines.append(f"{status} 第{idx}局")
            output_lines.append(f"  白方: {white} → {white_user}")
            output_lines.append(f"  黑方: {black} → {black_user}")
            output_lines.append("-" * 32)
        output_lines.append("")
        output_lines.append(f"共 {len(self.parsed_pairings)} 局，对应用户名 {matched} 局匹配成功")

        self.result_text.delete("1.0", tk.END)
        self.result_text.insert("1.0", "\n".join(output_lines))
        self.log("对阵解析完成")

    def parse_pairings_content(self, content: str) -> List[Dict[str, str]]:
        pairings: List[Dict[str, str]] = []
        for raw in content.splitlines():
            line = raw.strip()
            if not line:
                continue
            line = re.sub(r"^\s*\d+[\.、:：\)\-–—\s]*", "", line)
            line = line.strip()
            if not line:
                continue
            pairing = self._extract_pairing(line)
            if pairing:
                pairings.append(pairing)
        return pairings

    def _extract_pairing(self, line: str) -> Optional[Dict[str, str]]:
        patterns = [
            r"(.+?)\s+[vs对战VS]\s+(.+?)(?:\s|$)",
            r"(.+?)\s+-\s+(.+?)(?:\s|$)",
            r"(.+?)\s+对\s+(.+?)(?:\s|$)",
            r"(.+?)\s*[:：]\s*(.+?)(?:\s|$)",
            r"^(.+?)\s+(.+?)$",
        ]
        for pattern in patterns:
            match = re.search(pattern, line, re.IGNORECASE)
            if match:
                white = match.group(1).strip()
                black = match.group(2).strip()
                if white and black and not white.isdigit() and not black.isdigit():
                    return {
                        "white": white,
                        "black": black,
                        "white_username": self.find_username(white),
                        "black_username": self.find_username(black),
                    }
        return None

    def find_username(self, name: str) -> str:
        if not name:
            return "未找到"
        name = name.strip()
        if name in self.students:
            return self.students[name]
        lower = name.lower()
        for real_name, username in self.students.items():
            if real_name.lower() == lower:
                return username
        for real_name, username in self.students.items():
            if lower in real_name.lower():
                return username
        for real_name, username in self.students.items():
            if real_name.lower() in lower:
                return username
        parts = lower.split()
        first = parts[0] if parts else lower
        for real_name, username in self.students.items():
            real_first = real_name.lower().split()[0] if real_name.split() else real_name.lower()
            if first == real_first and len(first) > 1:
                return username
        filtered = [p for p in parts if len(p) > 1]
        for real_name, username in self.students.items():
            real_parts = [seg.lower() for seg in real_name.split() if len(seg) > 1]
            if set(filtered) & set(real_parts):
                return username
        clean = re.sub(r"[^\w]", "", lower)
        for real_name, username in self.students.items():
            if clean and clean == re.sub(r"[^\w]", "", real_name.lower()):
                return username
        return "未找到"

    # ------------------------------------------------------------------
    # 下载
    # ------------------------------------------------------------------
    def download_games(self) -> None:
        if not self.parsed_pairings:
            messagebox.showwarning("提示", "请先解析对阵表。")
            return
        round_name = simple_input("下载棋谱", "请输入轮次名称 (例如 Round1):")
        if round_name is None:
            return
        round_name = round_name.strip() or "Round"

        folder = self._prepare_download_folder(round_name)
        total = len(self.parsed_pairings)
        self.progress["maximum"] = total
        success = 0
        failed: List[str] = []

        for idx, pairing in enumerate(self.parsed_pairings, 1):
            self.progress["value"] = idx
            white_user = pairing.get("white_username", "未找到")
            black_user = pairing.get("black_username", "未找到")
            white_name = pairing.get("white", "")
            black_name = pairing.get("black", "")

            self.log(f"下载 {idx}/{total}: {white_name} vs {black_name}")

            if white_user == "未找到" or black_user == "未找到":
                failed.append(f"第{idx}局 {white_name} vs {black_name} (用户名未匹配)")
                continue

            ok, detail = self.download_pairing_games(
                white_user,
                black_user,
                white_name,
                black_name,
                folder,
            )
            if ok:
                success += 1
            else:
                failed.append(f"第{idx}局 {white_name} vs {black_name} ({detail})")

        self.progress["value"] = total
        report_path = folder / "下载报告.txt"
        with report_path.open("w", encoding="utf-8") as f:
            f.write(f"下载报告 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 60 + "\n")
            f.write(f"班级: {self.current_class}\n")
            f.write(f"轮次: {round_name}\n")
            f.write(f"对阵总数: {total}\n")
            f.write(f"成功下载: {success}\n")
            f.write(f"失败: {len(failed)}\n\n")
            if failed:
                f.write("失败详情:\n")
                for item in failed:
                    f.write(f"- {item}\n")

        messagebox.showinfo(
            "下载完成",
            f"总计 {total} 局\n成功 {success}\n失败 {len(failed)}\n保存目录: {folder}",
        )

        try:
            if platform.system() == "Windows":
                os.startfile(folder)
            elif platform.system() == "Darwin":
                subprocess.run(["open", str(folder)], check=False)
            else:
                subprocess.run(["xdg-open", str(folder)], check=False)
        except Exception:
            pass

        self.log("下载完成")

    def _prepare_download_folder(self, round_name: str) -> Path:
        safe_class = re.sub(r"[^0-9A-Za-z\-_]", "_", self.current_class) or "class"
        safe_round = re.sub(r"[^0-9A-Za-z\-_]", "_", round_name) or "Round"
        folder_name = f"{safe_class}-round-{safe_round}"
        folder = self.base_dir / folder_name
        if folder.exists():
            suffix = datetime.now().strftime("_%Y%m%d_%H%M%S")
            folder = self.base_dir / f"{folder_name}{suffix}"
        folder.mkdir(parents=True, exist_ok=True)
        return folder

    def download_pairing_games(
        self,
        white_username: str,
        black_username: str,
        white_display: str,
        black_display: str,
        folder: Path,
    ) -> tuple[bool, str]:
        white_api = self.sanitize_username(white_username).lower()
        black_api = self.sanitize_username(black_username).lower()

        if not white_api or not black_api:
            return False, "用户名无效"

        archives, white_err = self.get_player_archives(white_api)
        if not archives:
            archives, black_err = self.get_player_archives(black_api)
            if not archives:
                reason = white_err or black_err or "未找到棋谱归档"
                return False, reason
        else:
            black_err = ""

        matched_games: List[Dict[str, Any]] = []
        last_error = ""
        cutoff = datetime.utcnow() - timedelta(days=self.recent_days)
        considered_archives = [url for url in archives if self.is_archive_recent(url)]
        if not considered_archives:
            considered_archives = archives[-self.archive_month_limit :]

        for archive_url in reversed(considered_archives):
            games, err = self.get_archive_games(archive_url)
            if err:
                last_error = err
                continue
            for game in games:
                game_white = game.get("white", {}).get("username", "").lower()
                game_black = game.get("black", {}).get("username", "").lower()
                if {game_white, game_black} != {white_api, black_api}:
                    continue
                game_time = self.extract_game_time(game)
                if game_time and game_time < cutoff:
                    continue
                matched_games.append(game)

        if not matched_games:
            if white_err and not archives:
                return False, white_err
            if black_err and not archives:
                return False, black_err
            if last_error:
                return False, last_error
            return False, f"未找到双方在最近{self.recent_days}天的对局"

        safe_white = re.sub(r"[^0-9A-Za-z\-_]", "_", white_display) or "white"
        safe_black = re.sub(r"[^0-9A-Za-z\-_]", "_", black_display) or "black"

        saved_any = False
        for idx, game in enumerate(matched_games, 1):
            pgn = game.get("pgn")
            if not pgn:
                continue
            game_time = self.extract_game_time(game)
            if game_time:
                dt_str = game_time.strftime("%Y%m%d_%H%M")
            else:
                dt_str = datetime.utcnow().strftime("%Y%m%d_%H%M")
            filename = f"{safe_white}_vs_{safe_black}_{dt_str}_{idx}.pgn"
            filename = re.sub(r"[<>:\"/\\|?*]", "_", filename)
            filepath = folder / filename
            try:
                with filepath.open("w", encoding="utf-8") as f:
                    f.write(pgn)
                    if not pgn.endswith("\n"):
                        f.write("\n")
            except OSError:
                continue
            saved_any = True
            time.sleep(0.2)

        if not saved_any:
            return False, "找到对局但保存失败"

        return True, f"保存 {len(matched_games)} 局"

    def get_player_archives(self, username: str) -> tuple[List[str], str]:
        if not username:
            return [], "用户名缺失"
        url = f"https://api.chess.com/pub/player/{username}/games/archives"
        try:
            resp = self.http.get(url, timeout=10)
        except requests.RequestException as exc:
            return [], f"网络错误: {exc}"
        if resp.status_code == 404:
            return [], "用户不存在或无公开棋谱"
        if resp.status_code == 403:
            return [], "HTTP 403 (访问被拒绝)"
        if resp.status_code != 200:
            return [], f"HTTP {resp.status_code}"
        data = resp.json()
        archives = data.get("archives", [])
        return archives if isinstance(archives, list) else [], ""

    def get_archive_games(self, archive_url: str) -> tuple[List[Dict[str, Any]], str]:
        if archive_url in self.archive_cache:
            return self.archive_cache[archive_url], ""
        try:
            resp = self.http.get(archive_url, timeout=15)
        except requests.Timeout:
            return [], "请求超时"
        except requests.RequestException as exc:
            return [], f"网络错误: {exc}"
        if resp.status_code != 200:
            if resp.status_code == 403:
                return [], "HTTP 403 (访问被拒绝)"
            return [], f"HTTP {resp.status_code}"
        data = resp.json()
        games = data.get("games", []) if isinstance(data, dict) else []
        if isinstance(games, list):
            self.archive_cache[archive_url] = games
            return games, ""
        return [], "数据格式错误"

    def is_archive_recent(self, archive_url: str) -> bool:
        match = re.search(r"/(\d{4})/(\d{2})$", archive_url)
        if not match:
            return True
        year, month = map(int, match.groups())
        now = datetime.utcnow()
        current_total = now.year * 12 + now.month
        archive_total = year * 12 + month
        return (current_total - archive_total) <= 2

    def extract_game_time(self, game: Dict[str, Any]) -> Optional[datetime]:
        timestamp = (
            game.get("end_time")
            or game.get("finish_time")
            or game.get("start_time")
            or game.get("last_activity")
        )
        if isinstance(timestamp, (int, float)):
            try:
                return datetime.utcfromtimestamp(timestamp)
            except (OverflowError, OSError, ValueError):
                return None
        if isinstance(timestamp, str):
            for fmt in ("%Y-%m-%dT%H:%M:%S.%fZ", "%Y-%m-%dT%H:%M:%SZ"):
                try:
                    return datetime.strptime(timestamp, fmt)
                except ValueError:
                    continue
        return None


# ----------------------------------------------------------------------
# 辅助对话框
# ----------------------------------------------------------------------
class StudentDialog:
    def __init__(self, parent: tk.Tk, title: str, initial: Optional[tuple[str, str]] = None):
        self.result: Optional[tuple[str, str]] = None
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.transient(parent)
        self.top.grab_set()
        self.top.resizable(False, False)

        ttk.Label(self.top, text="姓名:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.name_var = tk.StringVar(value=initial[0] if initial else "")
        name_entry = ttk.Entry(self.top, textvariable=self.name_var, width=28)
        name_entry.grid(row=0, column=1, padx=10, pady=10)
        name_entry.focus()

        ttk.Label(self.top, text="Chess.com 用户名:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
        self.username_var = tk.StringVar(value=initial[1] if initial else "")
        ttk.Entry(self.top, textvariable=self.username_var, width=28).grid(row=1, column=1, padx=10, pady=10)

        btn_frame = ttk.Frame(self.top)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=(0, 10))

        ttk.Button(btn_frame, text="确定", width=10, command=self._on_ok).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_frame, text="取消", width=10, command=self.top.destroy).pack(side=tk.LEFT, padx=6)

        self.top.bind("<Return>", lambda _e: self._on_ok())
        self.top.bind("<Escape>", lambda _e: self.top.destroy())

    def _on_ok(self) -> None:
        name = self.name_var.get().strip()
        username = self.username_var.get().strip()
        if not name or not username:
            messagebox.showwarning("提示", "请填写姓名和用户名。", parent=self.top)
            return
        self.result = (name, username)
        self.top.destroy()


def simple_input(title: str, prompt: str) -> Optional[str]:
    dialog = simpledialog.askstring(title, prompt)
    if dialog is None:
        return None
    return dialog.strip()


def multiline_input(title: str, prompt: str, default: str = "") -> Optional[str]:
    dialog = tk.Toplevel()
    dialog.title(title)
    dialog.transient()
    dialog.grab_set()
    dialog.resizable(True, True)

    ttk.Label(dialog, text=prompt).pack(padx=10, pady=8, anchor=tk.W)
    text_widget = scrolledtext.ScrolledText(dialog, width=60, height=12, wrap=tk.WORD)
    text_widget.pack(padx=10, pady=(0, 10), fill=tk.BOTH, expand=True)
    if default:
        text_widget.insert("1.0", default)

    result: Dict[str, Optional[str]] = {"value": None}

    def on_ok() -> None:
        result["value"] = text_widget.get("1.0", tk.END).strip()
        dialog.destroy()

    def on_cancel() -> None:
        dialog.destroy()

    btn_frame = ttk.Frame(dialog)
    btn_frame.pack(pady=(0, 10))

    ttk.Button(btn_frame, text="确定", width=10, command=on_ok).pack(side=tk.LEFT, padx=6)
    ttk.Button(btn_frame, text="取消", width=10, command=on_cancel).pack(side=tk.LEFT, padx=6)

    dialog.bind("<Return>", lambda _e: on_ok())
    dialog.bind("<Escape>", lambda _e: on_cancel())

    dialog.wait_window()
    return result["value"]


def main() -> None:
    root = tk.Tk()
    app = TeacherChessApp(root)

    def on_close() -> None:
        try:
            app.save_data()
        finally:
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
