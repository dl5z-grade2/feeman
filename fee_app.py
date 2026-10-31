# -*- coding: utf-8 -*-
"""饭费管理系统 图形界面（Tkinter）。

基于 fee_core.py 的核心模型，提供：
  * .cfee 工程的新建 / 打开 / 保存 / 另存为（支持命令行直接打开工程文件）
  * 核算月份、全月单价、每周休息日与每日单独单价的设置
  * 名单的录入、修改、删除，以及 Excel / CSV / 文本粘贴导入
  * 每日不订（快捷标记“今天不订”、整月都不订）
  * 费用汇总与 Excel 核算表导出（班级 / 姓名 / 身份证号 / 费用）

运行：python fee_app.py [工程文件.cfee]
"""

from __future__ import annotations

import calendar
import os
import re
import sys
import tkinter as tk
from datetime import date
from tkinter import filedialog, messagebox, simpledialog, ttk
from typing import Dict, List, Optional

# 嵌入式 Python 用 python311._pth 限制 sys.path，脚本目录不会自动加入；
# 显式加入，确保能 import 同目录的 fee_core。
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import fee_core
from fee_core import Person, Project

APP_TITLE = "饭费管理系统"
WEEKDAYS = ["一", "二", "三", "四", "五", "六", "日"]
WEEKDAYS_CN = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
MONEY_FMT = "{:.2f}"


# ---------------------------------------------------------------------------
# 通用小工具
# ---------------------------------------------------------------------------

def fmt_money(value: float) -> str:
    return MONEY_FMT.format(float(value or 0))


def weekday_name(year: int, month: int, day: int) -> str:
    return "周" + WEEKDAYS[calendar.weekday(year, month, day)]


def safe_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", name).strip() or "未命名工程"


def center_window(win: tk.Misc, parent: Optional[tk.Misc] = None) -> None:
    """把窗口居中到屏幕或父窗口。"""
    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()
    if parent is not None and parent.winfo_width() > 1:
        x = parent.winfo_rootx() + (parent.winfo_width() - w) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - h) // 3
    else:
        x = (win.winfo_screenwidth() - w) // 2
        y = (win.winfo_screenheight() - h) // 3
    win.geometry(f"+{max(x, 0)}+{max(y, 0)}")


def parse_roster_text(text: str) -> tuple[List[Person], List[str]]:
    """解析“一行一人”的名单文本。

    每行格式：学号 姓名 [身份证] [性别]，列之间用空格、制表符、逗号或分号分隔。
    """
    persons: List[Person] = []
    warnings: List[str] = []
    seen = set()
    for i, raw in enumerate(text.splitlines(), start=1):
        line = raw.strip()
        if not line:
            continue
        parts = [p for p in re.split(r"[\s,，;；\t]+", line) if p != ""]
        if len(parts) < 2:
            warnings.append(f"第 {i} 行：至少需要“学号 姓名”两列，已跳过")
            continue
        sid, name = parts[0], parts[1]
        if sid in ("学号", "编号") or name == "姓名":
            continue  # 忽略表头行
        idc = parts[2] if len(parts) > 2 else ""
        gender = parts[3] if len(parts) > 3 else ""
        if sid in seen:
            warnings.append(f"第 {i} 行：学号 {sid} 重复，已跳过")
            continue
        ok, reason = fee_core.validate_id_card(idc)
        if not ok:
            warnings.append(f"第 {i} 行：{name} 身份证无效（{reason}），已跳过")
            continue
        if not gender and idc:
            _, g = fee_core.id_birth_gender(idc)
            gender = g or ""
        if gender and gender not in ("男", "女"):
            warnings.append(f"第 {i} 行：{name} 性别“{gender}”非男/女，已置空")
            gender = ""
        seen.add(sid)
        persons.append(Person(sid=sid, name=name, id_card=idc.upper() if idc else "", gender=gender))
    return persons, warnings


# ---------------------------------------------------------------------------
# 人员编辑对话框
# ---------------------------------------------------------------------------

class PersonDialog(tk.Toplevel):
    """添加 / 编辑单个人员。关闭后通过 .result 取回 Person 或 None。"""

    def __init__(self, master: tk.Misc, title: str = "添加人员",
                 person: Optional[Person] = None, existing_sids=None):
        super().__init__(master)
        self.title(title)
        self.resizable(False, False)
        self.result: Optional[Person] = None
        self._existing = set(existing_sids or [])
        self._original_sid = person.sid if person else None

        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)

        self.sid_var = tk.StringVar(value=person.sid if person else "")
        self.name_var = tk.StringVar(value=person.name if person else "")
        self.id_var = tk.StringVar(value=person.id_card if person else "")
        self.gender_var = tk.StringVar(value=person.gender if person else "")

        def add_row(row: int, label: str, widget: tk.Widget, required: bool = False):
            text = label + (" *" if required else "")
            ttk.Label(body, text=text).grid(row=row, column=0, sticky="w", pady=6, padx=(0, 12))
            widget.grid(row=row, column=1, sticky="ew", pady=6)

        self.sid_entry = ttk.Entry(body, textvariable=self.sid_var, width=34)
        add_row(0, "学号", self.sid_entry, True)
        self.name_entry = ttk.Entry(body, textvariable=self.name_var, width=34)
        add_row(1, "姓名", self.name_entry, True)
        self.id_entry = ttk.Entry(body, textvariable=self.id_var, width=34)
        add_row(2, "身份证号", self.id_entry)
        self.gender_box = ttk.Combobox(body, textvariable=self.gender_var, width=32,
                                       values=["", "男", "女"], state="readonly")
        add_row(3, "性别", self.gender_box)

        ttk.Label(body, text="身份证可留空；填写时会校验 18 位校验码，并可自动识别性别。",
                  style="Hint.TLabel").grid(row=4, column=0, columnspan=2, sticky="w", pady=(4, 10))

        btns = ttk.Frame(body)
        btns.grid(row=5, column=0, columnspan=2, sticky="e")
        ttk.Button(btns, text="确定", style="Accent.TButton", command=self._ok, width=10).pack(side="left", padx=4)
        ttk.Button(btns, text="取消", command=self._cancel, width=10).pack(side="left")

        body.columnconfigure(1, weight=1)
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self._cancel())
        self.transient(master)
        center_window(self, master)
        self.wait_visibility()
        self.grab_set()
        self.sid_entry.focus_set()
        self.wait_window(self)

    def _ok(self) -> None:
        sid = self.sid_var.get().strip()
        name = self.name_var.get().strip()
        idc = self.id_var.get().strip().upper()
        gender = self.gender_var.get().strip()
        if not sid:
            messagebox.showwarning("提示", "请填写学号。", parent=self)
            self.sid_entry.focus_set()
            return
        if not name:
            messagebox.showwarning("提示", "请填写姓名。", parent=self)
            self.name_entry.focus_set()
            return
        if sid != self._original_sid and sid in self._existing:
            messagebox.showwarning("提示", f"学号 {sid} 已存在，请勿重复添加。", parent=self)
            self.sid_entry.focus_set()
            return
        ok, reason = fee_core.validate_id_card(idc)
        if not ok:
            messagebox.showwarning("提示", "身份证号无效：" + reason, parent=self)
            self.id_entry.focus_set()
            return
        if not gender and idc:
            _, g = fee_core.id_birth_gender(idc)
            gender = g or ""
        self.result = Person(sid=sid, name=name, id_card=idc, gender=gender)
        self.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.destroy()


# ---------------------------------------------------------------------------
# 文本粘贴导入对话框
# ---------------------------------------------------------------------------

class TextImportDialog(tk.Toplevel):
    """粘贴多行名单文本并解析。"""

    def __init__(self, master: tk.Misc):
        super().__init__(master)
        self.title("从文本粘贴导入名单")
        self.resizable(True, True)
        self.result: Optional[List[Person]] = None
        self.warnings: List[str] = []

        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="每行一人，列之间用空格 / 制表符 / 逗号分隔：学号 姓名 [身份证] [性别]",
                  style="Hint.TLabel").pack(anchor="w", pady=(0, 8))

        text_frame = ttk.Frame(body)
        text_frame.pack(fill="both", expand=True)
        self.text = tk.Text(text_frame, width=72, height=14, wrap="none", undo=True)
        yscroll = ttk.Scrollbar(text_frame, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=yscroll.set)
        self.text.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")

        btns = ttk.Frame(body)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Label(btns, text="示例：20240001 张三 110101199003071234 男",
                  style="Hint.TLabel").pack(side="left")
        ttk.Button(btns, text="取消", command=self.destroy, width=10).pack(side="right", padx=4)
        ttk.Button(btns, text="导入", style="Accent.TButton", command=self._ok, width=10).pack(side="right")

        self.transient(master)
        center_window(self, master)
        self.wait_visibility()
        self.grab_set()
        self.text.focus_set()
        self.wait_window(self)

    def _ok(self) -> None:
        persons, warnings = parse_roster_text(self.text.get("1.0", "end"))
        if not persons:
            messagebox.showwarning("提示", "没有解析到有效的人员记录。\n\n" + "\n".join(warnings[:10]),
                                   parent=self)
            return
        self.result = persons
        self.warnings = warnings
        self.destroy()


SAMPLE_UNORDERED = """1. 苏煜博，10月不订午餐
2. 44号范付璐10月不订午餐、晚餐、周六餐
3. 41号陈心怡不订所有餐
4. 43号赵思琪10月不订晚餐"""


class UnorderedImportDialog(tk.Toplevel):
    """粘贴“本月不订”的自由格式文本，解析后预览再应用。"""

    def __init__(self, master: tk.Misc, proj: Project):
        super().__init__(master)
        self.title("导入本月不订（自由文本）")
        self.geometry("820x620")
        self.proj = proj
        self.result: Optional[List[Dict]] = None
        self.warnings: List[str] = []

        body = ttk.Frame(self, padding=14)
        body.pack(fill="both", expand=True)

        ttk.Label(
            body,
            text="每行一人，格式不限。示例：44号范付璐10月不订午餐、晚餐、周六餐；"
                 "陈心怡不订所有餐；赵思琪只订午餐",
            style="Hint.TLabel", wraplength=780, justify="left",
        ).pack(anchor="w", pady=(0, 8))

        # 数据中只保留这些餐别的不订记录；未勾选的餐别一律视为照常订餐
        keep_box = ttk.LabelFrame(
            body, text="只记录这些餐别的不订（未勾选的餐别一律视为照常订餐，文本里提到也不记录）", padding=8)
        keep_box.pack(fill="x", pady=(0, 8))
        self.meal_vars: Dict[str, tk.BooleanVar] = {}
        for meal in fee_core._ALL_MEALS:
            var = tk.BooleanVar(value=True)
            self.meal_vars[meal] = var
            ttk.Checkbutton(keep_box, text=meal, variable=var,
                            command=self._on_keep_changed).pack(side="left", padx=(0, 16))
        self.keep_hint_var = tk.StringVar()
        ttk.Label(keep_box, textvariable=self.keep_hint_var, style="Hint.TLabel").pack(side="left", padx=(8, 0))

        self.replace_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(body, text="覆盖这些人员的原有逐日不订记录（默认与原有记录合并）",
                        variable=self.replace_var).pack(anchor="w", pady=(0, 8))

        panes = ttk.Panedwindow(body, orient="horizontal")
        panes.pack(fill="both", expand=True)

        left = ttk.Frame(panes, padding=(0, 0, 8, 0))
        text_frame = ttk.Frame(left)
        text_frame.pack(fill="both", expand=True)
        self.text = tk.Text(text_frame, width=46, height=15, wrap="word", undo=True)
        yscroll = ttk.Scrollbar(text_frame, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=yscroll.set)
        self.text.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")
        self.text.insert("1.0", SAMPLE_UNORDERED)
        left_btns = ttk.Frame(left)
        left_btns.pack(fill="x", pady=(8, 0))
        ttk.Button(left_btns, text="从文件读取...", command=self._load_file).pack(side="left")
        ttk.Button(left_btns, text="解析预览", style="Accent.TButton",
                   command=self._preview).pack(side="left", padx=(8, 0))
        panes.add(left, weight=3)

        right = ttk.Frame(panes, padding=(8, 0, 0, 0))
        ttk.Label(right, text="解析结果").pack(anchor="w")
        tree_wrap = ttk.Frame(right)
        tree_wrap.pack(fill="both", expand=True, pady=(6, 0))
        columns = ("name", "meals", "days")
        self.preview_tree = ttk.Treeview(tree_wrap, columns=columns, show="headings", height=12)
        for c, (t, w) in {"name": ("姓名", 90), "meals": ("不订餐别", 130), "days": ("不订天数", 80)}.items():
            self.preview_tree.heading(c, text=t)
            self.preview_tree.column(c, width=w, anchor="center" if c != "meals" else "w")
        ys = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.preview_tree.yview)
        self.preview_tree.configure(yscrollcommand=ys.set)
        self.preview_tree.pack(side="left", fill="both", expand=True)
        ys.pack(side="right", fill="y")
        panes.add(right, weight=2)

        self.info_var = tk.StringVar(value="输入文本后点击“解析预览”。")
        ttk.Label(body, textvariable=self.info_var, style="Hint.TLabel").pack(anchor="w", pady=(8, 0))

        btns = ttk.Frame(body)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="取消", command=self.destroy, width=10).pack(side="right", padx=4)
        ttk.Button(btns, text="应用到工程", style="Accent.TButton",
                   command=self._apply, width=12).pack(side="right")

        self._on_keep_changed()

        self.transient(master)
        center_window(self, master)
        self.wait_visibility()
        self.grab_set()
        self.text.focus_set()
        self.wait_window(self)

    # ---- 只记录哪些餐别的不订 ----
    def allowed_meals(self) -> List[str]:
        """勾选的餐别：只有这些餐别会写进不订数据。"""
        return [m for m in fee_core._ALL_MEALS if self.meal_vars[m].get()]

    def dropped_meals(self) -> List[str]:
        """未勾选、即视为照常订餐的餐别。"""
        return [m for m in fee_core._ALL_MEALS if not self.meal_vars[m].get()]

    def _on_keep_changed(self) -> None:
        keep = self.allowed_meals()
        drop = self.dropped_meals()
        if not drop:
            self.keep_hint_var.set("记录午餐 / 晚餐 / 周六餐的全部不订。")
        elif not keep:
            self.keep_hint_var.set("未选择任何餐别：不会写入任何不订记录。")
        else:
            self.keep_hint_var.set(
                "只记录：" + "、".join(keep) + "；" + "、".join(drop) + " 视为照常订餐")

    def _load_file(self) -> None:
        path = filedialog.askopenfilename(
            parent=self, title="选择文本文件",
            filetypes=[("文本文件", "*.txt *.csv *.md"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8-sig") as f:
                content = f.read()
        except UnicodeDecodeError:
            try:
                with open(path, "r", encoding="gbk") as f:
                    content = f.read()
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror("读取失败", str(exc), parent=self)
                return
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("读取失败", str(exc), parent=self)
            return
        self.text.delete("1.0", "end")
        self.text.insert("1.0", content)

    def _parse(self):
        keep = self.allowed_meals()
        return fee_core.parse_unordered_text(
            self.proj,
            self.text.get("1.0", "end"),
            # 只有勾选的餐别会写入不订数据；未勾选视为照常订餐
            allowed_meals=keep,
            # 行内没写餐别时，按勾选的餐别处理（没勾就不推断）
            default_unordered=keep,
        )

    def _preview(self) -> bool:
        results, warnings = self._parse()
        self.preview_tree.delete(*self.preview_tree.get_children())
        for item in results:
            self.preview_tree.insert("", "end", values=(item["name"], item["note"] or "—", len(item["days"])))
        tip = f"识别到 {len(results)} 人"
        drop = self.dropped_meals()
        if drop:
            tip += "（不记录：" + "、".join(drop) + "）"
        if warnings:
            tip += f"，另有 {len(warnings)} 行未处理"
        self.info_var.set(tip + ("；点击“应用到工程”写入不订记录。" if results else "。"))
        self._warnings = warnings
        return bool(results)

    def _apply(self) -> None:
        if not self._preview():
            messagebox.showwarning("提示", "没有可应用的结果，请检查文本。", parent=self)
            return
        keep = self.allowed_meals()
        count, warnings = fee_core.apply_unordered_text(
            self.proj,
            self.text.get("1.0", "end"),
            allowed_meals=keep,
            default_unordered=keep,
            replace=bool(self.replace_var.get()),
        )
        self.applied_count = count
        self.warnings = list(getattr(self, "_warnings", []))
        self.result = True
        self.destroy()


# ---------------------------------------------------------------------------
# 主窗口
# ---------------------------------------------------------------------------

class FeeApp(tk.Tk):
    def __init__(self, project_path: Optional[str] = None):
        super().__init__()
        self.proj = Project(year=date.today().year, month=date.today().month)
        self.path: Optional[str] = None
        self.dirty = False

        self._day_vars: Dict[int, tk.BooleanVar] = {}
        self._day_cbs: Dict[int, ttk.Checkbutton] = {}
        self._person_order: List[str] = []
        self._current_sid: Optional[str] = None
        self._loading_skip = False
        self._weekday_vars: Dict[int, tk.BooleanVar] = {}

        self.title(APP_TITLE)
        self.geometry("1140x760")
        self.minsize(980, 660)
        self._build_style()
        self._build_menu()
        self._build_header()
        self._build_notebook()
        self._build_status()
        self._bind_shortcuts()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.refresh_all()
        center_window(self)

        # 命令行传入工程文件时直接打开（不弹“是否保存”确认）
        if project_path:
            self.after(80, lambda: self.load_from_path(project_path, quiet=False))

    # ---------------------------------------------------------------- 样式
    def _build_style(self) -> None:
        style = ttk.Style(self)
        try:
            themes = style.theme_names()
            if "vista" in themes:
                style.theme_use("vista")
            elif "clam" in themes:
                style.theme_use("clam")
        except tk.TclError:
            pass
        family = "Microsoft YaHei UI"
        base = (family, 10)
        # 只给文本 / 菜单这类经典 Tk 控件设默认字体。
        # 不能用 option_add("*Font", base)：这个通配项会同时覆盖 ttk 控件的 -font，
        # 令 Count.TLabel（38pt）等自定义样式一律退回 10pt，“今日订餐”的大号数字
        # 和窗口标题都会缩小。
        self.option_add("*Text.font", base)
        self.option_add("*Menu.font", base)
        self.option_add("*TCombobox*Listbox.font", base)
        style.configure(".", font=base)
        style.configure("Treeview", rowheight=27, font=base)
        style.configure("Treeview.Heading", font=(family, 10, "bold"))
        style.configure("Title.TLabel", font=(family, 15, "bold"))
        style.configure("Big.TLabel", font=(family, 12, "bold"))
        style.configure("Hint.TLabel", foreground="#707070")
        style.configure("Money.TLabel", font=(family, 11, "bold"), foreground="#B00020")
        style.configure("Count.TLabel", font=(family, 38, "bold"), foreground="#1A7F37")
        style.configure("CountUnit.TLabel", font=(family, 13, "bold"), foreground="#1A7F37")
        style.configure("Accent.TButton", font=(family, 10, "bold"))
        style.configure("Toolbar.TFrame", padding=(0, 6))

    # ---------------------------------------------------------------- 菜单
    def _build_menu(self) -> None:
        menubar = tk.Menu(self)

        file_menu = tk.Menu(menubar, tearoff=False)
        file_menu.add_command(label="新建工程", accelerator="Ctrl+N", command=self.new_project)
        file_menu.add_command(label="打开工程...", accelerator="Ctrl+O", command=self.open_project)
        file_menu.add_separator()
        file_menu.add_command(label="保存", accelerator="Ctrl+S", command=self.save_project)
        file_menu.add_command(label="另存为...", accelerator="Ctrl+Shift+S", command=self.save_project_as)
        file_menu.add_separator()
        file_menu.add_command(label="导入名单（Excel/CSV）...", accelerator="Ctrl+I", command=self.import_roster)
        file_menu.add_command(label="从文本粘贴名单...", command=self.import_roster_text)
        file_menu.add_command(label="导入本月不订（自由文本）...", accelerator="Ctrl+Shift+I",
                              command=self.import_unordered)
        file_menu.add_command(label="下载名单模板...", command=self.export_template)
        file_menu.add_command(label="导出核算表（Excel）...", accelerator="Ctrl+E", command=self.export_settlement)
        file_menu.add_separator()
        file_menu.add_command(label="退出", accelerator="Alt+F4", command=self._on_close)
        menubar.add_cascade(label="文件", menu=file_menu)

        edit_menu = tk.Menu(menubar, tearoff=False)
        edit_menu.add_command(label="添加人员...", accelerator="Ctrl+Shift+N", command=self.add_person)
        edit_menu.add_command(label="编辑选中人员...", accelerator="F2", command=self.edit_person)
        edit_menu.add_command(label="删除选中人员", accelerator="Delete", command=self.delete_person)
        edit_menu.add_separator()
        edit_menu.add_command(label="设置选中日期单价...", command=lambda: self.edit_selected_day())
        edit_menu.add_command(label="清除选中日期设置", command=lambda: self.clear_selected_day())
        edit_menu.add_separator()
        edit_menu.add_command(label="标记/取消今天不订", accelerator="空格",
                              command=self.toggle_selected_skip)
        edit_menu.add_command(label="整月都不订", command=lambda: self._set_all_days(True))
        edit_menu.add_command(label="清除该人全部不订", command=lambda: self._set_all_days(False))
        menubar.add_cascade(label="编辑", menu=edit_menu)

        view_menu = tk.Menu(menubar, tearoff=False)
        for idx, name in enumerate(["今日订餐", "每日单价", "人员名单", "每日不订", "核算汇总"], start=1):
            view_menu.add_command(label=name, accelerator=f"Ctrl+{idx}",
                                  command=lambda i=idx - 1: self.goto_tab(i))
        view_menu.add_separator()
        view_menu.add_command(label="下一个页签", accelerator="Ctrl+Tab",
                              command=lambda: self.goto_tab((self._current_tab() + 1) % 5))
        view_menu.add_command(label="上一个页签", accelerator="Ctrl+Shift+Tab",
                              command=lambda: self.goto_tab((self._current_tab() - 1) % 5))
        view_menu.add_separator()
        view_menu.add_command(label="刷新", accelerator="F5", command=self.refresh_all)
        menubar.add_cascade(label="查看", menu=view_menu)

        help_menu = tk.Menu(menubar, tearoff=False)
        help_menu.add_command(label="使用说明", accelerator="F1", command=self.show_help)
        help_menu.add_command(label="快捷键一览", accelerator="Ctrl+/", command=self.show_shortcuts)
        help_menu.add_command(label="关于", command=self.show_about)
        menubar.add_cascade(label="帮助", menu=help_menu)

        self.config(menu=menubar)

    # ---------------------------------------------------------------- 顶部
    def _build_header(self) -> None:
        outer = ttk.Frame(self, padding=(14, 12, 14, 6))
        outer.pack(fill="x")

        ttk.Label(outer, text="饭费管理", style="Title.TLabel").grid(row=0, column=0, rowspan=3, sticky="nw", padx=(0, 22), pady=(2, 0))

        ttk.Label(outer, text="工程名称").grid(row=0, column=1, sticky="e", padx=(0, 6))
        self.title_var = tk.StringVar()
        ttk.Entry(outer, textvariable=self.title_var, width=24).grid(row=0, column=2, sticky="w")

        ttk.Label(outer, text="班级").grid(row=0, column=3, sticky="e", padx=(12, 6))
        self.class_var = tk.StringVar()
        ttk.Entry(outer, textvariable=self.class_var, width=14).grid(row=0, column=4, columnspan=2, sticky="w")

        ttk.Label(outer, text="核算月份").grid(row=0, column=6, sticky="e", padx=(14, 6))
        self.year_var = tk.StringVar()
        years = [str(y) for y in range(date.today().year - 5, date.today().year + 6)]
        ttk.Combobox(outer, textvariable=self.year_var, width=7, values=years).grid(row=0, column=7, sticky="w")
        ttk.Label(outer, text="年").grid(row=0, column=8, sticky="w", padx=(2, 4))
        self.month_var = tk.StringVar()
        ttk.Combobox(outer, textvariable=self.month_var, width=4, state="readonly",
                     values=[str(m) for m in range(1, 13)]).grid(row=0, column=9, sticky="w")
        ttk.Label(outer, text="月").grid(row=0, column=10, sticky="w", padx=(2, 6))

        ttk.Label(outer, text="全月单价").grid(row=1, column=1, sticky="e", padx=(0, 6), pady=(6, 0))
        self.price_var = tk.StringVar()
        ttk.Entry(outer, textvariable=self.price_var, width=10).grid(row=1, column=2, sticky="w", pady=(6, 0))
        ttk.Label(outer, text="元 / 天（0 表示未设置）", style="Hint.TLabel").grid(
            row=1, column=3, columnspan=8, sticky="w", pady=(6, 0))

        # 每周休息日（这些天单价为 0、不计入订餐天数）
        ttk.Label(outer, text="每周休息日").grid(row=2, column=1, sticky="e", padx=(0, 6), pady=(6, 0))
        wd_box = ttk.Frame(outer)
        wd_box.grid(row=2, column=2, columnspan=9, sticky="w", pady=(6, 0))
        for idx, name in enumerate(WEEKDAYS_CN):
            var = tk.BooleanVar(value=(idx in self.proj.skip_weekdays))
            self._weekday_vars[idx] = var
            ttk.Checkbutton(wd_box, text=name, variable=var,
                            command=self._on_weekday_toggle).pack(side="left", padx=(0, 6))
        ttk.Label(wd_box, text="（勾选表示该星期不供餐，价格 0）", style="Hint.TLabel").pack(side="left", padx=(6, 0))

        btn_box = ttk.Frame(outer)
        btn_box.grid(row=0, column=11, rowspan=3, sticky="ne", padx=(16, 0))
        ttk.Button(btn_box, text="应用设置", style="Accent.TButton", command=self.apply_header, width=11).pack(side="left", padx=4)
        ttk.Button(btn_box, text="保存工程", command=self.save_project, width=11).pack(side="left")

        outer.columnconfigure(12, weight=1)
        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=14, pady=(0, 4))

    # ---------------------------------------------------------------- 选项卡
    def _build_notebook(self) -> None:
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=14, pady=(4, 6))

        self.tab_today = ttk.Frame(self.notebook, padding=10)
        self.tab_days = ttk.Frame(self.notebook, padding=10)
        self.tab_roster = ttk.Frame(self.notebook, padding=10)
        self.tab_skip = ttk.Frame(self.notebook, padding=10)
        self.tab_summary = ttk.Frame(self.notebook, padding=10)

        self.notebook.add(self.tab_today, text="  今日订餐  ")
        self.notebook.add(self.tab_days, text="  每日单价  ")
        self.notebook.add(self.tab_roster, text="  人员名单  ")
        self.notebook.add(self.tab_skip, text="  每日不订  ")
        self.notebook.add(self.tab_summary, text="  核算汇总  ")

        self._build_tab_today()
        self._build_tab_days()
        self._build_tab_roster()
        self._build_tab_skip()
        self._build_tab_summary()
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

    # -- 今日订餐 ------------------------------------------------------
    def _build_tab_today(self) -> None:
        bar = ttk.Frame(self.tab_today)
        bar.pack(fill="x", pady=(0, 6))
        ttk.Button(bar, text="刷新", command=self.refresh_today, width=8).pack(side="left", padx=(0, 6))
        ttk.Label(bar, text="统计口径：当天供餐且未标记“不订”的人数",
                  style="Hint.TLabel").pack(side="left", padx=(6, 0))
        self.today_date_var = tk.StringVar()
        ttk.Label(bar, textvariable=self.today_date_var, style="Hint.TLabel").pack(side="right")

        # 大号统计数字
        stat = ttk.Frame(self.tab_today, padding=(0, 2, 0, 10))
        stat.pack(fill="x")
        self.today_count_var = tk.StringVar(value="0")
        ttk.Label(stat, textvariable=self.today_count_var, style="Count.TLabel").pack(side="left")
        ttk.Label(stat, text="人", style="CountUnit.TLabel").pack(side="left", padx=(6, 18), pady=(22, 0))
        self.today_sub_var = tk.StringVar()
        ttk.Label(stat, textvariable=self.today_sub_var, style="Hint.TLabel").pack(side="left", pady=(26, 0))

        # 左右两栏名单
        panes = ttk.Panedwindow(self.tab_today, orient="horizontal")
        panes.pack(fill="both", expand=True)

        left = ttk.LabelFrame(panes, text="今日订餐", padding=8)
        self.today_diner_tree = ttk.Treeview(left, columns=("sid", "name", "gender"),
                                             show="headings", selectmode="browse")
        for c, (t, w) in {"sid": ("学号", 90), "name": ("姓名", 90), "gender": ("性别", 60)}.items():
            self.today_diner_tree.heading(c, text=t)
            self.today_diner_tree.column(c, width=w, minwidth=60, anchor="center", stretch=True)
        ys = ttk.Scrollbar(left, orient="vertical", command=self.today_diner_tree.yview)
        self.today_diner_tree.configure(yscrollcommand=ys.set)
        self.today_diner_tree.pack(side="left", fill="both", expand=True)
        ys.pack(side="right", fill="y")
        panes.add(left, weight=1)

        right = ttk.LabelFrame(panes, text="今日不订", padding=8)
        self.today_absent_tree = ttk.Treeview(right, columns=("sid", "name", "reason"),
                                              show="headings", selectmode="browse")
        for c, (t, w) in {"sid": ("学号", 70), "name": ("姓名", 80), "reason": ("原因", 110)}.items():
            self.today_absent_tree.heading(c, text=t)
            self.today_absent_tree.column(c, width=w, minwidth=50,
                                          anchor="center" if c != "reason" else "w", stretch=True)
        ys2 = ttk.Scrollbar(right, orient="vertical", command=self.today_absent_tree.yview)
        self.today_absent_tree.configure(yscrollcommand=ys2.set)
        self.today_absent_tree.pack(side="left", fill="both", expand=True)
        ys2.pack(side="right", fill="y")
        self.today_absent_tree.tag_configure("closed", foreground="#909090")
        panes.add(right, weight=1)

    # -- 今日订餐 ------------------------------------------------------
    def refresh_today(self) -> None:
        """刷新“今日订餐”页：显示今天（或工程当月）的订餐人数与名单。"""
        tree = self.today_diner_tree
        absent = self.today_absent_tree
        tree.delete(*tree.get_children())
        absent.delete(*absent.get_children())

        today = date.today()
        in_month = (today.year, today.month) == (self.proj.year, self.proj.month)
        total = len(self.proj.roster)

        if not in_month:
            self.today_date_var.set(f"今天 {today.year}年{today.month}月{today.day}日（不在本工程月份）")
            self.today_count_var.set("—")
            self.today_sub_var.set(
                f"本工程为 {self.proj.year}年{self.proj.month}月，"
                f"请切换到“每日单价”页查看当月的具体情况。"
            )
            return

        day = today.day
        self.today_date_var.set(
            f"今天 {today.year}年{today.month}月{today.day}日　{weekday_name(today.year, today.month, day)}"
        )

        if self.proj.is_closed_day(day):
            self.today_count_var.set("0")
            self.today_sub_var.set(f"今天是每周休息日，不供餐。名单共 {total} 人。")
            for sid in self.proj.sorted_sids():
                p = self.proj.roster[sid]
                absent.insert("", "end", values=(p.sid, p.name, "每周休息日"), tags=("closed",))
            return

        diners = self.proj.diners_on(day)
        total_open = len(self.proj.open_days())
        self.today_count_var.set(str(len(diners)))
        price = self.proj.price_for_day(day)
        self.today_sub_var.set(
            f"共 {total} 人，今天订餐 {len(diners)} 人、不订 {total - len(diners)} 人；"
            f"今日单价 {fmt_money(price)} 元，预计收入 {fmt_money(price * len(diners))} 元"
        )
        for item in diners:
            tree.insert("", "end", values=(item["sid"], item["name"], item["gender"]))

        for item in self.proj.non_diners_on(day):
            sid = item["sid"]
            n = len(set(self.proj.skips.get(sid, set())) & set(self.proj.day_list()))
            if n >= total_open and total_open:
                reason = "整月不订"
            elif self.proj.is_skipping(sid, day):
                reason = "今日不订"
            else:
                reason = "不订"
            absent.insert("", "end", values=(sid, item["name"], reason))

    # -- 每日单价 ------------------------------------------------------
    def _build_tab_days(self) -> None:
        bar = ttk.Frame(self.tab_days)
        bar.pack(fill="x", pady=(0, 8))
        ttk.Button(bar, text="修改选中日期单价...", command=self.edit_selected_day).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="清除选中日期设置", command=self.clear_selected_day).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="全部恢复全月单价", command=self.reset_day_prices).pack(side="left", padx=(0, 6))
        self.day_info_var = tk.StringVar()
        ttk.Label(bar, textvariable=self.day_info_var, style="Hint.TLabel").pack(side="right")

        wrap = ttk.Frame(self.tab_days)
        wrap.pack(fill="both", expand=True)
        columns = ("day", "week", "price", "note")
        self.day_tree = ttk.Treeview(wrap, columns=columns, show="headings", selectmode="browse")
        headings = {"day": "日期", "week": "星期", "price": "单价（元）", "note": "设置说明"}
        widths = {"day": 120, "week": 90, "price": 140, "note": 320}
        for c in columns:
            self.day_tree.heading(c, text=headings[c])
            self.day_tree.column(c, width=widths[c], anchor="center" if c != "note" else "w", stretch=(c == "note"))
        yscroll = ttk.Scrollbar(wrap, orient="vertical", command=self.day_tree.yview)
        self.day_tree.configure(yscrollcommand=yscroll.set)
        self.day_tree.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")
        self.day_tree.tag_configure("weekend", background="#F2F2F2")
        self.day_tree.tag_configure("closed", background="#E4E4E4", foreground="#808080")
        self.day_tree.tag_configure("today", background="#FFF4CE")
        self.day_tree.bind("<Double-1>", lambda e: self.edit_selected_day(event=e))

    # -- 人员名单 ------------------------------------------------------
    def _build_tab_roster(self) -> None:
        bar = ttk.Frame(self.tab_roster)
        bar.pack(fill="x", pady=(0, 8))
        ttk.Button(bar, text="添加人员", style="Accent.TButton", command=self.add_person, width=10).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="编辑", command=self.edit_person, width=8).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="删除", command=self.delete_person, width=8).pack(side="left", padx=(0, 12))
        ttk.Button(bar, text="导入名单（Excel/CSV）...", command=self.import_roster).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="文本粘贴导入...", command=self.import_roster_text).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="下载模板", command=self.export_template).pack(side="left", padx=(0, 6))

        wrap = ttk.Frame(self.tab_roster)
        wrap.pack(fill="both", expand=True)
        columns = ("sid", "name", "gender", "id_card", "skip_days", "days", "cost")
        self.roster_tree = ttk.Treeview(wrap, columns=columns, show="headings", selectmode="browse")
        headings = {
            "sid": "学号", "name": "姓名", "gender": "性别", "id_card": "身份证号",
            "skip_days": "不订天数", "days": "订餐天数", "cost": "应缴金额（元）",
        }
        widths = {"sid": 130, "name": 110, "gender": 60, "id_card": 210, "skip_days": 90, "days": 90, "cost": 130}
        for c in columns:
            self.roster_tree.heading(c, text=headings[c])
            anchor = "w" if c in ("sid", "name", "id_card") else "center"
            self.roster_tree.column(c, width=widths[c], anchor=anchor, stretch=(c == "id_card"))
        yscroll = ttk.Scrollbar(wrap, orient="vertical", command=self.roster_tree.yview)
        self.roster_tree.configure(yscrollcommand=yscroll.set)
        self.roster_tree.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")
        self.roster_tree.tag_configure("optout", foreground="#999999")
        self.roster_tree.bind("<Double-1>", lambda e: self.edit_person())

    # -- 每日不订 ------------------------------------------------------
    def _build_tab_skip(self) -> None:
        body = ttk.Frame(self.tab_skip)
        body.pack(fill="both", expand=True)

        left = ttk.LabelFrame(body, text="人员名单", padding=8)
        left.pack(side="left", fill="y", padx=(0, 12))
        self.person_list = tk.Listbox(left, width=24, height=22, exportselection=False,
                                      activestyle="dotbox", font=("Microsoft YaHei UI", 10))
        self.person_list.pack(side="left", fill="y")
        lb_scroll = ttk.Scrollbar(left, orient="vertical", command=self.person_list.yview)
        self.person_list.configure(yscrollcommand=lb_scroll.set)
        lb_scroll.pack(side="right", fill="y")
        self.person_list.bind("<<ListboxSelect>>", self._on_person_selected)

        right = ttk.Frame(body)
        right.pack(side="left", fill="both", expand=True)

        head = ttk.Frame(right)
        head.pack(fill="x", pady=(0, 8))
        self.skip_title_var = tk.StringVar(value="请先在左侧选择人员")
        ttk.Label(head, textvariable=self.skip_title_var, style="Big.TLabel").pack(side="left")
        self.skip_month_info_var = tk.StringVar()
        ttk.Label(head, textvariable=self.skip_month_info_var, style="Hint.TLabel").pack(side="right")

        quick = ttk.Frame(right)
        quick.pack(fill="x", pady=(0, 8))
        ttk.Button(quick, text="标记今天不订", command=self.mark_today_skip, width=14).pack(side="left", padx=(0, 6))
        ttk.Button(quick, text="整月都不订", command=lambda: self._set_all_days(True), width=12).pack(side="left", padx=(0, 6))
        ttk.Button(quick, text="全部清除", command=lambda: self._set_all_days(False), width=10).pack(side="left", padx=(0, 6))
        ttk.Button(quick, text="导入本月不订...", command=self.import_unordered, width=14).pack(side="left")
        self.skip_hint_var = tk.StringVar()
        ttk.Label(quick, textvariable=self.skip_hint_var, style="Hint.TLabel").pack(side="right")

        self.day_frame = ttk.LabelFrame(right, text="不订日期（勾选表示当天不订餐）", padding=10)
        self.day_frame.pack(fill="both", expand=True)

        ttk.Label(right, text="提示：“整月都不订”就是把当月每一天都标记为不订；逐日勾选会实时保存到工程。",
                  style="Hint.TLabel").pack(anchor="w", pady=(8, 0))

    # -- 核算汇总 ------------------------------------------------------
    def _build_tab_summary(self) -> None:
        bar = ttk.Frame(self.tab_summary)
        bar.pack(fill="x", pady=(0, 8))
        ttk.Button(bar, text="刷新", command=self.refresh_summary, width=8).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="导出核算表（Excel）...", style="Accent.TButton",
                   command=self.export_settlement).pack(side="left", padx=(0, 14))

        self.skip_zero_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar, text="导出时不包括费用为 0 的人",
                        variable=self.skip_zero_var,
                        command=self._on_skip_zero_toggle).pack(side="left")

        self.total_var = tk.StringVar(value="合计：0.00 元")
        ttk.Label(bar, textvariable=self.total_var, style="Money.TLabel").pack(side="right")

        wrap = ttk.Frame(self.tab_summary)
        wrap.pack(fill="both", expand=True)
        columns = ("sid", "name", "gender", "id_card", "days", "note", "cost")
        self.summary_tree = ttk.Treeview(wrap, columns=columns, show="headings")
        headings = {
            "sid": "学号", "name": "姓名", "gender": "性别", "id_card": "身份证号",
            "days": "订餐天数", "note": "备注", "cost": "应缴金额（元）",
        }
        widths = {"sid": 120, "name": 100, "gender": 60, "id_card": 200, "days": 90, "note": 140, "cost": 130}
        for c in columns:
            self.summary_tree.heading(c, text=headings[c])
            anchor = "w" if c in ("sid", "name", "id_card") else "center"
            self.summary_tree.column(c, width=widths[c], anchor=anchor, stretch=(c == "id_card"))
        yscroll = ttk.Scrollbar(wrap, orient="vertical", command=self.summary_tree.yview)
        self.summary_tree.configure(yscrollcommand=yscroll.set)
        self.summary_tree.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")
        self.summary_tree.tag_configure("optout", foreground="#999999")

    # ---------------------------------------------------------------- 状态栏
    def _build_status(self) -> None:
        bar = ttk.Frame(self, padding=(14, 2, 14, 6))
        bar.pack(fill="x")
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(bar, textvariable=self.status_var, anchor="w").pack(side="left")
        self.file_var = tk.StringVar(value="未保存")
        ttk.Label(bar, textvariable=self.file_var, anchor="e", style="Hint.TLabel").pack(side="right")

    def _bind_shortcuts(self) -> None:
        """注册全局快捷键。

        只用**一个** ``<KeyPress>`` 通配绑定处理全部按键（含 Ctrl 组合）。
        绑定多个序列时 Tcl 的 ``break`` 会让后续绑定不执行，而
        ``<Control-KeyPress>`` 与 ``<Control-Key-2>`` 还会重复命中同一按键，
        因此统一入口自行分派。
        """
        self.bind_all("<KeyPress>", self._dispatch, add="+")

        # Ctrl+Tab 的 keysym 仍是 Tab，会先命中 Tk 内置的 <Tab>（焦点遍历）并
        # break，导致通用 <KeyPress> 收不到。必须显式绑定更具体的序列。
        self.bind_all("<Control-Tab>", self._dispatch, add="+")
        self.bind_all("<Control-Shift-Tab>", self._dispatch, add="+")

        # TNotebook 自带 <Control-Tab>/<Control-Shift-Tab> 类绑定，也会在
        # 处理器之前执行并 break（切换后还不刷新页面数据）。清空它们。
        for seq in ("<Control-Tab>", "<Control-Shift-Tab>"):
            try:
                self.notebook.bind_class("TNotebook", seq, "")
            except tk.TclError:
                pass

    # ------------------------------------------------------------------
    # 快捷键分发
    # ------------------------------------------------------------------
    def _dispatch(self, event):
        """统一快捷键入口。返回 "break" 表示已处理并阻断后续传递。"""
        if event.keysym in ("Shift_L", "Shift_R", "Control_L", "Control_R",
                            "Alt_L", "Alt_R", "Win_L", "Win_R", "Caps_Lock"):
            return None

        state = int(event.state or 0)
        ctrl = bool(state & 0x0004)      # Control
        shift = bool(state & 0x0001)     # Shift
        alt = bool(state & 0x0008)       # Alt
        key = (event.keysym or "").lower()

        # ---- Ctrl 组合 ----
        if ctrl and not alt:
            if key == "tab":
                return self._run_shortcut("prev_tab" if shift else "next_tab", event)
            if not shift and key in ("1", "2", "3", "4", "5"):
                return self._run_shortcut("tab_" + key, event)

            table = {
                "n": "add_person" if shift else "new",
                "o": "open",
                "s": "save_as" if shift else "save",
                "i": "import_unordered" if shift else "import_roster",
                "e": "export",
                "slash": "shortcuts",
            }
            name = table.get(key)
            if name:
                return self._run_shortcut(name, event)
            return None

        # ---- 无修饰键 ----
        if ctrl or alt:
            return None
        if key in ("f1", "f2", "f5", "delete", "space"):
            return self._run_shortcut(key, event)
        return None

    def _run_shortcut(self, name: str, event=None):
        """执行一个快捷键动作，并决定是否阻断事件继续传递。"""
        # 光标在输入框时只保留文件级快捷键，其余放行给控件（避免影响打字）
        if self._focus_is_text_input() and name not in ("save", "save_as"):
            return None

        if name == "next_tab":
            handler = lambda: self.goto_tab((self._current_tab() + 1) % self._tab_count())
        elif name == "prev_tab":
            handler = lambda: self.goto_tab((self._current_tab() - 1) % self._tab_count())
        elif name.startswith("tab_"):
            idx = int(name.split("_")[1]) - 1
            handler = lambda: self.goto_tab(idx)
        else:
            handler = {
                "new": self.new_project,
                "open": self.open_project,
                "save": self.save_project,
                "save_as": self.save_project_as,
                "import_roster": self.import_roster,
                "import_unordered": self.import_unordered,
                "export": self.export_settlement,
                "add_person": self.add_person,
                "shortcuts": self.show_shortcuts,
                "f1": self.show_help,
                "f2": self.edit_person,
                "f5": self.refresh_all,
                "delete": self.delete_person,
                "space": self.toggle_selected_skip,
            }.get(name)

        if handler is None:
            return None
        try:
            handler()
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("操作失败", str(exc), parent=self)
        return "break"

    def _focus_is_text_input(self) -> bool:
        """当前焦点是否在可编辑的文本控件里。"""
        try:
            w = self.focus_get()
        except (KeyError, tk.TclError):
            return False
        if w is None:
            return False
        if isinstance(w, (tk.Text, tk.Entry)):
            return True
        try:
            return w.winfo_class() in ("TEntry", "TCombobox", "TSpinbox", "Entry", "Text")
        except tk.TclError:
            return False

    def _current_tab(self) -> int:
        try:
            return self.notebook.index(self.notebook.select())
        except tk.TclError:
            return 0

    def _tab_count(self) -> int:
        try:
            return len(self.notebook.tabs()) or 1
        except tk.TclError:
            return 1

    def goto_tab(self, idx: int) -> None:
        """切换到指定页签并刷新其内容。"""
        try:
            self.notebook.select(idx)
        except tk.TclError:
            return
        self._on_tab_changed()

    def toggle_selected_skip(self) -> None:
        """空格键：给选中人员标记 / 取消“今天不订”。"""
        sid = self._selected_sid_any()
        if not sid:
            messagebox.showinfo("提示", "请先选择一名人员，再按空格。", parent=self)
            return

        today = date.today()
        if (today.year, today.month) != (self.proj.year, self.proj.month):
            messagebox.showinfo(
                "提示",
                f"今天是 {today.year}年{today.month}月{today.day}日，"
                f"不在当前工程（{self.proj.year}年{self.proj.month}月）内。",
                parent=self,
            )
            return
        day = today.day
        if self.proj.is_closed_day(day):
            messagebox.showinfo("提示", f"今天（{day}日）是每周休息日，本来就不供餐。", parent=self)
            return

        person = self.proj.roster[sid]
        days = set(self.proj.skips.get(sid, set()))
        if day in days:
            days.discard(day)
            action = "已取消今天不订"
        else:
            days.add(day)
            action = "已标记今天不订"
        if days:
            self.proj.set_skip(sid, days)
        else:
            self.proj.skips.pop(sid, None)

        self.dirty = True
        self._current_sid = sid
        self.refresh_all()
        self._status(f"{person.name}：{action}（{self.proj.month}月{day}日）")

    def _selected_sid_any(self) -> Optional[str]:
        """按当前页签取选中人员：名单页取 Treeview，不订页取左列表。"""
        tab = self._current_tab()
        if tab == 2:
            sel = self.roster_tree.selection()
            if sel:
                return sel[0]
        if tab == 3:
            sel = self.person_list.curselection()
            if sel:
                idx = sel[0]
                if 0 <= idx < len(self._person_order):
                    return self._person_order[idx]
        if self._current_sid and self._current_sid in self.proj.roster:
            return self._current_sid
        return None

    # ---------------------------------------------------------------- 数据刷新
    def refresh_all(self) -> None:
        self._sync_header_vars()
        self.refresh_today()
        self.refresh_days()
        self.refresh_roster()
        self.refresh_person_list()
        self.refresh_skip_grid()
        self.refresh_summary()
        self._update_title()

    def _sync_header_vars(self) -> None:
        self.title_var.set(self.proj.title)
        self.class_var.set(self.proj.class_name)
        self.year_var.set(str(self.proj.year))
        self.month_var.set(str(self.proj.month))
        self.price_var.set(fmt_money(self.proj.month_price))
        for idx, var in self._weekday_vars.items():
            var.set(idx in self.proj.skip_weekdays)
        self.skip_zero_var.set(bool(getattr(self.proj, "export_skip_zero", True)))

    def _update_title(self) -> None:
        mark = "*" if self.dirty else ""
        name = os.path.basename(self.path) if self.path else "未保存工程"
        self.title(f"{mark}{self.proj.title} - {name} - {APP_TITLE}")
        self.file_var.set(self.path or "未保存")

    def _touch(self) -> None:
        self.dirty = True
        self._update_title()

    def _status(self, message: str) -> None:
        self.status_var.set(message)

    # ---------------------------------------------------------------- 顶部应用
    def apply_header(self) -> None:
        try:
            year = int(str(self.year_var.get()).strip())
        except ValueError:
            messagebox.showwarning("提示", "年份必须是整数。", parent=self)
            return
        if not 1900 <= year <= 2100:
            messagebox.showwarning("提示", "年份需在 1900 ~ 2100 之间。", parent=self)
            return
        try:
            month = int(str(self.month_var.get()).strip())
        except ValueError:
            messagebox.showwarning("提示", "月份必须是 1 ~ 12 的整数。", parent=self)
            return
        if not 1 <= month <= 12:
            messagebox.showwarning("提示", "月份必须是 1 ~ 12 的整数。", parent=self)
            return
        raw_price = str(self.price_var.get()).strip().replace("，", ".")
        try:
            price = float(raw_price) if raw_price else 0.0
        except ValueError:
            messagebox.showwarning("提示", "全月单价必须是数字。", parent=self)
            return
        if price < 0:
            messagebox.showwarning("提示", "全月单价不能为负数。", parent=self)
            return

        title = self.title_var.get().strip() or "未命名工程"
        changed_month = (year != self.proj.year) or (month != self.proj.month)
        self.proj.title = title
        self.proj.class_name = self.class_var.get().strip()
        self.proj.year = year
        self.proj.month = month
        self.proj.month_price = round(price, 2)

        dropped = 0
        if changed_month:
            dim = self.proj.days_in_month
            for sid in list(self.proj.skips.keys()):
                kept = {d for d in self.proj.skips[sid] if 1 <= d <= dim}
                dropped += len(self.proj.skips[sid]) - len(kept)
                if kept:
                    self.proj.skips[sid] = kept
                else:
                    self.proj.skips.pop(sid, None)
            for key in list(self.proj.day_prices.keys()):
                try:
                    d = int(key)
                except ValueError:
                    self.proj.day_prices.pop(key, None)
                    continue
                if not 1 <= d <= dim:
                    self.proj.day_prices.pop(key, None)

        self.dirty = True
        self.refresh_all()
        self._status("设置已应用" + (f"，已清理 {dropped} 条超出本月天数的记录" if dropped else ""))

    def _on_weekday_toggle(self) -> None:
        """勾选/取消每周休息日：立即写入工程并刷新所有视图。"""
        self.proj.skip_weekdays = {idx for idx, var in self._weekday_vars.items() if var.get()}
        self.dirty = True
        self.refresh_days()
        self.refresh_roster()
        self.refresh_skip_grid()
        self.refresh_summary()
        self._touch()
        names = "、".join(WEEKDAYS_CN[i] for i in sorted(self.proj.skip_weekdays)) or "无"
        self._status(f"每周休息日已设为：{names}（这些天价格为 0）")

    # ---------------------------------------------------------------- 每日单价
    def refresh_days(self) -> None:
        tree = self.day_tree
        tree.delete(*tree.get_children())
        today = date.today()
        count_override = 0
        total = 0.0
        for day in self.proj.day_list():
            price = self.proj.price_for_day(day)
            total += price
            override = str(day) in self.proj.day_prices
            if override:
                count_override += 1
            if self.proj.is_closed_day(day):
                note = "每周休息日，价格 0"
            elif override:
                note = "单独设置"
            else:
                note = "按全月单价"
            tags = []
            wd = calendar.weekday(self.proj.year, self.proj.month, day)
            if self.proj.is_closed_day(day):
                tags.append("closed")
            elif wd >= 5:
                tags.append("weekend")
            if (today.year, today.month, today.day) == (self.proj.year, self.proj.month, day):
                tags.append("today")
                note += "（今天）"
            tree.insert("", "end", iid=f"d{day}",
                        values=(f"{self.proj.month}月{day}日", weekday_name(self.proj.year, self.proj.month, day),
                                fmt_money(price), note),
                        tags=tuple(tags))
        if self.proj.month_price <= 0:
            self.day_info_var.set(f"共 {self.proj.days_in_month} 天　尚未设置全月单价")
        else:
            self.day_info_var.set(
                f"共 {self.proj.days_in_month} 天　供餐 {len(self.proj.open_days())} 天　"
                f"全月单价 {fmt_money(self.proj.month_price)} 元/天　"
                f"按天合计 {fmt_money(total)} 元　单独设置 {count_override} 天"
            )

    def _selected_day(self, event=None) -> Optional[int]:
        if event is not None:
            iid = self.day_tree.identify_row(event.y)
            if iid:
                self.day_tree.selection_set(iid)
        sel = self.day_tree.selection()
        if not sel:
            return None
        try:
            return int(sel[0][1:])
        except (ValueError, IndexError):
            return None

    def edit_selected_day(self, event=None) -> None:
        day = self._selected_day(event)
        if day is None:
            messagebox.showinfo("提示", "请先在列表中选择一个日期。", parent=self)
            return
        if self.proj.is_closed_day(day):
            messagebox.showinfo(
                "提示",
                f"{self.proj.month}月{day}日 是{weekday_name(self.proj.year, self.proj.month, day)}，"
                "已被设为每周休息日，价格为 0。\n\n如需供餐，请在上方“每周休息”中取消勾选该星期。",
                parent=self,
            )
            return
        current = self.proj.price_for_day(day)
        value = simpledialog.askfloat(
            "设置单价",
            f"请输入 {self.proj.year}年{self.proj.month}月{day}日 的单价（元）：",
            initialvalue=current, minvalue=0.0, parent=self,
        )
        if value is None:
            return
        self.proj.day_prices[str(day)] = round(float(value), 2)
        self.dirty = True
        self.refresh_days()
        self.refresh_roster()
        self.refresh_summary()
        self._touch()
        self._status(f"已设置 {self.proj.month}月{day}日 单价为 {fmt_money(value)} 元")

    def clear_selected_day(self, event=None) -> None:
        day = self._selected_day(event)
        if day is None:
            messagebox.showinfo("提示", "请先在列表中选择一个日期。", parent=self)
            return
        if str(day) not in self.proj.day_prices:
            self._status(f"{self.proj.month}月{day}日 本来就没有单独设置")
            return
        self.proj.day_prices.pop(str(day), None)
        self.dirty = True
        self.refresh_days()
        self.refresh_roster()
        self.refresh_summary()
        self._touch()

    def reset_day_prices(self) -> None:
        if not self.proj.day_prices:
            messagebox.showinfo("提示", "当前没有单独设置的日期单价。", parent=self)
            return
        if not messagebox.askyesno("确认", f"将清除 {len(self.proj.day_prices)} 个日期的单独单价，全部改为全月单价，是否继续？",
                                   parent=self):
            return
        self.proj.day_prices.clear()
        self.dirty = True
        self.refresh_all()
        self._status("已清除全部日单价设置")

    # ---------------------------------------------------------------- 名单
    def refresh_roster(self) -> None:
        tree = self.roster_tree
        tree.delete(*tree.get_children())
        for sid in self.proj.sorted_sids():
            p = self.proj.roster[sid]
            skip_n = len(set(self.proj.skips.get(sid, set())) & set(self.proj.day_list()))
            values = (p.sid, p.name, p.gender, p.id_card, skip_n,
                      self.proj.ordered_days(sid), fmt_money(self.proj.person_cost(sid)))
            tree.insert("", "end", iid=sid, values=values,
                        tags=("optout",) if skip_n >= len(self.proj.open_days()) and skip_n else ())

    def refresh_person_list(self) -> None:
        self._person_order = self.proj.sorted_sids()
        self.person_list.delete(0, "end")
        for sid in self._person_order:
            p = self.proj.roster[sid]
            skip_n = len(set(self.proj.skips.get(sid, set())) & set(self.proj.day_list()))
            if skip_n and skip_n >= len(self.proj.open_days()):
                suffix = "（整月不订）"
            elif skip_n:
                suffix = f"（不订 {skip_n} 天）"
            else:
                suffix = ""
            self.person_list.insert("end", f"{sid}  {p.name}{suffix}")
        if self._current_sid in self._person_order:
            idx = self._person_order.index(self._current_sid)
            self.person_list.selection_set(idx)
            self.person_list.see(idx)
        else:
            self._current_sid = None

    def _selected_person_sid(self) -> Optional[str]:
        sel = self.roster_tree.selection()
        return sel[0] if sel else None

    def add_person(self) -> None:
        dlg = PersonDialog(self, "添加人员", existing_sids=self.proj.roster.keys())
        if dlg.result:
            self.proj.add_person(dlg.result)
            self._current_sid = dlg.result.sid
            self.dirty = True
            self.refresh_all()
            self._status(f"已添加 {dlg.result.name}（{dlg.result.sid}）")

    def edit_person(self) -> None:
        sid = self._selected_person_sid()
        if not sid:
            messagebox.showinfo("提示", "请先在名单中选择一名人员。", parent=self)
            return
        old = self.proj.roster[sid]
        dlg = PersonDialog(self, "编辑人员", person=old, existing_sids=self.proj.roster.keys())
        new = dlg.result
        if not new:
            return
        if new.sid != old.sid:
            skips = self.proj.skips.pop(old.sid, None)
            self.proj.roster.pop(old.sid, None)
            if skips:
                self.proj.skips[new.sid] = skips
        self.proj.roster[new.sid] = new
        self._current_sid = new.sid
        self.dirty = True
        self.refresh_all()
        self._status(f"已保存 {new.name}（{new.sid}）的信息")

    def delete_person(self) -> None:
        sid = self._selected_person_sid()
        if not sid:
            messagebox.showinfo("提示", "请先在名单中选择一名人员。", parent=self)
            return
        p = self.proj.roster[sid]
        if not messagebox.askyesno("确认删除", f"确定删除 {p.name}（{sid}）吗？\n其不订记录也会一并删除。", parent=self):
            return
        self.proj.remove_person(sid)
        if self._current_sid == sid:
            self._current_sid = None
        self.dirty = True
        self.refresh_all()
        self._status(f"已删除 {p.name}（{sid}）")

    def _merge_persons(self, persons: List[Person]) -> tuple[int, int]:
        added = updated = 0
        for p in persons:
            old = self.proj.roster.get(p.sid)
            if old:
                old.name = p.name
                old.id_card = p.id_card
                old.gender = p.gender
                updated += 1
            else:
                self.proj.add_person(p)
                added += 1
        return added, updated

    def _after_import(self, added: int, updated: int, warnings: List[str], source: str) -> None:
        self.dirty = True
        self.refresh_all()
        msg = f"已从{source}导入：新增 {added} 人，更新 {updated} 人。"
        if warnings:
            self._status(msg + f" 有 {len(warnings)} 条提示")
            self._show_warnings("导入提示", msg, warnings)
        else:
            self._status(msg)
            messagebox.showinfo("导入完成", msg, parent=self)

    def _show_warnings(self, title: str, summary: str, warnings: List[str]) -> None:
        win = tk.Toplevel(self)
        win.title(title)
        win.geometry("560x360")
        body = ttk.Frame(win, padding=12)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=summary).pack(anchor="w", pady=(0, 8))
        frame = ttk.Frame(body)
        frame.pack(fill="both", expand=True)
        text = tk.Text(frame, wrap="word", height=12)
        scroll = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        text.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        text.insert("1.0", "\n".join(warnings))
        text.configure(state="disabled")
        ttk.Button(body, text="知道了", command=win.destroy).pack(anchor="e", pady=(10, 0))
        win.transient(self)
        center_window(win, self)
        win.grab_set()

    def import_roster(self) -> None:
        path = filedialog.askopenfilename(
            parent=self, title="选择名单文件",
            filetypes=[("Excel / CSV 名单", "*.xlsx *.csv"), ("Excel 工作簿", "*.xlsx"), ("CSV 文件", "*.csv"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            persons, warnings = fee_core.import_roster_from_excel(path)
        except Exception as exc:  # noqa: BLE001 - 统一提示给用户
            messagebox.showerror("导入失败", str(exc), parent=self)
            return
        if not persons:
            messagebox.showwarning("导入失败", "没有导入到任何有效人员，请检查表格内容。", parent=self)
            return
        added, updated = self._merge_persons(persons)
        self._after_import(added, updated, warnings, os.path.basename(path))

    def import_roster_text(self) -> None:
        dlg = TextImportDialog(self)
        if not dlg.result:
            return
        added, updated = self._merge_persons(dlg.result)
        self._after_import(added, updated, dlg.warnings, "粘贴文本")

    def import_unordered(self) -> None:
        """从自由格式文本导入“本月不订”。"""
        if not self.proj.roster:
            messagebox.showinfo("提示", "请先导入或添加人员名单，再导入不订记录。", parent=self)
            return
        dlg = UnorderedImportDialog(self, self.proj)
        if not dlg.result:
            return
        self.dirty = True
        self.refresh_all()
        count = getattr(dlg, "applied_count", 0)
        msg = f"已把 {count} 人的不订记录写入工程。"
        if dlg.warnings:
            self._status(f"{msg} 有 {len(dlg.warnings)} 行未处理")
            self._show_warnings("导入本月不订", msg, dlg.warnings)
        else:
            self._status(msg)
            messagebox.showinfo("导入完成", msg, parent=self)

    # ---------------------------------------------------------------- 不订
    def refresh_skip_grid(self) -> None:
        for child in self.day_frame.winfo_children():
            child.destroy()
        self._day_vars.clear()
        self._day_cbs.clear()
        today = date.today()
        in_month = (today.year, today.month) == (self.proj.year, self.proj.month)
        for day in self.proj.day_list():
            var = tk.BooleanVar(value=False)
            self._day_vars[day] = var
            wd = weekday_name(self.proj.year, self.proj.month, day)
            closed = self.proj.is_closed_day(day)
            text = f"{day}日 {wd}"
            if closed:
                text += "·休息"
            elif in_month and today.day == day:
                text += "·今天"
            cb = ttk.Checkbutton(self.day_frame, text=text, variable=var,
                                 command=lambda d=day: self._on_day_toggle(d))
            cb.grid(row=(day - 1) // 7, column=(day - 1) % 7, sticky="w", padx=6, pady=5)
            if closed:
                # 每周休息日不供餐，无需也不能单独勾选“不订”
                cb.state(["disabled"])
            self._day_cbs[day] = cb
        self.skip_hint_var.set(f"今天是 {today.year}年{today.month}月{today.day}日" if in_month
                               else f"今天 {today.month}月{today.day}日 不在本工程月份内")
        self._load_skip_state()

    def _update_skip_month_info(self, sid: str) -> None:
        """更新“每日不订”页右上角的整月统计。"""
        total = len(self.proj.open_days())
        skipped = len(set(self.proj.skips.get(sid, set())) & set(self.proj.day_list()))
        if total and skipped >= total:
            self.skip_month_info_var.set(f"整月都不订（{skipped} 天）")
        elif skipped:
            self.skip_month_info_var.set(f"已不订 {skipped} 天 / 供餐 {total} 天")
        else:
            self.skip_month_info_var.set(f"每天订餐（供餐 {total} 天）")

    def _load_skip_state(self) -> None:
        self._loading_skip = True
        try:
            sid = self._current_sid
            if not sid or sid not in self.proj.roster:
                self.skip_title_var.set("请先在左侧选择人员")
                self.skip_month_info_var.set("")
                for var in self._day_vars.values():
                    var.set(False)
                for cb in self._day_cbs.values():
                    cb.state(["disabled"])
                return
            p = self.proj.roster[sid]
            self.skip_title_var.set(f"{p.name}（{sid}）")
            days = self.proj.skips.get(sid, set())
            for d, var in self._day_vars.items():
                var.set(d in days)
            for d, cb in self._day_cbs.items():
                cb.state(["disabled"] if self.proj.is_closed_day(d) else ["!disabled"])
            self._update_skip_month_info(sid)
        finally:
            self._loading_skip = False

    def _on_person_selected(self, event=None) -> None:
        sel = self.person_list.curselection()
        if not sel:
            return
        idx = sel[0]
        if 0 <= idx < len(self._person_order):
            self._current_sid = self._person_order[idx]
            self._load_skip_state()

    def _on_day_toggle(self, day: int) -> None:
        if self._loading_skip or not self._current_sid:
            return
        sid = self._current_sid
        if sid not in self.proj.roster:
            return
        days = {d for d, var in self._day_vars.items() if var.get() and not self.proj.is_closed_day(d)}
        if days:
            self.proj.set_skip(sid, days)
        else:
            self.proj.skips.pop(sid, None)
        self.dirty = True
        self.refresh_roster()
        self.refresh_person_list()
        self.refresh_summary()
        self.refresh_today()
        self._touch()
        self._update_skip_month_info(sid)
        self._status(f"{self.proj.roster[sid].name}：本月不订 {len(days)} 天")

    def _set_all_days(self, value: bool) -> None:
        if self._loading_skip or not self._current_sid:
            messagebox.showinfo("提示", "请先在左侧选择人员。", parent=self)
            return
        sid = self._current_sid
        if value and not messagebox.askyesno(
                "确认", f"将把 {self.proj.roster[sid].name} 当月每一天都标记为不订，是否继续？", parent=self):
            return
        for var in self._day_vars.values():
            var.set(value)
        # 休息日的复选框是禁用的，其变量也一并设值但不会写入 skips
        self._on_day_toggle(0)

    def mark_today_skip(self) -> None:
        if not self._current_sid:
            messagebox.showinfo("提示", "请先在左侧选择人员。", parent=self)
            return
        today = date.today()
        if (today.year, today.month) != (self.proj.year, self.proj.month):
            messagebox.showinfo(
                "提示",
                f"今天是 {today.year}年{today.month}月{today.day}日，不在当前工程（{self.proj.year}年{self.proj.month}月）内。",
                parent=self,
            )
            return
        var = self._day_vars.get(today.day)
        if var is None:
            return
        if self.proj.is_closed_day(today.day):
            messagebox.showinfo("提示", f"今天（{today.day}日）是每周休息日，本来就不供餐。", parent=self)
            return
        var.set(True)
        self._on_day_toggle(today.day)
        self._status(f"已为 {self.proj.roster[self._current_sid].name} 标记今天（{today.day}日）不订")

    # ---------------------------------------------------------------- 汇总
    def _on_skip_zero_toggle(self) -> None:
        """勾选“导出时不包括费用为 0 的人”，随工程保存。"""
        self.proj.export_skip_zero = bool(self.skip_zero_var.get())
        self.dirty = True
        self._touch()
        self._status("导出设置已更新：" + ("不包括费用为 0 的人" if self.proj.export_skip_zero
                                           else "包括所有人员"))

    def refresh_summary(self) -> None:
        tree = self.summary_tree
        tree.delete(*tree.get_children())
        rows = self.proj.settle()
        total_open = len(self.proj.open_days())
        for item in rows:
            n = len(set(self.proj.skips.get(item["sid"], set())) & set(self.proj.day_list()))
            if n and n >= total_open:
                note = "整月不订"
            else:
                note = f"不订 {n} 天" if n else ""
            tree.insert("", "end", iid=item["sid"],
                        values=(item["sid"], item["name"], item["gender"], item["id_card"],
                                item["days"], note, fmt_money(item["cost"])),
                        tags=("optout",) if note == "整月不订" else ())
        closed = len(self.proj.closed_days())
        zero = sum(1 for r in rows if float(r["cost"]) == 0.0)
        extra = f"　费用为 0 的 {zero} 人" if zero else ""
        self.total_var.set(
            f"合计：{fmt_money(self.proj.total_cost())} 元　共 {len(rows)} 人{extra}　"
            f"供餐 {len(self.proj.open_days())} 天" + (f"（每周休息 {closed} 天）" if closed else "")
        )

    def _on_tab_changed(self, event=None) -> None:
        try:
            idx = self.notebook.index(self.notebook.select())
        except tk.TclError:
            return
        if idx == 0:
            self.refresh_today()
        elif idx == 4:
            self.refresh_summary()

    # ---------------------------------------------------------------- 文件
    def _confirm_discard(self) -> bool:
        if not self.dirty:
            return True
        answer = messagebox.askyesnocancel("保存更改", "当前工程有未保存的更改，是否先保存？", parent=self)
        if answer is None:
            return False
        if answer:
            return self.save_project()
        return True

    def new_project(self) -> None:
        if not self._confirm_discard():
            return
        today = date.today()
        self.proj = Project(year=today.year, month=today.month)
        self.path = None
        self.dirty = False
        self._current_sid = None
        self.refresh_all()
        self._status("已新建工程")

    def open_project(self) -> None:
        if not self._confirm_discard():
            return
        path = filedialog.askopenfilename(
            parent=self, title="打开工程",
            filetypes=[("饭费工程文件", "*.cfee"), ("所有文件", "*.*")],
        )
        if not path:
            return
        self.load_from_path(path)

    def load_from_path(self, path: str, quiet: bool = False) -> bool:
        """打开指定路径的 .cfee 工程（供菜单与命令行共用）。"""
        path = os.path.abspath(os.path.expanduser(str(path).strip().strip('"')))
        if not os.path.exists(path):
            messagebox.showerror("打开失败", f"文件不存在：\n{path}", parent=self)
            return False
        try:
            proj = fee_core.load_project(path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("打开失败", f"无法打开该工程：\n{path}\n\n{exc}", parent=self)
            return False
        self.proj = proj
        self.path = path
        self.dirty = False
        self._current_sid = None
        self.refresh_all()
        self._status(f"已打开 {os.path.basename(path)}")
        if not quiet:
            self.lift()
            self.focus_force()
        return True

    def save_project(self) -> bool:
        if not self.path:
            return self.save_project_as()
        try:
            target = self.path
            if not target.lower().endswith(fee_core.FILE_SUFFIX):
                target += fee_core.FILE_SUFFIX
            fee_core.save_project(self.proj, target)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("保存失败", f"保存工程时出错：\n{exc}", parent=self)
            return False
        self.path = target
        self.dirty = False
        self._update_title()
        self._status(f"已保存到 {target}")
        return True

    def save_project_as(self) -> bool:
        path = filedialog.asksaveasfilename(
            parent=self, title="另存为",
            defaultextension=fee_core.FILE_SUFFIX,
            initialfile=safe_filename(self.proj.title) + fee_core.FILE_SUFFIX,
            filetypes=[("饭费工程文件", "*.cfee"), ("所有文件", "*.*")],
        )
        if not path:
            return False
        if not path.lower().endswith(fee_core.FILE_SUFFIX):
            path += fee_core.FILE_SUFFIX
        try:
            fee_core.save_project(self.proj, path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("保存失败", f"保存工程时出错：\n{exc}", parent=self)
            return False
        self.path = path
        self.dirty = False
        self._update_title()
        self._status(f"已保存到 {path}")
        return True

    def export_template(self) -> None:
        path = filedialog.asksaveasfilename(
            parent=self, title="保存名单模板",
            defaultextension=".xlsx", initialfile="名单模板.xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")],
        )
        if not path:
            return
        try:
            fee_core.export_roster_template(path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("导出失败", str(exc), parent=self)
            return
        self._status(f"名单模板已保存到 {path}")
        messagebox.showinfo("完成", f"名单模板已保存到：\n{path}", parent=self)

    def export_settlement(self) -> None:
        if not self.proj.roster:
            messagebox.showwarning("提示", "名单为空，无法导出核算表。", parent=self)
            return
        if self.proj.month_price <= 0 and not any(float(v) > 0 for v in self.proj.day_prices.values()):
            messagebox.showwarning("提示", "尚未设置任何单价（全月单价与每日单价均为 0），导出结果将全部为 0 元。",
                                   parent=self)
            return
        default = f"{safe_filename(self.proj.title)}_{self.proj.year}年{self.proj.month:02d}月核算.xlsx"
        path = filedialog.asksaveasfilename(
            parent=self, title="导出核算表",
            defaultextension=".xlsx", initialfile=default,
            filetypes=[("Excel 工作簿", "*.xlsx")],
        )
        if not path:
            return
        try:
            fee_core.export_settlement(self.proj, path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("导出失败", f"导出核算表时出错：\n{exc}", parent=self)
            return
        rows = self.proj.settle()
        exported = [r for r in rows if not self.skip_zero_var.get() or float(r["cost"]) != 0.0]
        skipped = len(rows) - len(exported)
        msg = f"核算表已导出到：\n{path}\n\n共 {len(exported)} 人"
        if skipped:
            msg += f"（已排除 {skipped} 位费用为 0 的人员）"
        msg += f"\n合计：{fmt_money(self.proj.total_cost())} 元"
        self._status(f"核算表已导出到 {path}")
        messagebox.showinfo("完成", msg, parent=self)

    # ---------------------------------------------------------------- 关于/帮助
    def show_help(self) -> None:
        messagebox.showinfo(
            "使用说明",
            "1. 顶部填写工程名称、班级、核算年月、全月单价，并勾选每周休息日，点击“应用设置”。\n"
            "   每周休息日默认为周六、周日：这些天价格为 0，不计入订餐天数。\n"
            "2. “每日单价”中可双击某天单独设置价格（如加餐日）。\n"
            "3. “人员名单”支持手动添加，或从 Excel/CSV、文本粘贴批量导入。\n"
            "4. “每日不订”中选择人员后勾选不订日期；“标记今天不订”可一键记录当天；\n"
            "   “导入本月不订...”可粘贴一整段自由文本，自动识别人员与“午餐/晚餐/周六餐”。\n"
            "5. “整月都不订”按钮会把当月每一天都标记为不订（等价于逐日勾选），费用为 0；\n"
            "   之后仍可单独取消其中某几天。\n"
            "6. “核算汇总”可查看每人应缴金额，并导出 Excel 核算表（班级/姓名/身份证号/费用）；\n"
            "   右上角“导出时不包括费用为 0 的人”默认勾选，并随工程保存。\n"
            "7. 工程保存为 .cfee 文件，可随时打开继续编辑；\n"
            "   也可用命令行直接打开：" + _program_name() + " 某工程.cfee\n"
            "8. 常用快捷键：Ctrl+S 保存、Ctrl+E 导出、Ctrl+1~5 切换页签、F5 刷新；\n"
            "   选中人员后按空格可标记 / 取消“今天不订”；按 Ctrl+/ 查看完整列表。",
            parent=self,
        )

    def show_shortcuts(self) -> None:
        """快捷键一览窗口。"""
        win = tk.Toplevel(self)
        win.title("快捷键一览")
        win.geometry("540x540")
        win.resizable(False, False)

        body = ttk.Frame(win, padding=14)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="快捷键", style="Big.TLabel").pack(anchor="w", pady=(0, 10))

        groups = [
            ("文件", [
                ("Ctrl+N", "新建工程"),
                ("Ctrl+O", "打开工程"),
                ("Ctrl+S", "保存"),
                ("Ctrl+Shift+S", "另存为"),
                ("Ctrl+I", "导入名单（Excel/CSV）"),
                ("Ctrl+Shift+I", "导入本月不订（自由文本）"),
                ("Ctrl+E", "导出核算表"),
                ("Alt+F4", "退出"),
            ]),
            ("编辑", [
                ("Ctrl+Shift+N", "添加人员"),
                ("F2", "编辑选中人员"),
                ("Delete", "删除选中人员"),
                ("空格", "标记 / 取消选中人员的“今天不订”"),
            ]),
            ("视图", [
                ("Ctrl+1 ~ Ctrl+5", "切换到第 1~5 个页签"),
                ("Ctrl+Tab / Ctrl+Shift+Tab", "下一个 / 上一个页签"),
                ("F5", "刷新全部数据"),
            ]),
            ("帮助", [
                ("F1", "使用说明"),
                ("Ctrl+/", "快捷键一览"),
            ]),
        ]

        for title, items in groups:
            frame = ttk.LabelFrame(body, text=title, padding=(10, 6))
            frame.pack(fill="x", pady=(0, 8))
            for key, desc in items:
                row = ttk.Frame(frame)
                row.pack(fill="x", pady=1)
                ttk.Label(row, text=key, width=26, font=("Consolas", 9, "bold")).pack(side="left")
                ttk.Label(row, text=desc).pack(side="left")

        ttk.Label(body, text="提示：光标在输入框内时，Delete / F2 / 空格不会触发，避免误操作；保存类快捷键始终有效。",
                  style="Hint.TLabel", wraplength=490, justify="left").pack(anchor="w", pady=(4, 8))
        ttk.Button(body, text="关闭", command=win.destroy, width=10).pack(anchor="e")

        win.transient(self)
        center_window(win, self)
        win.grab_set()

    def show_about(self) -> None:
        messagebox.showinfo(
            "关于",
            f"{APP_TITLE}\n\n用于按月统计学生/员工的订餐费用。\n"
            "工程文件扩展名：.cfee（JSON 格式）\n"
            "核算表导出：Excel .xlsx（班级 / 姓名 / 身份证号 / 费用）\n\n"
            "命令行用法：" + _program_name() + " [工程文件.cfee]",
            parent=self,
        )

    # ---------------------------------------------------------------- 关闭
    def _on_close(self) -> None:
        if self._confirm_discard():
            self.destroy()


def _program_name() -> str:
    """打包后显示 exe 名，源码运行时显示脚本名。"""
    if getattr(sys, "frozen", False):
        return os.path.basename(sys.executable)
    return "python fee_app.py"


def main(argv: Optional[List[str]] = None) -> int:
    """程序入口。argv[0] 之后的第一个参数视为要打开的 .cfee 工程路径。"""
    args = list(sys.argv[1:] if argv is None else argv)

    project_path: Optional[str] = None
    if args:
        first = args[0].strip().strip('"').strip("'")
        if first in ("-h", "--help", "/?"):
            print(f"用法: {_program_name()} [工程文件.cfee]")
            print("  不带参数启动空白工程；带路径时直接打开该 .cfee 工程。")
            return 0
        project_path = os.path.abspath(os.path.expanduser(first))
        if not os.path.exists(project_path):
            # 控制台提示 + 图形提示，避免双击关联文件时无反馈
            print(f"文件不存在：{project_path}", file=sys.stderr)
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("打开失败", f"文件不存在：\n{project_path}")
            root.destroy()
            return 2
        if not project_path.lower().endswith(fee_core.FILE_SUFFIX):
            candidate = project_path + fee_core.FILE_SUFFIX
            if os.path.exists(candidate):
                project_path = candidate

    app = FeeApp(project_path=project_path)
    app.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
