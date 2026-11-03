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
from typing import Dict, List, Optional, Tuple

# 嵌入式 Python 用 python311._pth 限制 sys.path，脚本目录不会自动加入；
# 显式加入，确保能 import 同目录的 fee_core。
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import fee_core
import fluent_ui as ui
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
        self.configure(bg=ui.C.bg)
        self.resizable(False, False)
        self.result: Optional[Person] = None
        self._existing = set(existing_sids or [])
        self._original_sid = person.sid if person else None

        self.sid_var = tk.StringVar(value=person.sid if person else "")
        self.name_var = tk.StringVar(value=person.name if person else "")
        self.id_var = tk.StringVar(value=person.id_card if person else "")
        self.gender_var = tk.StringVar(value=person.gender if person else "")

        card = ui.Card(self, padding=(24, 20))
        card.pack(fill="both", expand=True, padx=ui.S(14), pady=ui.S(14))
        body = card.body
        ui.text_label(body, title, font_key="h2").pack(anchor="w", pady=(0, ui.S(4)))
        ui.text_label(body, "带 * 的为必填项", font_key="caption", fg=ui.C.text3,
                      bg=ui.C.card).pack(anchor="w", pady=(0, ui.S(16)))

        def add_field(label: str, widget: tk.Widget, required: bool = False) -> None:
            row = ui.Surface(body, bg=ui.C.card)
            row.pack(fill="x", pady=(0, ui.S(14)))
            text = label + (" *" if required else "")
            ui.text_label(row, text, font_key="caption", fg=ui.C.text3).pack(
                anchor="w", pady=(0, ui.S(5)))
            widget.pack(fill="x")

        self.sid_entry = ui.FluentEntry(body, self.sid_var, width=34)
        add_field("学号", self.sid_entry, True)
        self.name_entry = ui.FluentEntry(body, self.name_var, width=34)
        add_field("姓名", self.name_entry, True)
        self.id_entry = ui.FluentEntry(body, self.id_var, width=34)
        add_field("身份证号", self.id_entry)
        self.gender_box = ui.FluentSelect(body, self.gender_var, ["", "男", "女"])
        add_field("性别", self.gender_box)

        ui.text_label(body, "身份证可留空；填写时会校验 18 位校验码，并可自动识别性别。",
                      font_key="caption", fg=ui.C.text3, bg=ui.C.card,
                      wrap=380).pack(anchor="w", pady=(0, ui.S(18)))

        btns = ui.Surface(body, bg=ui.C.card)
        btns.pack(fill="x")
        ui.FluentButton(btns, "确定", kind="accent", command=self._ok,
                        min_width=ui.S(88)).pack(side="right")
        ui.FluentButton(btns, "取消", command=self._cancel,
                        min_width=ui.S(76)).pack(side="right", padx=(0, ui.S(8)))

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
        self.configure(bg=ui.C.bg)
        self.geometry(f"{ui.S(820)}x{ui.S(580)}")
        self.resizable(True, True)
        self.result: Optional[List[Person]] = None
        self.warnings: List[str] = []

        card = ui.Card(self, padding=(22, 18))
        card.pack(fill="both", expand=True, padx=ui.S(14), pady=ui.S(14))
        body = card.body
        ui.text_label(body, "从文本粘贴导入名单", font_key="h2").pack(anchor="w")
        ui.text_label(body, "每行一人，列之间用空格 / 制表符 / 逗号分隔：学号 姓名 [身份证] [性别]",
                      font_key="caption", fg=ui.C.text3, bg=ui.C.card).pack(
                          anchor="w", pady=(ui.S(4), ui.S(10)))

        area = ui.TextArea(body, height=320)
        area.pack(fill="both", expand=True)
        self.text = area.text

        btns = ui.Surface(body, bg=ui.C.card)
        btns.pack(fill="x", pady=(ui.S(12), 0))
        ui.text_label(btns, "示例：20240001 张三 110101199003071234 男",
                      font_key="caption", fg=ui.C.text3).pack(side="left")
        ui.FluentButton(btns, "导入", kind="accent", command=self._ok,
                        min_width=ui.S(88)).pack(side="right")
        ui.FluentButton(btns, "取消", command=self.destroy,
                        min_width=ui.S(76)).pack(side="right", padx=(0, ui.S(8)))

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
        self.configure(bg=ui.C.bg)
        self.geometry(f"{ui.S(1000)}x{ui.S(700)}")
        self.minsize(ui.S(840), ui.S(580))
        self.proj = proj
        self.result: Optional[List[Dict]] = None
        self.warnings: List[str] = []

        card = ui.Card(self, padding=(22, 16))
        card.pack(fill="both", expand=True, padx=ui.S(14), pady=ui.S(14))
        body = card.body

        head = ui.Surface(body, bg=ui.C.card)
        head.pack(fill="x")
        ui.text_label(head, "导入本月不订（自由文本）", font_key="h2").pack(anchor="w")
        ui.text_label(
            head,
            "每行一人，格式不限。示例：44号范付璐10月不订午餐、晚餐、周六餐；"
            "陈心怡不订所有餐；赵思琪只订午餐",
            font_key="caption", fg=ui.C.text3, bg=ui.C.card, wrap=880,
        ).pack(anchor="w", pady=(ui.S(4), ui.S(12)))

        # 数据中只保留这些餐别的不订记录；未勾选的餐别一律视为照常订餐
        keep_box = ui.Card(body, padding=(14, 10), fill=ui.C.card_alt, border=ui.C.border)
        keep_box.pack(fill="x", pady=(0, ui.S(10)))
        ui.text_label(keep_box.body, "只记录这些餐别的不订（未勾选的餐别一律视为照常订餐，文本里提到也不记录）",
                      font_key="caption", fg=ui.C.text3, bg=ui.C.card_alt).pack(anchor="w", pady=(0, ui.S(6)))
        meal_row = ui.Surface(keep_box.body, bg=ui.C.card_alt)
        meal_row.pack(fill="x")
        self.meal_vars: Dict[str, tk.BooleanVar] = {}
        for meal in fee_core._ALL_MEALS:
            var = tk.BooleanVar(value=True)
            self.meal_vars[meal] = var
            ui.FluentCheck(meal_row, meal, var, kind="chip",
                           command=self._on_keep_changed).pack(side="left", padx=(0, ui.S(8)))
        self.keep_hint_var = tk.StringVar()
        ui.text_label(meal_row, textvariable=self.keep_hint_var, font_key="caption",
                      fg=ui.C.text3, bg=ui.C.card_alt).pack(side="left", padx=(ui.S(8), 0))

        self.replace_var = tk.BooleanVar(value=False)
        ui.FluentCheck(body, "覆盖这些人员的原有逐日不订记录（默认与原有记录合并）",
                       self.replace_var).pack(anchor="w", pady=(0, ui.S(10)))

        cols = ui.Surface(body, bg=ui.C.card)
        cols.pack(fill="both", expand=True)
        cols.columnconfigure(0, weight=3, uniform="col")
        cols.columnconfigure(1, weight=2, uniform="col")
        cols.rowconfigure(0, weight=1)

        left = ui.Surface(cols, bg=ui.C.card)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, ui.S(12)))
        area = ui.TextArea(left, height=240)
        area.pack(fill="both", expand=True)
        self.text = area.text
        self.text.insert("1.0", SAMPLE_UNORDERED)
        left_btns = ui.Surface(left, bg=ui.C.card)
        left_btns.pack(fill="x", pady=(ui.S(10), 0))
        ui.FluentButton(left_btns, "解析预览", kind="accent", command=self._preview,
                        min_width=ui.S(96)).pack(side="left")
        ui.FluentButton(left_btns, "从文件读取...", command=self._load_file).pack(
            side="left", padx=(ui.S(8), 0))

        self.preview_card = ui.TableCard(
            cols,
            {"name": ("姓名", 90, "w"), "meals": ("不订餐别", 150, "w"),
             "days": ("不订天数", 80, "center")},
            title="解析结果", empty_text="输入文本后点击“解析预览”",
            border=ui.C.border_strong)
        self.preview_card.grid(row=0, column=1, sticky="nsew")
        self.preview_tree = self.preview_card.tree

        self.info_var = tk.StringVar(value="输入文本后点击“解析预览”。")
        ui.text_label(body, textvariable=self.info_var, font_key="caption",
                      fg=ui.C.text3, bg=ui.C.card).pack(anchor="w", pady=(ui.S(10), 0))

        btns = ui.Surface(body, bg=ui.C.card)
        btns.pack(fill="x", pady=(ui.S(10), 0))
        ui.FluentButton(btns, "应用到工程", kind="accent", command=self._apply,
                        min_width=ui.S(104)).pack(side="right")
        ui.FluentButton(btns, "取消", command=self.destroy,
                        min_width=ui.S(76)).pack(side="right", padx=(0, ui.S(8)))

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
        self.preview_card.set_empty(not results, "没有识别到人员")
        self.preview_card.set_subtitle(f"{len(results)} 人")
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


class _PersonListAdapter:
    """把单列 Treeview 包成 Listbox 的常用接口，让“每日不订”页风格统一。"""

    def __init__(self, tree: ttk.Treeview, empty_label: tk.Label) -> None:
        self._tree = tree
        self._empty = empty_label
        self._empty_visible = False

    def delete(self, first: Any, last: Any = None) -> None:
        self._tree.delete(*self._tree.get_children())

    def insert(self, index: Any, text: str) -> None:
        self._tree.insert("", "end", values=(text,))

    def selection_set(self, index: int) -> None:
        children = self._tree.get_children()
        if 0 <= index < len(children):
            self._tree.selection_set(children[index])

    def curselection(self) -> Tuple[int, ...]:
        selected = self._tree.selection()
        if not selected:
            return ()
        children = self._tree.get_children()
        try:
            return (children.index(selected[0]),)
        except ValueError:
            return ()

    def see(self, index: int) -> None:
        children = self._tree.get_children()
        if 0 <= index < len(children):
            self._tree.see(children[index])

    def set_empty(self, visible: bool) -> None:
        if visible and not self._empty_visible:
            self._empty.place(relx=0.5, rely=0.5, anchor="center")
            self._empty_visible = True
        elif not visible and self._empty_visible:
            self._empty.place_forget()
            self._empty_visible = False


# ---------------------------------------------------------------------------
# 主窗口
# ---------------------------------------------------------------------------

class FeeApp(tk.Tk):
    def __init__(self, project_path: Optional[str] = None):
        super().__init__()
        self.proj = Project(year=date.today().year, month=date.today().month)
        self.path: Optional[str] = None
        self.dirty = False
        self._settings_open = True

        self._day_vars: Dict[int, tk.BooleanVar] = {}
        self._day_cbs: Dict[int, ui.FluentCheck] = {}
        self._person_order: List[str] = []
        self._current_sid: Optional[str] = None
        self._loading_skip = False
        self._weekday_vars: Dict[int, tk.BooleanVar] = {}

        self.title(APP_TITLE)
        self.configure(bg=ui.C.bg)
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        width = min(ui.S(1200), max(screen_w - ui.S(80), ui.S(720)))
        height = min(ui.S(800), max(screen_h - ui.S(130), ui.S(520)))
        self.geometry(f"{width}x{height}")
        self.minsize(min(ui.S(1040), width), min(ui.S(700), height))
        self._build_style()
        self._apply_app_icon()
        self._build_menu()
        self._build_header()
        self._build_status()      # 先占住底部，避免被内容挤掉
        self._build_notebook()
        self._bind_shortcuts()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.refresh_all()
        center_window(self)
        self.after(120, lambda: ui.style_titlebar(self))

        # 命令行传入工程文件时直接打开（不弹“是否保存”确认）
        if project_path:
            self.after(80, lambda: self.load_from_path(project_path, quiet=False))

    def _apply_app_icon(self) -> None:
        icon = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "app.ico")
        try:
            if os.path.exists(icon):
                self.iconbitmap(icon)
        except tk.TclError:
            pass


    # ---------------------------------------------------------------- 样式
    def _build_style(self) -> None:
        ui.init(self)

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
        bar = ui.Surface(self)
        bar.pack(fill="x", padx=ui.S(18), pady=(ui.S(14), ui.S(2)))

        ui.BrandTile(bar, "¥", size=40).pack(side="left", padx=(0, ui.S(12)))

        brand = ui.Surface(bar)
        brand.pack(side="left", fill="y")
        ui.text_label(brand, "饭费管理系统", font_key="h1").pack(anchor="w")
        self.brand_sub_var = tk.StringVar()
        ui.text_label(brand, textvariable=self.brand_sub_var, font_key="caption",
                      fg=ui.C.text3).pack(anchor="w", pady=(ui.S(2), 0))

        actions = ui.Surface(bar)
        actions.pack(side="right")
        self.save_btn = ui.FluentButton(actions, "保存工程", kind="accent",
                                        command=self.save_project, min_width=ui.S(96))
        self.save_btn.pack(side="right")
        self.settings_btn = ui.FluentButton(actions, "收起设置", kind="standard",
                                            command=self._toggle_settings, min_width=ui.S(88))
        self.settings_btn.pack(side="right", padx=(0, ui.S(8)))

        # 工程设置卡片（可折叠）：直接挂在窗口上，展开时插到页签容器之前
        self.tab_host = ui.Surface(self)
        self.tab_host.pack(fill="both", expand=True)
        self.settings_card = self._build_settings_card(self)
        self.settings_card.pack(fill="x", padx=ui.S(18), pady=(ui.S(8), 0),
                                before=self.tab_host)

    def _build_settings_card(self, master: tk.Misc) -> ui.Card:
        card = ui.Card(master, padding=(18, 14), title="工程设置")
        self.title_var = tk.StringVar()
        self.class_var = tk.StringVar()
        self.year_var = tk.StringVar()
        self.month_var = tk.StringVar()
        self.price_var = tk.StringVar()

        ui.FluentButton(card.actions, "应用设置", kind="accent",
                        command=self.apply_header, min_width=ui.S(88)).pack(side="right")

        row = ui.Surface(card.body, bg=ui.C.card)
        row.pack(fill="x")

        def field(label: str, factory, gap: int = 18) -> None:
            cell = ui.Surface(row, bg=ui.C.card)
            cell.pack(side="left", padx=(0, ui.S(gap)))
            ui.text_label(cell, label, font_key="caption", fg=ui.C.text3).pack(
                anchor="w", pady=(0, ui.S(4)))
            factory(cell).pack(anchor="w")

        field("工程名称", lambda m: ui.FluentEntry(m, self.title_var, width=20))
        field("班级", lambda m: ui.FluentEntry(m, self.class_var, width=14))

        # 核算月份：年 / 月 两个下拉
        year_cell = ui.Surface(row, bg=ui.C.card)
        year_cell.pack(side="left", padx=(0, ui.S(18)))
        ui.text_label(year_cell, "核算月份", font_key="caption",
                      fg=ui.C.text3).pack(anchor="w", pady=(0, ui.S(4)))
        month_row = ui.Surface(year_cell, bg=ui.C.card)
        month_row.pack(anchor="w")
        years = [str(y) for y in range(date.today().year - 10, date.today().year + 11)]
        self.year_select = ui.FluentSelect(month_row, self.year_var, years, width=ui.S(96))
        self.year_select.pack(side="left")
        ui.text_label(month_row, "年", font_key="body", fg=ui.C.text2).pack(
            side="left", padx=(ui.S(6), ui.S(10)))
        self.month_select = ui.FluentSelect(month_row, self.month_var,
                                            [str(m) for m in range(1, 13)], width=ui.S(76))
        self.month_select.pack(side="left")
        ui.text_label(month_row, "月", font_key="body", fg=ui.C.text2).pack(
            side="left", padx=(ui.S(6), 0))

        price_cell = ui.Surface(row, bg=ui.C.card)
        price_cell.pack(side="left")
        ui.text_label(price_cell, "全月单价", font_key="caption",
                      fg=ui.C.text3).pack(anchor="w", pady=(0, ui.S(4)))
        price_row = ui.Surface(price_cell, bg=ui.C.card)
        price_row.pack(anchor="w")
        ui.FluentEntry(price_row, self.price_var, width=8).pack(side="left")
        ui.text_label(price_row, "元 / 天（0 表示未设置）", font_key="body",
                      fg=ui.C.text2).pack(side="left", padx=(ui.S(6), 0))

        rest = ui.Surface(card.body, bg=ui.C.card)
        rest.pack(fill="x", pady=(ui.S(12), 0))
        ui.text_label(rest, "每周休息日", font_key="caption", fg=ui.C.text3).pack(
            side="left", padx=(0, ui.S(12)))
        for idx, name in enumerate(WEEKDAYS_CN):
            var = tk.BooleanVar(value=(idx in self.proj.skip_weekdays))
            self._weekday_vars[idx] = var
            ui.FluentCheck(rest, name, var, kind="chip",
                           command=self._on_weekday_toggle).pack(side="left",
                                                                padx=(0, ui.S(6)))
        ui.text_label(rest, "勾选表示该星期不供餐，价格按 0 计。", font_key="caption",
                      fg=ui.C.text3).pack(side="left", padx=(ui.S(8), 0))
        return card

    def _toggle_settings(self) -> None:
        self._settings_open = not self._settings_open
        if self._settings_open:
            self.settings_card.pack(fill="x", padx=ui.S(18), pady=(ui.S(8), 0),
                                    before=self.tab_host)
            self.settings_btn.set_text("收起设置")
            self._status("已展开工程设置")
        else:
            self.settings_card.pack_forget()
            self.settings_btn.set_text("工程设置")
            self._status("已收起工程设置")
        # pack 顺序变化后立即重算几何，避免这一帧内容被挤压/裁掉
        self.update_idletasks()
        self.update()

    # ---------------------------------------------------------------- 选项卡
    def _build_notebook(self) -> None:
        self.notebook = ui.TabView(self.tab_host)
        self.notebook.pack(fill="both", expand=True, padx=ui.S(18), pady=(ui.S(10), ui.S(6)))

        self.tab_today = ui.Surface(self.notebook.host)
        self.tab_days = ui.Surface(self.notebook.host)
        self.tab_roster = ui.Surface(self.notebook.host)
        self.tab_skip = ui.Surface(self.notebook.host)
        self.tab_summary = ui.Surface(self.notebook.host)

        for page, name in ((self.tab_today, "今日订餐"), (self.tab_days, "每日单价"),
                           (self.tab_roster, "人员名单"), (self.tab_skip, "每日不订"),
                           (self.tab_summary, "核算汇总")):
            self.notebook.add(page, name)

        self._build_tab_today()
        self._build_tab_days()
        self._build_tab_roster()
        self._build_tab_skip()
        self._build_tab_summary()
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

    # -- 今日订餐 ------------------------------------------------------
    def _build_tab_today(self) -> None:
        head = ui.Surface(self.tab_today)
        head.pack(fill="x", pady=(0, ui.S(10)))
        ui.text_label(head, "今日订餐", font_key="h2").pack(side="left")
        self.today_date_var = tk.StringVar()
        ui.text_label(head, textvariable=self.today_date_var, font_key="caption",
                      fg=ui.C.text3).pack(side="left", padx=(ui.S(10), 0))
        ui.FluentButton(head, "刷新", kind="subtle", command=self.refresh_today,
                        min_width=ui.S(64)).pack(side="right")

        stats = ui.Surface(self.tab_today)
        stats.pack(fill="x", pady=(0, ui.S(12)))
        for i in range(3):
            stats.columnconfigure(i, weight=1, uniform="today-stat")
        self.stat_today = ui.StatCard(stats, "今日订餐", unit="人", value="—", accent=True)
        self.stat_today.grid(row=0, column=0, sticky="nsew", padx=(0, ui.S(12)))
        self.stat_price = ui.StatCard(stats, "今日单价", unit="元/天", value="—")
        self.stat_price.grid(row=0, column=1, sticky="nsew", padx=(0, ui.S(12)))
        self.stat_income = ui.StatCard(stats, "预计收入", unit="元", value="—")
        self.stat_income.grid(row=0, column=2, sticky="nsew")

        panes = ui.Surface(self.tab_today)
        panes.pack(fill="both", expand=True)
        panes.columnconfigure(0, weight=1, uniform="today-pane")
        panes.columnconfigure(1, weight=1, uniform="today-pane")
        panes.rowconfigure(0, weight=1)

        self.diner_card = ui.TableCard(
            panes,
            {"sid": ("学号", 110, "w"), "name": ("姓名", 100, "w"),
             "gender": ("性别", 64, "center")},
            title="今日订餐", empty_text="今天还没有人订餐")
        self.diner_card.grid(row=0, column=0, sticky="nsew", padx=(0, ui.S(12)))
        self.absent_card = ui.TableCard(
            panes,
            {"sid": ("学号", 90, "w"), "name": ("姓名", 90, "w"),
             "reason": ("原因", 120, "w")},
            title="今日不订", empty_text="今天所有人都订餐")
        self.absent_card.grid(row=0, column=1, sticky="nsew")

        self.today_diner_tree = self.diner_card.tree
        self.today_absent_tree = self.absent_card.tree
        self.today_absent_tree.tag_configure("closed", foreground=ui.C.text3)

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
            self.stat_today.set("—", f"本工程为 {self.proj.year}年{self.proj.month}月")
            self.stat_price.set("—", "请切换到“每日单价”页查看")
            self.stat_income.set("—", "—")
            self.diner_card.set_subtitle("—")
            self.absent_card.set_subtitle("—")
            self.diner_card.set_empty(True, "不在本工程月份")
            self.absent_card.set_empty(True, "不在本工程月份")
            return

        day = today.day
        self.today_date_var.set(
            f"今天 {today.year}年{today.month}月{today.day}日　{weekday_name(today.year, today.month, day)}"
        )

        if self.proj.is_closed_day(day):
            self.stat_today.set("0", f"今天是每周休息日，不供餐 · 名单共 {total} 人")
            self.stat_price.set(fmt_money(self.proj.price_for_day(day)), "每周休息日单价为 0")
            self.stat_income.set("¥0.00", "今天不供餐")
            for sid in self.proj.sorted_sids():
                p = self.proj.roster[sid]
                absent.insert("", "end", values=(p.sid, p.name, "每周休息日"), tags=("closed",))
            self.diner_card.set_subtitle("0 人")
            self.absent_card.set_subtitle(f"{total} 人")
            self.diner_card.set_empty(True, "今天是每周休息日")
            self.absent_card.set_empty(total == 0)
            return

        diners = self.proj.diners_on(day)
        total_open = len(self.proj.open_days())
        price = self.proj.price_for_day(day)
        absent_count = total - len(diners)
        self.stat_today.set(str(len(diners)), f"共 {total} 人 · 不订 {absent_count} 人")
        self.stat_price.set(
            fmt_money(price),
            "按全月单价" if price == self.proj.month_price else "单独设置的单价")
        self.stat_income.set("¥" + fmt_money(price * len(diners)),
                             f"订餐 {len(diners)} 人 × {fmt_money(price)} 元")
        self.diner_card.set_subtitle(f"{len(diners)} 人")
        self.absent_card.set_subtitle(f"{absent_count} 人")
        self.diner_card.set_empty(not diners, "今天还没有人订餐")
        self.absent_card.set_empty(absent_count == 0, "今天所有人都订餐")

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
        strip = ui.StatStrip(self.tab_days,
                             ["供餐天数", "全月单价", "按天合计", "单独设置"])
        strip.pack(fill="x", pady=(0, ui.S(12)))
        self.day_info_strip = strip

        bar = ui.Surface(self.tab_days)
        bar.pack(fill="x", pady=(0, ui.S(10)))
        ui.FluentButton(bar, "修改选中日期单价...", kind="accent",
                        command=self.edit_selected_day).pack(side="left")
        ui.FluentButton(bar, "清除选中日期设置", kind="subtle",
                        command=self.clear_selected_day).pack(side="left", padx=(ui.S(8), 0))
        ui.FluentButton(bar, "全部恢复全月单价", kind="subtle",
                        command=self.reset_day_prices).pack(side="left", padx=(ui.S(8), 0))
        self.day_info_var = tk.StringVar()
        ui.text_label(bar, textvariable=self.day_info_var, font_key="caption",
                      fg=ui.C.text3).pack(side="right")

        self.day_card = ui.TableCard(
            self.tab_days,
            {"day": ("日期", 140, "w"), "week": ("星期", 100, "center"),
             "price": ("单价（元）", 140, "center"), "note": ("设置说明", 320, "w")},
            title="每日单价", empty_text="本月还没有日期")
        self.day_card.pack(fill="both", expand=True)
        self.day_tree = self.day_card.tree
        self.day_tree.tag_configure("weekend", background="#FAFAFA")
        self.day_tree.tag_configure("closed", background="#F0F0F0", foreground=ui.C.text3)
        self.day_tree.tag_configure("today", background="#FFF8E1")
        self.day_tree.tag_configure("stripe", background="#FAFAFA")
        self.day_tree.bind("<Double-1>", lambda e: self.edit_selected_day(event=e))

    # -- 人员名单 ------------------------------------------------------
    def _build_tab_roster(self) -> None:
        bar = ui.Surface(self.tab_roster)
        bar.pack(fill="x", pady=(0, ui.S(10)))
        ui.FluentButton(bar, "添加人员", kind="accent", command=self.add_person).pack(
            side="left", padx=(0, ui.S(8)))
        ui.FluentButton(bar, "编辑", command=self.edit_person).pack(side="left", padx=(0, ui.S(6)))
        ui.FluentButton(bar, "删除", kind="danger", command=self.delete_person).pack(
            side="left", padx=(0, ui.S(16)))
        ui.FluentButton(bar, "导入名单（Excel/CSV）...", kind="subtle",
                        command=self.import_roster).pack(side="left", padx=(0, ui.S(6)))
        ui.FluentButton(bar, "文本粘贴导入...", kind="subtle",
                        command=self.import_roster_text).pack(side="left", padx=(0, ui.S(6)))
        ui.FluentButton(bar, "下载模板", kind="subtle",
                        command=self.export_template).pack(side="left")

        self.roster_card = ui.TableCard(
            self.tab_roster,
            {"sid": ("学号", 130, "w"), "name": ("姓名", 110, "w"),
             "gender": ("性别", 64, "center"), "id_card": ("身份证号", 210, "w"),
             "skip_days": ("不订天数", 96, "center"), "days": ("订餐天数", 96, "center"),
             "cost": ("应缴金额（元）", 130, "center")},
            title="人员名单", empty_text="还没有人员，点“添加人员”或导入名单")
        self.roster_card.pack(fill="both", expand=True)
        self.roster_tree = self.roster_card.tree
        self.roster_tree.tag_configure("optout", foreground=ui.C.text3)
        self.roster_tree.tag_configure("stripe", background="#FAFAFA")
        self.roster_tree.bind("<Double-1>", lambda e: self.edit_person())

    # -- 每日不订 ------------------------------------------------------
    def _build_tab_skip(self) -> None:
        body = ui.Surface(self.tab_skip)
        body.pack(fill="both", expand=True)

        left = ui.Card(body, padding=(0, 14), title="人员名单")
        left.pack(side="left", fill="y", padx=(0, ui.S(12)))
        left.body.pack_configure(padx=0)
        if left.header is not None:
            left.header.pack_configure(padx=ui.S(16))

        self.person_list_card = left
        ui.divider(left.body).pack(fill="x")
        holder = ui.Surface(left.body, bg=ui.C.card)
        holder.pack(fill="both", expand=True)
        self.person_tree = ttk.Treeview(holder, columns=("name",), show="headings",
                                        selectmode="browse")
        self.person_tree.heading("name", text="人员", anchor="w")
        self.person_tree.column("name", width=ui.S(232), anchor="w", stretch=True)
        scroll = ui.FluentScrollbar(holder, command=self.person_tree.yview)
        self.person_tree.configure(yscrollcommand=scroll.set)
        self.person_tree.pack(side="left", fill="both", expand=True, padx=(ui.S(6), 0))
        scroll.pack(side="right", fill="y", padx=(0, ui.S(4)))
        self.person_tree.bind("<<TreeviewSelect>>", self._on_person_selected)
        self.person_empty = tk.Label(holder, text="名单为空", bg=ui.C.card,
                                     fg=ui.C.text3, font=ui.F("caption"))
        self.person_list = _PersonListAdapter(self.person_tree, self.person_empty)

        right = ui.Surface(body)
        right.pack(side="left", fill="both", expand=True)

        head = ui.Surface(right)
        head.pack(fill="x")
        info = ui.Surface(head)
        info.pack(side="left")
        ui.text_label(info, "每日不订", font_key="caption", fg=ui.C.text3).pack(anchor="w")
        self.skip_title_var = tk.StringVar(value="请先在左侧选择人员")
        ui.text_label(info, textvariable=self.skip_title_var, font_key="h2").pack(anchor="w")
        self.skip_month_info_var = tk.StringVar()
        self.skip_badge = ui.FluentCheck(head, "—", kind="chip", min_width=ui.S(150),
                                         enabled=False)
        self.skip_badge.pack(side="right", pady=(ui.S(10), 0))

        quick = ui.Surface(right)
        quick.pack(fill="x", pady=(ui.S(10), ui.S(10)))
        self.skip_buttons = [
            ui.FluentButton(quick, "标记今天不订", kind="accent",
                            command=self.mark_today_skip, enabled=False),
            ui.FluentButton(quick, "整月都不订",
                            command=lambda: self._set_all_days(True), enabled=False),
            ui.FluentButton(quick, "全部清除", kind="subtle",
                            command=lambda: self._set_all_days(False), enabled=False),
        ]
        self.skip_buttons[0].pack(side="left", padx=(0, ui.S(8)))
        self.skip_buttons[1].pack(side="left", padx=(0, ui.S(8)))
        self.skip_buttons[2].pack(side="left")
        ui.FluentButton(quick, "导入本月不订...", kind="subtle",
                        command=self.import_unordered).pack(side="left", padx=(ui.S(8), 0))
        self.skip_hint_var = tk.StringVar()
        ui.text_label(quick, textvariable=self.skip_hint_var, font_key="caption",
                      fg=ui.C.text3).pack(side="right")

        day_card = ui.Card(right, padding=(14, 12), title="不订日期",
                           subtitle="勾选表示当天不订餐")
        day_card.pack(fill="both", expand=True)
        self.day_frame = ui.Surface(day_card.body, bg=ui.C.card)
        self.day_frame.pack(fill="both", expand=True)

        ui.text_label(right, "提示：“整月都不订”就是把当月每一天都标记为不订；逐日勾选会实时保存到工程。",
                      font_key="caption", fg=ui.C.text3).pack(anchor="w", pady=(ui.S(8), 0))

    # -- 核算汇总 ------------------------------------------------------
    def _build_tab_summary(self) -> None:
        strip = ui.StatStrip(self.tab_summary, ["合计金额", "人数", "供餐天数", "每周休息"])
        strip.pack(fill="x", pady=(0, ui.S(12)))
        self.summary_strip = strip

        bar = ui.Surface(self.tab_summary)
        bar.pack(fill="x", pady=(0, ui.S(10)))
        ui.FluentButton(bar, "导出核算表（Excel）...", kind="accent",
                        command=self.export_settlement).pack(side="left")
        ui.FluentButton(bar, "刷新", kind="subtle",
                        command=self.refresh_summary).pack(side="left", padx=(ui.S(8), ui.S(16)))
        self.skip_zero_var = tk.BooleanVar(value=True)
        ui.FluentCheck(bar, "导出时不包括费用为 0 的人", self.skip_zero_var,
                       command=self._on_skip_zero_toggle).pack(side="left", pady=(ui.S(4), 0))
        self.total_var = tk.StringVar(value="合计：0.00 元")

        self.summary_card = ui.TableCard(
            self.tab_summary,
            {"sid": ("学号", 120, "w"), "name": ("姓名", 100, "w"),
             "gender": ("性别", 64, "center"), "id_card": ("身份证号", 200, "w"),
             "days": ("订餐天数", 96, "center"), "note": ("备注", 140, "w"),
             "cost": ("应缴金额（元）", 130, "center")},
            title="核算汇总", empty_text="还没有可汇总的数据")
        self.summary_card.pack(fill="both", expand=True)
        self.summary_tree = self.summary_card.tree
        self.summary_tree.tag_configure("optout", foreground=ui.C.text3)
        self.summary_tree.tag_configure("stripe", background="#FAFAFA")

    # ---------------------------------------------------------------- 状态栏
    def _build_status(self) -> None:
        ui.divider(self).pack(side="bottom", fill="x")
        bar = ui.Surface(self)
        bar.pack(side="bottom", fill="x", padx=ui.S(18), pady=(ui.S(6), ui.S(8)))
        self.status_dot = ui.StatusDot(bar, ui.C.success)
        self.status_dot.pack(side="left")
        self.status_var = tk.StringVar(value="就绪")
        ui.text_label(bar, textvariable=self.status_var, font_key="caption",
                      fg=ui.C.text2).pack(side="left", padx=(ui.S(8), 0))
        self.file_var = tk.StringVar(value="未保存")
        ui.text_label(bar, textvariable=self.file_var, font_key="caption",
                      fg=ui.C.text3).pack(side="right")

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
        # 打开旧工程时，年份可能不在默认下拉范围里，补进去
        years = {str(y) for y in range(date.today().year - 10, date.today().year + 11)}
        years.add(str(self.proj.year))
        self.year_select.set_values(sorted(years))
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
        self.file_var.set(("● " if self.dirty else "") + (self.path or "未保存"))
        if getattr(self, "save_btn", None) is not None:
            self.save_btn.set_enabled(self.dirty or not self.path)
        if getattr(self, "status_dot", None) is not None:
            self.status_dot.set_color(ui.C.warning if self.dirty else ui.C.success)
        if getattr(self, "brand_sub_var", None) is not None:
            closed = len(self.proj.closed_days())
            rest = "、".join(WEEKDAYS_CN[i] for i in sorted(self.proj.skip_weekdays)) or "无"
            parts = [
                f"{self.proj.year}年{self.proj.month}月",
                self.proj.class_name or "未填写班级",
                f"全月单价 {fmt_money(self.proj.month_price)} 元/天",
                f"名单 {len(self.proj.roster)} 人",
            ]
            if self.proj.month_price > 0:
                parts.append(f"供餐 {len(self.proj.open_days())} 天")
            parts.append(f"休息日 {rest}")
            self.brand_sub_var.set("　·　".join(parts))
            if getattr(self, "settings_card", None) is not None:
                self.settings_card.set_subtitle(
                    f"{self.proj.year}年{self.proj.month}月 · "
                    f"{self.proj.class_name or '未填写班级'} · "
                    f"{fmt_money(self.proj.month_price)} 元/天")

    def _touch(self) -> None:
        self.dirty = True
        self._update_title()

    def _status(self, message: str) -> None:
        self.status_var.set(message)
        if getattr(self, "status_dot", None) is not None:
            self.status_dot.set_color(ui.C.success)

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
            if not tags and day % 2 == 0:
                tags.append("stripe")
            tree.insert("", "end", iid=f"d{day}",
                        values=(f"{self.proj.month}月{day}日", weekday_name(self.proj.year, self.proj.month, day),
                                fmt_money(price), note),
                        tags=tuple(tags))
        self.day_card.set_subtitle(f"{self.proj.days_in_month} 天")
        if self.proj.month_price <= 0:
            self.day_info_var.set(f"共 {self.proj.days_in_month} 天 · 尚未设置全月单价")
        else:
            closed = len(self.proj.closed_days())
            self.day_info_var.set(
                f"共 {self.proj.days_in_month} 天 · 每周休息 {closed} 天 · "
                f"单独设置 {count_override} 天"
            )
        self.day_info_strip.set(0, f"{len(self.proj.open_days())} 天")
        self.day_info_strip.set(1, f"{fmt_money(self.proj.month_price)} 元/天")
        self.day_info_strip.set(2, f"{fmt_money(total)} 元")
        self.day_info_strip.set(3, f"{count_override} 天",
                                ui.C.accent if count_override else ui.C.text)

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
        total_open = len(self.proj.open_days())
        for index, sid in enumerate(self.proj.sorted_sids()):
            p = self.proj.roster[sid]
            skip_n = len(set(self.proj.skips.get(sid, set())) & set(self.proj.day_list()))
            values = (p.sid, p.name, p.gender, p.id_card, skip_n,
                      self.proj.ordered_days(sid), fmt_money(self.proj.person_cost(sid)))
            tags = []
            if skip_n >= total_open and skip_n:
                tags.append("optout")
            elif index % 2:
                tags.append("stripe")
            tree.insert("", "end", iid=sid, values=values, tags=tuple(tags))
        self.roster_card.set_subtitle(f"{len(self.proj.roster)} 人")
        self.roster_card.set_empty(not self.proj.roster)

    def refresh_person_list(self) -> None:
        self._person_order = self.proj.sorted_sids()
        self.person_list.delete(0, "end")
        total_open = len(self.proj.open_days())
        for sid in self._person_order:
            p = self.proj.roster[sid]
            skip_n = len(set(self.proj.skips.get(sid, set())) & set(self.proj.day_list()))
            if skip_n and skip_n >= total_open:
                suffix = " · 整月不订"
            elif skip_n:
                suffix = f" · 不订 {skip_n} 天"
            else:
                suffix = ""
            self.person_list.insert("end", f"{p.name}  {sid}{suffix}")
        self.person_list.set_empty(not self._person_order)
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
        win.configure(bg=ui.C.bg)
        win.geometry(f"{ui.S(640)}x{ui.S(460)}")
        card = ui.Card(win, padding=(20, 16))
        card.pack(fill="both", expand=True, padx=ui.S(12), pady=ui.S(12))
        body = card.body
        ui.text_label(body, title, font_key="h2").pack(anchor="w")
        ui.text_label(body, summary, font_key="caption", fg=ui.C.text3,
                      bg=ui.C.card, wrap=560).pack(anchor="w", pady=(ui.S(4), ui.S(10)))
        frame = ui.Surface(body, bg=ui.C.card)
        frame.pack(fill="both", expand=True)
        text = tk.Text(frame, wrap="word", height=12, bd=0, relief="flat",
                       bg=ui.C.card_alt, fg=ui.C.text, font=ui.F("body"),
                       padx=ui.S(10), pady=ui.S(8), highlightthickness=0)
        scroll = ui.FluentScrollbar(frame, command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        text.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        text.insert("1.0", "\n".join(warnings))
        text.configure(state="disabled")
        ui.FluentButton(body, "知道了", kind="accent", command=win.destroy,
                        min_width=ui.S(88)).pack(anchor="e", pady=(ui.S(12), 0))
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
            sub = wd
            if closed:
                sub = "休息"
            elif in_month and today.day == day:
                sub = "今天"
            cb = ui.FluentCheck(self.day_frame, f"{day} 日", var, kind="chip",
                                subtext=sub, min_width=ui.S(78),
                                command=lambda d=day: self._on_day_toggle(d))
            cb.grid(row=(day - 1) // 7, column=(day - 1) % 7, sticky="we",
                    padx=ui.S(3), pady=ui.S(3))
            if closed:
                # 每周休息日不供餐，无需也不能单独勾选“不订”
                cb.state(["disabled"])
            self._day_cbs[day] = cb
        for col in range(7):
            self.day_frame.columnconfigure(col, weight=1, uniform="days")
        self.skip_hint_var.set(f"今天是 {today.year}年{today.month}月{today.day}日" if in_month
                               else f"今天 {today.month}月{today.day}日 不在本工程月份内")
        self._load_skip_state()

    def _update_skip_month_info(self, sid: str) -> None:
        """更新“每日不订”页右侧的整月统计徽标。"""
        total = len(self.proj.open_days())
        skipped = len(set(self.proj.skips.get(sid, set())) & set(self.proj.day_list()))
        if total and skipped >= total:
            text = f"整月都不订 · {skipped} 天"
        elif skipped:
            text = f"已不订 {skipped} 天 / 供餐 {total} 天"
        else:
            text = f"每天订餐 · 供餐 {total} 天"
        self.skip_month_info_var.set(text)
        self.skip_badge.configure_text(text)

    def _load_skip_state(self) -> None:
        self._loading_skip = True
        try:
            sid = self._current_sid
            has_person = bool(sid and sid in self.proj.roster)
            for button in getattr(self, "skip_buttons", []):
                button.set_enabled(has_person)
            if not has_person:
                self.skip_title_var.set("请先在左侧选择人员")
                self.skip_month_info_var.set("")
                self.skip_badge.configure_text("未选择人员")
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
        for index, item in enumerate(rows):
            n = len(set(self.proj.skips.get(item["sid"], set())) & set(self.proj.day_list()))
            if n and n >= total_open:
                note = "整月不订"
            else:
                note = f"不订 {n} 天" if n else ""
            tags = []
            if note == "整月不订":
                tags.append("optout")
            elif index % 2:
                tags.append("stripe")
            tree.insert("", "end", iid=item["sid"],
                        values=(item["sid"], item["name"], item["gender"], item["id_card"],
                                item["days"], note, fmt_money(item["cost"])),
                        tags=tuple(tags))
        closed = len(self.proj.closed_days())
        zero = sum(1 for r in rows if float(r["cost"]) == 0.0)
        extra = f"　费用为 0 的 {zero} 人" if zero else ""
        self.total_var.set(
            f"合计：{fmt_money(self.proj.total_cost())} 元　共 {len(rows)} 人{extra}　"
            f"供餐 {len(self.proj.open_days())} 天" + (f"（每周休息 {closed} 天）" if closed else "")
        )
        self.summary_card.set_subtitle(f"{len(rows)} 人")
        self.summary_card.set_empty(not rows)
        self.summary_strip.set(0, f"{fmt_money(self.proj.total_cost())} 元", ui.C.accent)
        self.summary_strip.set(1, f"{len(rows)} 人")
        self.summary_strip.set(2, f"{len(self.proj.open_days())} 天")
        self.summary_strip.set(3, f"{closed} 天", ui.C.text3 if closed else ui.C.text)

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
        win.configure(bg=ui.C.bg)
        height = min(ui.S(760), self.winfo_screenheight() - ui.S(120))
        win.geometry(f"{ui.S(640)}x{height}")
        win.minsize(ui.S(580), ui.S(520))
        win.resizable(False, True)

        card = ui.Card(win, padding=(20, 16))
        card.pack(fill="both", expand=True, padx=ui.S(12), pady=ui.S(12))
        body = card.body
        ui.text_label(body, "快捷键", font_key="h2").pack(anchor="w")
        ui.text_label(body, "光标在输入框内时，Delete / F2 / 空格不会触发，避免误操作；保存类快捷键始终有效。",
                      font_key="caption", fg=ui.C.text3, bg=ui.C.card, wrap=540).pack(
                          anchor="w", pady=(ui.S(4), ui.S(12)))

        # 先占住底部按钮，内容再多也不会把它挤掉
        footer = ui.Surface(body, bg=ui.C.card)
        footer.pack(side="bottom", fill="x", pady=(ui.S(12), 0))
        ui.FluentButton(footer, "关闭", kind="accent", command=win.destroy,
                        min_width=ui.S(88)).pack(side="right")

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

        grid = ui.Surface(body, bg=ui.C.card)
        grid.pack(fill="both", expand=True)
        row = 0
        for title, items in groups:
            ui.text_label(grid, title, font_key="caption_strong", fg=ui.C.accent).grid(
                row=row, column=0, columnspan=2, sticky="w", pady=(ui.S(8), ui.S(4)))
            row += 1
            for key, desc in items:
                tk.Label(grid, text=key, font=ui.F("mono"), fg=ui.C.text2, bg=ui.C.card,
                         anchor="w").grid(row=row, column=0, sticky="w", padx=(ui.S(2), ui.S(18)))
                tk.Label(grid, text=desc, font=ui.F("body"), fg=ui.C.text, bg=ui.C.card,
                         anchor="w").grid(row=row, column=1, sticky="w")
                row += 1
        grid.columnconfigure(1, weight=1)

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

    # 必须早于创建 Tk，否则高 DPI 屏幕上界面会被系统拉伸得发虚
    ui.enable_dpi_awareness()

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
