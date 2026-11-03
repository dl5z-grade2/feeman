# -*- coding: utf-8 -*-
"""Fluent Design 风格的 Tkinter 基础控件。

只依赖标准库：圆角、描边、悬停 / 按下状态全部手绘在 Canvas 上，
这样在 Windows 10 / 11 上外观一致，也更接近 WinUI 3 的 Fluent 2 视觉。

主要导出：
  * ``C``            颜色令牌（Fluent 2 浅色主题）
  * ``init(root)``   初始化字体与 ttk 样式，必须在创建控件前调用一次
  * ``FluentButton`` 圆角按钮（accent / standard / subtle / danger）
  * ``FluentCheck``  复选框，或 ``kind="chip"`` 的标签式开关
  * ``FluentEntry``  圆角文本框（聚焦时强调色描边）
  * ``FluentSelect`` 圆角下拉选择（自绘弹出列表）
  * ``Card``         圆角卡片容器，``.body`` 里放内容
  * ``TableCard``    带标题的表格卡片（ttk.Treeview + 空状态提示）
  * ``StatCard`` / ``StatStrip``  指标卡片 / 指标条
  * ``TabView``      胶囊式选项卡
  * ``TextArea``     圆角多行文本框
  * ``BrandTile`` / ``divider`` / ``text_label`` 等小部件
"""

from __future__ import annotations

import base64
import ctypes
import math
import struct
import sys
import tkinter as tk
import zlib
from tkinter import font as tkfont
from tkinter import ttk
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

# ---------------------------------------------------------------------------
# 设计令牌
# ---------------------------------------------------------------------------

class C:
    """Fluent 2 浅色主题的颜色令牌。"""

    bg = "#F3F3F3"            # 窗口背景（类似 Mica 的浅灰）
    bg_alt = "#FAFAFA"
    card = "#FFFFFF"          # 卡片 / 面板
    card_alt = "#FAFAFA"
    border = "#E5E5E5"        # 常规描边
    border_strong = "#D6D6D6"  # 控件描边
    border_input = "#C7C7C7"   # 输入类控件的描边
    border_hover = "#B3B3B3"
    text = "#1B1B1B"
    text2 = "#5C5C5C"
    text3 = "#8A8A8A"
    disabled_fill = "#F0F0F0"
    disabled_text = "#A6A6A6"

    accent = "#0F6CBD"
    accent_hover = "#115EA3"
    accent_press = "#0C3B5E"
    accent_tint = "#EBF3FC"
    accent_tint2 = "#DCEBFB"

    success = "#0F7B0F"
    warning = "#9D5D00"
    danger = "#C42B1C"
    danger_tint = "#FDF3F2"

    white = "#FFFFFF"


# 高 DPI 缩放系数：96 DPI 下为 1.0
SCALE: float = 1.0


def S(value: float) -> int:
    """把设计稿里的逻辑像素换算成当前 DPI 下的物理像素。"""
    return int(round(value * SCALE))


_FONTS: Dict[str, tkfont.Font] = {}


def F(name: str = "body") -> tkfont.Font:
    """取一个已注册的字体。"""
    return _FONTS.get(name) or _FONTS["body"]


# ---------------------------------------------------------------------------
# 初始化
# ---------------------------------------------------------------------------

def enable_dpi_awareness() -> None:
    """在创建 Tk 之前调用，让窗口在高 DPI 下不被系统拉伸模糊。"""
    if sys.platform != "win32":
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)   # system aware
    except Exception:  # noqa: BLE001 - 老系统没有 shcore
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:  # noqa: BLE001
            pass


def _pick_family(root: tk.Misc) -> str:
    families = set(tkfont.families(root))
    for name in ("Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑",
                 "Segoe UI", "Noto Sans CJK SC"):
        if name in families:
            return name
    return str(tkfont.nametofont("TkDefaultFont").actual("family"))


def _build_fonts(root: tk.Misc, family: str) -> Dict[str, tkfont.Font]:
    def mk(size: int, weight: str = "normal") -> tkfont.Font:
        return tkfont.Font(root=root, family=family, size=size, weight=weight)

    families = set(tkfont.families(root))
    mono = "Consolas" if "Consolas" in families else "Cascadia Mono" if "Cascadia Mono" in families else family
    return {
        "h1": mk(15, "bold"),
        "h2": mk(12, "bold"),
        "h3": mk(10, "bold"),
        "body": mk(10),
        "body_strong": mk(10, "bold"),
        "caption": mk(9),
        "caption_strong": mk(9, "bold"),
        "metric": mk(24, "bold"),
        "metric_sm": mk(15, "bold"),
        "chip": mk(9),
        "chip_strong": mk(9, "bold"),
        "mono": tkfont.Font(root=root, family=mono, size=9),
    }


def init(root: tk.Misc) -> None:
    """初始化缩放、字体与 ttk 样式（每个进程调用一次）。"""
    global SCALE
    try:
        SCALE = max(1.0, root.winfo_fpixels("1i") / 96.0)
    except tk.TclError:
        SCALE = 1.0

    family = _pick_family(root)
    _FONTS.clear()
    _FONTS.update(_build_fonts(root, family))

    # 只给经典 Tk 控件设默认字体；ttk 控件由样式控制，
    # 避免通配 *Font 把自定义样式（如大号数字）覆盖回 10pt。
    for pattern in ("*Text.font", "*Menu.font", "*Listbox.font",
                    "*TCombobox*Listbox.font"):
        root.option_add(pattern, F("body"))

    _install_ttk(root)


def _install_ttk(root: tk.Misc) -> None:
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    style.configure(".", font=F("body"), background=C.bg, foreground=C.text,
                    focuscolor=C.accent, bordercolor=C.border)

    # --- 表格 -------------------------------------------------------------
    style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
    style.configure(
        "Treeview",
        background=C.card, fieldbackground=C.card, foreground=C.text,
        borderwidth=0, relief="flat", rowheight=S(36), font=F("body"),
    )
    style.map(
        "Treeview",
        background=[("selected", C.accent_tint2)],
        foreground=[("selected", C.text)],
    )
    style.configure(
        "Treeview.Heading",
        background=C.card, foreground=C.text3, relief="flat", borderwidth=0,
        font=F("caption_strong"), padding=(S(10), S(7)),
    )
    style.map("Treeview.Heading", background=[("active", C.card_alt)])
    # clam 默认表头带一块多余的内边距，用自己的布局把高度压到约 30px
    style.layout("Treeview.Heading", [
        ("Treeheading.cell", {"sticky": "nswe"}),
        ("Treeheading.border", {"sticky": "nswe", "children": [
            ("Treeheading.padding", {"sticky": "nswe", "children": [
                ("Treeheading.image", {"side": "right", "sticky": ""}),
                ("Treeheading.text", {"sticky": "we"}),
            ]}),
        ]}),
    ])
    style.configure("Treeheading.cell", background=C.card, borderwidth=0, padding=0)
    style.configure("Treeheading.border", background=C.card, borderwidth=0, relief="flat")
    style.configure("Treeheading.padding", background=C.card, padding=(S(12), S(6)))
    style.configure("Treeheading.text", font=F("caption_strong"), foreground=C.text3)

    # --- 滚动条：去掉箭头，细滚动条 ---------------------------------------
    for orient, axis in (("Vertical", "ns"), ("Horizontal", "ew")):
        layout_name = f"Fluent.{orient}.TScrollbar"
        style.layout(layout_name, [
            (f"{orient}.Scrollbar.trough", {
                "sticky": axis,
                "children": [(f"{orient}.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})],
            })
        ])
        style.configure(
            layout_name,
            background="#C9C9C9", troughcolor=C.card, bordercolor=C.card,
            lightcolor="#C9C9C9", darkcolor="#C9C9C9", gripcount=0,
            relief="flat", arrowsize=S(1), width=S(10),
        )
        style.map(
            layout_name,
            background=[("pressed", "#9A9A9A"), ("active", "#ADADAD")],
        )
        style.layout(f"{orient}.TScrollbar", style.layout(layout_name))
        style.configure(f"{orient}.TScrollbar", **{
            k: v for k, v in style.configure(layout_name).items()
        })


def style_titlebar(root: tk.Misc, color: str = C.bg, text_color: str = "#1B1B1B") -> None:
    """尽量把 Windows 11 的标题栏染成应用背景色（老系统上静默忽略）。"""
    if sys.platform != "win32":
        return
    try:
        root.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(root.winfo_id()) or root.winfo_id()
        if not hwnd:
            return

        def bgr(hex_color: str) -> int:
            r = int(hex_color[1:3], 16)
            g = int(hex_color[3:5], 16)
            b = int(hex_color[5:7], 16)
            return r | (g << 8) | (b << 16)

        dwm = ctypes.windll.dwmapi
        for attr, value in ((35, bgr(color)), (36, bgr(text_color))):
            data = ctypes.c_int(value)
            dwm.DwmSetWindowAttribute(hwnd, attr, ctypes.byref(data), ctypes.sizeof(data))
    except Exception:  # noqa: BLE001 - 纯装饰，失败不影响功能
        pass


# ---------------------------------------------------------------------------
# 绘制小工具
# ---------------------------------------------------------------------------

def _hex_rgb(color: str) -> Tuple[int, int, int]:
    value = color.lstrip("#")
    if len(value) == 3:
        value = "".join(ch * 2 for ch in value)
    return int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16)


def _png_bytes(width: int, height: int, pixels: bytes) -> bytes:
    raw = b"".join(b"\x00" + pixels[y * width * 4:(y + 1) * width * 4]
                   for y in range(height))

    def chunk(tag: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))

    header = struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header)
            + chunk(b"IDAT", zlib.compress(raw, 6)) + chunk(b"IEND", b""))


_ROUND_IMAGES: Dict[Tuple[int, int, int, str, str, int], tk.PhotoImage] = {}
_ROUND_IMAGES_INTERP: Any = None


def _rounded_image(master: tk.Misc, width: int, height: int, radius: int,
                   fill: str, outline: str, border_width: int) -> tk.PhotoImage:
    """生成一张带抗锯齿的圆角矩形 PNG 并缓存。

    只用于小尺寸控件（按钮、标签、输入框等）；大卡片仍用多边形绘制，
    避免在纯 Python 里逐像素渲染超大画布。
    """
    global _ROUND_IMAGES_INTERP
    interp = master.tk
    if interp is not _ROUND_IMAGES_INTERP:
        # 换了一个 Tk 解释器（例如测试里重建窗口），旧图片不能再用
        _ROUND_IMAGES.clear()
        _ROUND_IMAGES_INTERP = interp

    key = (width, height, radius, fill, outline, border_width)
    cached = _ROUND_IMAGES.get(key)
    if cached is not None:
        return cached

    fill_rgb = _hex_rgb(fill)
    out_rgb = _hex_rgb(outline)
    samples = 3
    total = samples * samples
    hw, hh = width / 2.0, height / 2.0
    inner = radius - border_width
    radius = min(radius, hw, hh)
    inner = max(min(inner, radius), 0.0)
    pixels = bytearray()

    for y in range(height):
        for x in range(width):
            outer = 0
            core = 0
            for sy in range(samples):
                py = y + (sy + 0.5) / samples
                qy = abs(py - hh) - (hh - radius)
                for sx in range(samples):
                    px = x + (sx + 0.5) / samples
                    qx = abs(px - hw) - (hw - radius)
                    dist = (math.hypot(max(qx, 0.0), max(qy, 0.0))
                            + min(max(qx, qy), 0.0) - radius)
                    if dist <= 0.0:
                        outer += 1
                        if dist <= -border_width:
                            core += 1
            if outer == 0:
                pixels += b"\x00\x00\x00\x00"
                continue
            edge = outer - core
            red = (fill_rgb[0] * core + out_rgb[0] * edge) // outer
            green = (fill_rgb[1] * core + out_rgb[1] * edge) // outer
            blue = (fill_rgb[2] * core + out_rgb[2] * edge) // outer
            alpha = 255 * outer // total
            pixels += bytes((red, green, blue, alpha))

    image = tk.PhotoImage(master=master, data=base64.b64encode(
        _png_bytes(width, height, bytes(pixels))).decode("ascii"))
    if len(_ROUND_IMAGES) > 400:
        # 拖动窗口时会不断产生新尺寸，限制缓存规模，避免内存无限增长
        _ROUND_IMAGES.clear()
    _ROUND_IMAGES[key] = image
    return image


def round_rect(canvas: tk.Canvas, x1: float, y1: float, x2: float, y2: float,
               radius: float, *, steps: int = 8, **kwargs: Any) -> int:
    """在画布上画圆角矩形。

    小尺寸图形用带 alpha 的 PNG 贴图渲染，圆角和描边都带抗锯齿；
    大尺寸容器用 4 段圆弧采样成多边形（纯 Tk，不占内存）。
    """
    fill = kwargs.get("fill")
    outline = kwargs.get("outline")
    border_width = int(kwargs.get("width") or 1)
    width = int(round(x2 - x1))
    height = int(round(y2 - y1))
    radius = max(0.0, min(radius, width / 2, height / 2))
    if radius <= 0.5:
        return canvas.create_rectangle(x1, y1, x2, y2, **kwargs)

    tags = kwargs.get("tags")
    smoothable = (
        isinstance(fill, str) and fill.startswith("#")
        and (outline is None or (isinstance(outline, str) and outline.startswith("#")))
        and width >= 6 and height >= 6
        and width * height <= 24000
    )
    if smoothable:
        image = _rounded_image(canvas, width, height, int(round(radius)),
                               fill, outline or fill, max(border_width, 1))
        item = canvas.create_image(x1, y1, image=image, anchor="nw")
        if tags:
            canvas.itemconfigure(item, tags=tags)
        return item

    corners = (
        (x2 - radius, y1 + radius, -90.0, 0.0),    # 右上
        (x2 - radius, y2 - radius, 0.0, 90.0),     # 右下
        (x1 + radius, y2 - radius, 90.0, 180.0),   # 左下
        (x1 + radius, y1 + radius, 180.0, 270.0),  # 左上
    )
    points: List[float] = []
    for cx, cy, start, end in corners:
        for i in range(steps + 1):
            angle = math.radians(start + (end - start) * i / steps)
            points.append(cx + radius * math.cos(angle))
            points.append(cy + radius * math.sin(angle))
    return canvas.create_polygon(points, smooth=False, **kwargs)


def surface_of(widget: tk.Misc) -> str:
    """向上查找控件所在的背景色，用于让自绘控件“透明”融入父容器。"""
    w: Optional[tk.Misc] = widget
    while w is not None:
        color = getattr(w, "_fluent_surface", None)
        if color:
            return str(color)
        try:
            bg = w.cget("bg")
            if bg:
                return str(bg)
        except (tk.TclError, AttributeError):
            pass
        w = getattr(w, "master", None)
    return C.bg


class Surface(tk.Frame):
    """带背景色的普通容器，子控件可以据此取到“透明”背景。"""

    def __init__(self, master: tk.Misc, bg: str = C.bg, **kwargs: Any) -> None:
        kwargs.setdefault("bg", bg)
        super().__init__(master, **kwargs)
        self._fluent_surface = bg


def divider(master: tk.Misc, color: str = C.border) -> tk.Frame:
    line = tk.Frame(master, bg=color, height=S(1))
    line._fluent_surface = color  # type: ignore[attr-defined]
    return line


def text_label(master: tk.Misc, text: Optional[str] = None, *,
               textvariable: Optional[tk.Variable] = None,
               font_key: str = "body", fg: Optional[str] = None,
               bg: Optional[str] = None, wrap: Optional[int] = None,
               anchor: str = "w", justify: str = "left",
               cursor: Optional[str] = None) -> tk.Label:
    label = tk.Label(
        master, text=text, textvariable=textvariable, font=F(font_key),
        fg=fg or C.text, bg=bg or surface_of(master), anchor=anchor,
        justify=justify, bd=0, padx=0, pady=0,
    )
    if wrap:
        label.configure(wraplength=S(wrap))
    if cursor:
        label.configure(cursor=cursor)
    return label


# ---------------------------------------------------------------------------
# 按钮
# ---------------------------------------------------------------------------

class FluentButton(tk.Canvas):
    """圆角按钮。kind: accent / standard / subtle / danger。"""

    _STYLES: Dict[str, Dict[str, Optional[str]]] = {
        "accent": dict(fill=C.accent, hover=C.accent_hover, press=C.accent_press,
                       fg="#FFFFFF", border=C.accent, hover_border=C.accent_hover),
        "standard": dict(fill=C.card, hover="#F7F7F7", press="#EFEFEF",
                         fg=C.text, border=C.border_strong, hover_border=C.border_hover),
        "subtle": dict(fill=None, hover="#E7E7E7", press="#DCDCDC",
                       fg=C.text2, border=None, hover_border=None),
        "danger": dict(fill=C.card, hover=C.danger_tint, press="#FBE7E5",
                       fg=C.danger, border=C.border_strong, hover_border="#E0B4AF"),
    }

    def __init__(self, master: tk.Misc, text: str = "", command: Optional[Callable[[], Any]] = None,
                 kind: str = "standard", *, surface: Optional[str] = None,
                 min_width: Optional[int] = None, height: int = 32, radius: int = 4,
                 pad_x: int = 16, font_key: Optional[str] = None,
                 enabled: bool = True, takefocus: bool = True) -> None:
        self._surface = surface or surface_of(master)
        self._kind = kind if kind in self._STYLES else "standard"
        self._spec = self._STYLES[self._kind]
        self._font = F(font_key or ("body_strong" if kind == "accent" else "body"))
        self._pad_x = S(pad_x)
        self._radius = radius
        self._h = S(height)
        self._text = text
        self._command = command
        self._enabled = bool(enabled)
        self._hover = False
        self._pressed = False
        self._focus = False
        width = max(S(min_width or 0), self._font.measure(text) + self._pad_x * 2 + S(2))
        super().__init__(master, width=width, height=self._h, bg=self._surface,
                         highlightthickness=0, bd=0, takefocus=1 if takefocus else 0)

        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)
        self.bind("<FocusIn>", self._on_focus_in)
        self.bind("<FocusOut>", self._on_focus_out)
        self.bind("<Return>", self._on_key)
        self.bind("<space>", self._on_key)
        self._draw()

    def _on_key(self, _event: tk.Event) -> str:
        self.invoke()
        return "break"

    # -- 对外接口 ------------------------------------------------------
    def invoke(self) -> Any:
        if self._enabled and self._command is not None:
            return self._command()
        return None

    def set_enabled(self, enabled: bool) -> None:
        enabled = bool(enabled)
        if enabled != self._enabled:
            self._enabled = enabled
            self.configure(cursor="hand2" if enabled else "")
            self._draw()

    def set_text(self, text: str) -> None:
        self._text = text
        need = self._font.measure(text) + self._pad_x * 2 + S(2)
        if need > int(self.cget("width")):
            self.configure(width=need)
        self._draw()

    # -- 事件 ----------------------------------------------------------
    def _on_enter(self, _event: tk.Event) -> None:
        self._hover = True
        self._draw()

    def _on_leave(self, _event: tk.Event) -> None:
        self._hover = False
        self._pressed = False
        self._draw()

    def _on_press(self, _event: tk.Event) -> None:
        if not self._enabled:
            return
        self._pressed = True
        self.focus_set()
        self._draw()

    def _on_release(self, event: tk.Event) -> None:
        if not self._enabled or not self._pressed:
            return
        self._pressed = False
        self._draw()
        if 0 <= event.x <= self.winfo_width() and 0 <= event.y <= self.winfo_height():
            self.invoke()

    def _on_focus_in(self, _event: tk.Event) -> None:
        self._focus = True
        self._draw()

    def _on_focus_out(self, _event: tk.Event) -> None:
        self._focus = False
        self._draw()

    # -- 绘制 ----------------------------------------------------------
    def _draw(self) -> None:
        w = max(self.winfo_width(), self.winfo_reqwidth())
        h = max(self.winfo_height(), self.winfo_reqheight())
        if w < 4 or h < 4:
            return
        self.delete("all")
        spec = self._spec
        fill = spec["fill"] or self._surface
        border = spec["border"] or fill
        fg = spec["fg"] or C.text

        if not self._enabled:
            fill, border, fg = C.disabled_fill, C.border, C.disabled_text
        elif self._pressed and spec["press"]:
            fill = spec["press"] or fill
        elif self._hover and spec["hover"]:
            fill = spec["hover"]
            if spec["hover_border"]:
                border = spec["hover_border"]

        r = min(float(S(self._radius)), h / 2)
        round_rect(self, 0.5, 0.5, w - 1.5, h - 1.5, r, fill=fill,
                   outline=border, width=1)
        if self._focus and self._enabled:
            round_rect(self, 1.5, 1.5, w - 2.5, h - 2.5, max(r - 1, 1),
                       fill="", outline=C.accent, width=2)
        self.create_text(w / 2, h / 2 + 0.5, text=self._text, font=self._font,
                         fill=fg, anchor="center")


# ---------------------------------------------------------------------------
# 复选框 / 标签开关
# ---------------------------------------------------------------------------

class FluentCheck(tk.Canvas):
    """Fluent 复选框；``kind="chip"`` 时是圆角的标签式开关。"""

    def __init__(self, master: tk.Misc, text: str = "",
                 variable: Optional[tk.Variable] = None,
                 command: Optional[Callable[[], Any]] = None, *,
                 kind: str = "check", subtext: Optional[str] = None,
                 surface: Optional[str] = None, min_width: int = 0,
                 height: Optional[int] = None, pad_x: int = 12,
                 font_key: str = "body", gap: int = 10,
                 enabled: bool = True) -> None:
        self._surface = surface or surface_of(master)
        self._kind = kind
        self._text = text
        self._subtext = subtext
        self._var = variable
        self._command = command
        self._enabled = bool(enabled)
        self._hover = False
        self._focus = False
        self._font = F(font_key if kind == "check" else "chip")
        self._gap = S(gap)
        self._pad_x = S(pad_x)
        self._radius = 4

        if kind == "chip":
            h = S(height or (44 if subtext else 30))
            base = max(self._font.measure(text), F("chip").measure(subtext or ""))
            width = max(S(min_width), base + self._pad_x * 2 + S(4))
        else:
            h = S(height or 24)
            width = S(18) + self._gap + self._font.measure(text) + S(4)

        super().__init__(master, width=width, height=h, bg=self._surface,
                         highlightthickness=0, bd=0, takefocus=1)
        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Enter>", lambda e: self._hover_set(True))
        self.bind("<Leave>", lambda e: self._hover_set(False))
        self.bind("<Button-1>", self._on_click)
        self.bind("<FocusIn>", lambda e: self._focus_set(True))
        self.bind("<FocusOut>", lambda e: self._focus_set(False))
        self.bind("<Return>", self._on_key)
        self.bind("<space>", self._on_key)
        if variable is not None:
            try:
                variable.trace_add("write", lambda *_: self._draw())
            except tk.TclError:
                pass
        self._draw()

    def _on_key(self, event: tk.Event) -> str:
        self._on_click(event)
        return "break"

    # -- 对外接口 ------------------------------------------------------
    def get(self) -> bool:
        return bool(self._var.get()) if self._var is not None else False

    def set_enabled(self, enabled: bool) -> None:
        enabled = bool(enabled)
        if enabled != self._enabled:
            self._enabled = enabled
            self.configure(cursor="hand2" if enabled else "")
            self._draw()

    def state(self, spec: Optional[Sequence[str]] = None) -> Any:
        """兼容 ttk 的 ``state(["disabled"])`` / ``state(["!disabled"])``。"""
        if spec is None:
            return () if self._enabled else ("disabled",)
        for token in spec:
            if token.lstrip("!") == "disabled":
                self.set_enabled(token.startswith("!"))
        return ()

    def configure_text(self, text: str) -> None:
        """只改文字、不改尺寸（用于字数会变的徽标）。"""
        self._text = text
        self._draw()

    def configure(self, cnf: Any = None, **kwargs: Any) -> Any:  # noqa: D401
        if isinstance(cnf, dict) and "text" in cnf:
            self._text = cnf.pop("text")
        if "text" in kwargs:
            self._text = kwargs.pop("text")
        return super().configure(cnf, **kwargs)

    # -- 事件 ----------------------------------------------------------
    def _hover_set(self, value: bool) -> None:
        self._hover = value
        self._draw()

    def _focus_set(self, value: bool) -> None:
        self._focus = value
        self._draw()

    def _on_click(self, _event: Optional[tk.Event] = None) -> None:
        if not self._enabled:
            return
        if self._var is not None:
            self._var.set(not bool(self._var.get()))
        self.focus_set()
        if self._command is not None:
            self._command()
        self._draw()

    # -- 绘制 ----------------------------------------------------------
    def _draw(self) -> None:
        w = max(self.winfo_width(), self.winfo_reqwidth())
        h = max(self.winfo_height(), self.winfo_reqheight())
        if w < 4 or h < 4:
            return
        self.delete("all")
        checked = self.get()
        enabled = self._enabled
        if self._kind == "chip":
            self._draw_chip(w, h, checked, enabled)
        else:
            self._draw_check(w, h, checked, enabled)

    def _draw_chip(self, w: int, h: int, checked: bool, enabled: bool) -> None:
        if not enabled:
            fill, border = C.disabled_fill, C.border
            fg, fg2 = C.disabled_text, C.disabled_text
        elif checked:
            fill, border = C.accent_tint, C.accent
            fg, fg2 = C.accent, "#3B6FA8"
        elif self._hover:
            fill, border = "#EFEFEF", C.border_hover
            fg, fg2 = C.text, C.text3
        elif self._focus:
            fill, border = C.card, C.accent
            fg, fg2 = C.text, C.text3
        else:
            fill, border = self._surface, C.border_strong
            fg, fg2 = C.text2, C.text3

        if fill == self._surface:
            self.create_rectangle(0, 0, w, h, fill=self._surface, outline="")
        round_rect(self, 0.5, 0.5, w - 1.5, h - 1.5, S(self._radius),
                   fill=fill, outline=border, width=1)
        if self._subtext:
            self.create_text(w / 2, h / 2 - S(7), text=self._text,
                             font=F("chip_strong" if checked else "body_strong"),
                             fill=fg, anchor="center")
            self.create_text(w / 2, h / 2 + S(9), text=self._subtext,
                             font=F("chip"), fill=fg2, anchor="center")
        else:
            self.create_text(w / 2, h / 2 + 0.5, text=self._text,
                             font=F("chip_strong" if checked else "chip"),
                             fill=fg, anchor="center")

    def _draw_check(self, w: int, h: int, checked: bool, enabled: bool) -> None:
        box = S(18)
        top = (h - box) / 2
        if not enabled:
            box_fill, box_border = C.disabled_fill, C.border_strong
            fg = C.disabled_text
        else:
            box_fill = C.accent if checked else C.card
            box_border = C.accent if checked else (C.border_hover if self._hover else C.text3)
            fg = C.text if self._hover or checked else C.text2
        round_rect(self, 0.5, top + 0.5, box + 0.5, top + box + 0.5, S(4),
                   fill=box_fill, outline=box_border, width=1)
        if checked:
            pad = S(4)
            self.create_line(pad + 1, top + box * 0.55,
                             box * 0.42, top + box - pad,
                             box - pad + 1, top + pad,
                             fill="#FFFFFF", width=max(S(2), 2),
                             capstyle="round", joinstyle="round")
        if self._focus and enabled:
            round_rect(self, 0.5, top + 0.5, box + 0.5, top + box + 0.5, S(4),
                       fill="", outline=C.accent, width=2)
            self.create_line(box + 1, top + box / 2, w - 1, top + box / 2,
                             fill=C.accent, width=1, dash=(2, 2))
        self.create_text(box + self._gap, h / 2 + 0.5, text=self._text,
                         font=self._font, fill=fg, anchor="w")


# ---------------------------------------------------------------------------
# 文本框
# ---------------------------------------------------------------------------

class FluentEntry(tk.Canvas):
    """圆角单行文本框，聚焦时显示强调色描边。"""

    def __init__(self, master: tk.Misc, textvariable: Optional[tk.Variable] = None, *,
                 width: int = 20, surface: Optional[str] = None, height: int = 34,
                 radius: int = 4, font_key: str = "body", justify: str = "left",
                 show: Optional[str] = None, enabled: bool = True) -> None:
        self._surface = surface or surface_of(master)
        self._font = F(font_key)
        self._radius = radius
        self._h = S(height)
        self._hover = False
        self._focus = False
        self._enabled = enabled
        w = int(width * max(self._font.measure("0"), self._font.measure("字") * 0.62)) + S(26)
        super().__init__(master, width=w, height=self._h, bg=self._surface,
                         highlightthickness=0, bd=0, takefocus=0)

        self.entry = tk.Entry(
            self, textvariable=textvariable, bd=0, relief="flat", bg=C.card,
            fg=C.text, insertbackground=C.text, font=self._font, justify=justify,
            highlightthickness=0, disabledbackground=C.disabled_fill,
            disabledforeground=C.disabled_text, readonlybackground=C.card,
        )
        if show:
            self.entry.configure(show=show)
        if not enabled:
            self.entry.configure(state="disabled")
        self._win = self.create_window(S(13), self._h // 2, window=self.entry,
                                       anchor="w", width=w - S(26), height=self._h - S(10))
        self.bind("<Configure>", self._on_configure)
        self.bind("<Button-1>", lambda e: self.entry.focus_set())
        self.bind("<Enter>", lambda e: self._set_hover(True))
        self.bind("<Leave>", lambda e: self._set_hover(False))
        self.entry.bind("<FocusIn>", lambda e: self._set_focus(True), add="+")
        self.entry.bind("<FocusOut>", lambda e: self._set_focus(False), add="+")
        self.entry.bind("<Enter>", lambda e: self._set_hover(True), add="+")
        self.entry.bind("<Leave>", lambda e: self._set_hover(False), add="+")
        self._draw()

    # -- 对外接口 ------------------------------------------------------
    def get(self) -> str:
        return self.entry.get()

    def set(self, value: str) -> None:
        self.entry.delete(0, "end")
        self.entry.insert(0, value)

    def focus_set(self) -> None:  # noqa: D401 - 与 tk 接口同名
        self.entry.focus_set()

    def set_enabled(self, enabled: bool) -> None:
        self._enabled = enabled
        self.entry.configure(state="normal" if enabled else "disabled")
        self._draw()

    # -- 事件 ----------------------------------------------------------
    def _on_configure(self, _event: tk.Event) -> None:
        w = self.winfo_width()
        h = self.winfo_height()
        self.coords(self._win, S(13), h // 2)
        self.itemconfigure(self._win, width=max(w - S(26), S(10)), height=max(h - S(10), S(8)))
        self._draw()

    def _set_hover(self, value: bool) -> None:
        self._hover = value
        self._draw()

    def _set_focus(self, value: bool) -> None:
        self._focus = value
        self._draw()

    # -- 绘制 ----------------------------------------------------------
    def _draw(self) -> None:
        w = max(self.winfo_width(), self.winfo_reqwidth())
        h = max(self.winfo_height(), self.winfo_reqheight())
        if w < 4 or h < 4:
            return
        self.delete("bg")
        if not self._enabled:
            border, fill, width_px = C.border, C.disabled_fill, 1
        elif self._focus:
            border, fill, width_px = C.accent, C.card, 2
        elif self._hover:
            border, fill, width_px = C.border_hover, C.card, 1
        else:
            border, fill, width_px = C.border_input, C.card, 1
        self.entry.configure(bg=fill)
        round_rect(self, width_px / 2, width_px / 2, w - width_px / 2, h - width_px / 2,
                   S(self._radius), fill=fill, outline=border, width=width_px, tags="bg")
        self.tag_lower("bg")


# ---------------------------------------------------------------------------
# 下拉选择
# ---------------------------------------------------------------------------

class _Dropdown(tk.Toplevel):
    """FluentSelect 的弹出列表。"""

    def __init__(self, owner: "FluentSelect", values: Sequence[str], current: str) -> None:
        super().__init__(owner)
        self.withdraw()
        self.overrideredirect(True)
        try:
            self.attributes("-topmost", True)
        except tk.TclError:
            pass
        self.owner = owner
        self.values = list(values)
        self._index = self.values.index(current) if current in self.values else -1

        outer = tk.Frame(self, bg=C.border_strong)
        outer.pack(fill="both", expand=True)
        inner = tk.Frame(outer, bg=C.card)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        outer.bind("<Button-1>", lambda e: self.close())

        self.rows: List[Tuple[tk.Frame, tk.Label]] = []
        for i, value in enumerate(values):
            row = tk.Frame(inner, bg=C.card, cursor="hand2")
            row.pack(fill="x")
            label = tk.Label(row, text=value, bg=C.card, fg=C.text, font=F("body"),
                             anchor="w", padx=S(12), pady=S(7))
            label.pack(fill="x")
            for widget in (row, label):
                widget.bind("<Enter>", lambda e, i=i: self._hover(i))
                widget.bind("<Leave>", lambda e, i=i: self._hover(-1))
                widget.bind("<Button-1>", lambda e, i=i: self._choose(i))
            self.rows.append((row, label))

        self.bind("<Escape>", lambda e: self.close())
        self.bind("<Down>", lambda e: self._move(1))
        self.bind("<Up>", lambda e: self._move(-1))
        self.bind("<Return>", lambda e: self._choose(self._index if self._index >= 0 else 0))
        self.bind("<FocusOut>", lambda e: self.after(1, self._close_if_unfocused))

        self.update_idletasks()
        width = max(owner.winfo_width(), self.winfo_reqwidth())
        height = self.winfo_reqheight()
        x = owner.winfo_rootx()
        y = owner.winfo_rooty() + owner.winfo_height() + S(4)
        screen_h = self.winfo_screenheight()
        if y + height > screen_h - S(8):
            # 下方空间不足时向上弹出，但不能盖住输入框本身
            above = owner.winfo_rooty() - height - S(4)
            if above >= S(8):
                y = above
            else:
                y = max(S(8), min(owner.winfo_rooty(), screen_h - height - S(8)))
        x = max(S(4), min(x, self.winfo_screenwidth() - width - S(4)))
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.deiconify()
        self.lift()
        try:
            self.grab_set()
        except tk.TclError:
            pass
        self.focus_set()
        self._hover(self._index)

    def _hover(self, index: int) -> None:
        for i, (row, label) in enumerate(self.rows):
            if i == index:
                row.configure(bg=C.accent_tint)
                label.configure(bg=C.accent_tint, fg=C.accent)
            else:
                selected = i == self._index
                row.configure(bg=C.card)
                label.configure(bg=C.card, fg=C.accent if selected else C.text)

    def _move(self, delta: int) -> None:
        if not self.values:
            return
        index = min(max(self._index + delta, 0), len(self.values) - 1)
        self._index = index
        self._hover(index)

    def _choose(self, index: int) -> None:
        if 0 <= index < len(self.values):
            self.owner._pick(self.values[index])
        self.close()

    def close(self) -> None:
        try:
            self.grab_release()
        except tk.TclError:
            pass
        owner = self.owner
        self.destroy()
        try:
            owner._popup_closed()
        except tk.TclError:
            pass

    def _close_if_unfocused(self) -> None:
        try:
            if self.focus_displayof() is None:
                self.close()
        except tk.TclError:
            self.close()


class FluentSelect(tk.Canvas):
    """圆角下拉选择器（只读，风格与 Fluent 的 ComboBox 一致）。"""

    def __init__(self, master: tk.Misc, variable: tk.Variable,
                 values: Sequence[str], *, width: Optional[int] = None,
                 surface: Optional[str] = None, height: int = 34,
                 font_key: str = "body", command: Optional[Callable[[str], Any]] = None) -> None:
        self._surface = surface or surface_of(master)
        self._font = F(font_key)
        self._var = variable
        self._values = list(values)
        self._command = command
        self._hover = False
        self._open = False
        self._popup: Optional["_Dropdown"] = None
        self._h = S(height)
        chars = self._font.measure("0") * 8
        w = S(width) if width else chars + S(44)
        super().__init__(master, width=w, height=self._h, bg=self._surface,
                         highlightthickness=0, bd=0, takefocus=1, cursor="hand2")
        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Enter>", lambda e: self._set_hover(True))
        self.bind("<Leave>", lambda e: self._set_hover(False))
        self.bind("<Button-1>", lambda e: self.open())
        self.bind("<Return>", self._on_key)
        self.bind("<space>", self._on_key)
        try:
            variable.trace_add("write", lambda *_: self._draw())
        except tk.TclError:
            pass
        self._draw()

    # -- 对外接口 ------------------------------------------------------
    def set_values(self, values: Sequence[str]) -> None:
        self._values = list(values)
        self._draw()

    def _on_key(self, _event: tk.Event) -> str:
        self.open()
        return "break"

    def get(self) -> str:
        return str(self._var.get())

    def set(self, value: str) -> None:
        self._var.set(value)

    def open(self) -> None:
        if self._open:
            return
        self._open = True
        self._draw()
        self._popup = _Dropdown(self, self._values, self.get())

    def _reset_open(self) -> None:
        self._open = False
        self._draw()

    def _pick(self, value: str) -> None:
        self._var.set(value)
        self._open = False
        self._draw()
        if self._command is not None:
            self._command(value)

    def _popup_closed(self) -> None:
        self._open = False
        self._popup = None
        self._draw()

    def _set_hover(self, value: bool) -> None:
        self._hover = value
        self._draw()

    # -- 绘制 ----------------------------------------------------------
    def _draw(self) -> None:
        w = max(self.winfo_width(), self.winfo_reqwidth())
        h = max(self.winfo_height(), self.winfo_reqheight())
        if w < 4 or h < 4:
            return
        self.delete("all")
        border = C.accent if self._open else (C.border_hover if self._hover else C.border_input)
        round_rect(self, 0.5, 0.5, w - 1.5, h - 1.5, S(4), fill=C.card,
                   outline=border, width=1)
        self.create_text(S(12), h / 2 + 0.5, text=self.get(), font=self._font,
                         fill=C.text, anchor="w")
        # 自绘的“v”形箭头，避免依赖字体里的符号
        cx, cy = w - S(14), h / 2
        self.create_line(cx - S(4), cy - S(2), cx, cy + S(2.5), cx + S(4), cy - S(2),
                         fill=C.text2, width=max(S(1.5), 1), capstyle="round",
                         joinstyle="round")


# ---------------------------------------------------------------------------
# 卡片
# ---------------------------------------------------------------------------

class Card(tk.Frame):
    """圆角卡片：``canvas`` 画底与描边，``body`` 放内容。"""

    def __init__(self, master: tk.Misc, *, padding: Any = (16, 14), radius: int = 8,
                 fill: str = C.card, border: str = C.border,
                 surface: Optional[str] = None, title: Optional[str] = None,
                 subtitle: Optional[str] = None) -> None:
        self._surface = surface or surface_of(master)
        self._fill = fill
        self._border = border
        self._radius = radius
        self._pad = padding if isinstance(padding, tuple) else (padding, padding)
        super().__init__(master, bg=self._surface)

        self._canvas = tk.Canvas(self, bg=self._surface, highlightthickness=0, bd=0)
        self._canvas.place(x=0, y=0, relwidth=1, relheight=1)
        self._canvas.bind("<Configure>", lambda e: self._redraw())

        self.body = tk.Frame(self, bg=fill)
        self.body._fluent_surface = fill  # type: ignore[attr-defined]
        self.body.pack(fill="both", expand=True,
                       padx=S(self._pad[0]), pady=S(self._pad[1]))

        self.header: Optional[tk.Frame] = None
        self.actions: Optional[tk.Frame] = None
        self.subtitle_label: Optional[tk.Label] = None
        self._subtitle_var = tk.StringVar(value=subtitle or "")
        if title is not None:
            self.header = tk.Frame(self.body, bg=fill)
            self.header.pack(fill="x", pady=(0, S(10)))
            text_label(self.header, title, font_key="h2", bg=fill).pack(side="left")
            self.actions = tk.Frame(self.header, bg=fill)
            self.actions.pack(side="right")
            if subtitle is not None:
                self._build_subtitle()

        self._redraw()

    def _build_subtitle(self) -> None:
        assert self.header is not None
        self.subtitle_label = text_label(
            self.header, textvariable=self._subtitle_var, font_key="caption",
            fg=C.text3, bg=self._fill)
        self.subtitle_label.pack(side="left", padx=(S(10), 0))

    def set_subtitle(self, text: str) -> None:
        if self.subtitle_label is None and self.header is not None:
            self._build_subtitle()
        self._subtitle_var.set(text)

    def _redraw(self) -> None:
        w = self._canvas.winfo_width()
        h = self._canvas.winfo_height()
        if w < 4 or h < 4:
            return
        if getattr(self, "_drawn_size", None) == (w, h):
            return
        self._drawn_size = (w, h)
        self._canvas.delete("all")
        round_rect(self._canvas, 0.5, 0.5, w - 0.5, h - 0.5, S(self._radius),
                   fill=self._fill, outline=self._border, width=1)


class TableCard(Card):
    """带标题的表格卡片，自带空状态提示。"""

    def __init__(self, master: tk.Misc, columns: Dict[str, Tuple[str, int, str]], *,
                 title: str = "", subtitle: str = "", surface: Optional[str] = None,
                 height: int = 6, selectmode: str = "browse",
                 empty_text: str = "暂无数据",
                 border: Optional[str] = None,
                 on_double: Optional[Callable[[tk.Event], Any]] = None) -> None:
        super().__init__(master, padding=(0, 14), surface=surface,
                         title=title or None, subtitle=subtitle or None,
                         border=border or C.border)
        # 让表格贴到卡片左右内边距，标题仍缩进
        if self.header is not None:
            self.header.pack_configure(padx=S(16))
        self.body.pack_configure(padx=0)

        # 表头下方的分隔线 + 真正的表格容器（Treeview 必须是 holder 的子控件，
        # 否则 pack 会把它排到卡片里而不是表格区域）
        divider(self.body, C.border).pack(fill="x", pady=(0, 0))
        holder = tk.Frame(self.body, bg=C.card)
        holder._fluent_surface = C.card  # type: ignore[attr-defined]
        holder.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(holder, columns=list(columns.keys()),
                                 show="headings", selectmode=selectmode, height=height)
        for name, (heading, width, anchor) in columns.items():
            self.tree.heading(name, text=heading, anchor=anchor)
            self.tree.column(name, width=S(width), minwidth=S(40), anchor=anchor,
                             stretch=(anchor == "w"))
        self.scroll = FluentScrollbar(holder, command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.scroll.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(S(6), 0))
        self.scroll.pack(side="right", fill="y", padx=(0, S(4)))

        self._empty = tk.Label(holder, text=empty_text, bg=C.card, fg=C.text3,
                               font=F("caption"))
        self._empty_visible = False
        if on_double is not None:
            self.tree.bind("<Double-1>", on_double)

    def set_empty(self, visible: bool, text: Optional[str] = None) -> None:
        if text is not None:
            self._empty.configure(text=text)
        if visible:
            if not self._empty_visible:
                self._empty.place(relx=0.5, rely=0.55, anchor="center")
                self._empty_visible = True
        elif self._empty_visible:
            self._empty.place_forget()
            self._empty_visible = False


class StatCard(Card):
    """一个指标卡片：说明文字 + 大号数值 + 附属说明。"""

    def __init__(self, master: tk.Misc, label: str, *, unit: str = "",
                 value: str = "—", caption: str = "", accent: bool = False,
                 surface: Optional[str] = None) -> None:
        super().__init__(master, padding=(18, 14), surface=surface)
        self._value_color = C.accent if accent else C.text
        text_label(self.body, label, font_key="caption", fg=C.text3).pack(anchor="w")

        row = tk.Frame(self.body, bg=C.card)
        row.pack(anchor="w", fill="x", pady=(S(2), 0))
        self.value_label = tk.Label(row, text=value, font=F("metric"),
                                    fg=self._value_color, bg=C.card, anchor="w")
        self.value_label.pack(side="left")
        self.unit_label = tk.Label(row, text=unit, font=F("caption"),
                                   fg=C.text3, bg=C.card, anchor="sw", pady=S(3))
        self.unit_label.pack(side="left", padx=(S(4), 0))

        self.caption_label = tk.Label(self.body, text=caption, font=F("caption"),
                                      fg=C.text2, bg=C.card, anchor="w",
                                      justify="left", wraplength=S(240))
        self.caption_label.pack(anchor="w", fill="x")
        self.body.bind("<Configure>", self._on_resize)

    def _on_resize(self, event: tk.Event) -> None:
        self.caption_label.configure(wraplength=max(event.width - S(4), S(80)))

    def set(self, value: Optional[str] = None, caption: Optional[str] = None,
            unit: Optional[str] = None, color: Optional[str] = None) -> None:
        if value is not None:
            self.value_label.configure(text=value)
        if caption is not None:
            self.caption_label.configure(text=caption)
        if unit is not None:
            self.unit_label.configure(text=unit)
        if color is not None:
            self.value_label.configure(fg=color)


class StatStrip(Card):
    """横向排列的一组小指标（标签在左、数值在右）。"""

    def __init__(self, master: tk.Misc, labels: Sequence[str], *,
                 surface: Optional[str] = None) -> None:
        super().__init__(master, padding=(4, 10), surface=surface)
        self._values: List[tk.Label] = []
        for i, label in enumerate(labels):
            if i:
                divider(self.body, C.border).grid(row=0, column=2 * i - 1, sticky="ns",
                                                  padx=S(4), pady=S(2))
            cell = tk.Frame(self.body, bg=C.card)
            cell.grid(row=0, column=2 * i, sticky="nsew", padx=(S(12), S(6)))
            self.body.columnconfigure(2 * i, weight=1, uniform="strip")
            text_label(cell, label, font_key="caption", fg=C.text3).pack(anchor="w")
            value = tk.Label(cell, text="—", font=F("metric_sm"), fg=C.text,
                             bg=C.card, anchor="w")
            value.pack(anchor="w", pady=(S(1), 0))
            self._values.append(value)

    def set(self, index: int, text: str, color: Optional[str] = None) -> None:
        if 0 <= index < len(self._values):
            self._values[index].configure(text=text)
            if color:
                self._values[index].configure(fg=color)


# ---------------------------------------------------------------------------
# 选项卡
# ---------------------------------------------------------------------------

class TabView(tk.Frame):
    """胶囊式选项卡，接口与 ttk.Notebook 的常用子集一致。"""

    def __init__(self, master: tk.Misc, *, surface: str = C.bg) -> None:
        super().__init__(master, bg=surface)
        self._surface = surface
        self._items: List[Dict[str, Any]] = []
        self._active = 0
        self._hover = -1
        self._strip_h = S(46)

        self.strip = tk.Canvas(self, height=self._strip_h, bg=surface,
                               highlightthickness=0, bd=0)
        self.strip.pack(fill="x")
        self.host = tk.Frame(self, bg=surface)
        self.host._fluent_surface = surface  # type: ignore[attr-defined]
        self.host.pack(fill="both", expand=True, pady=(S(6), 0))

        self.strip.bind("<Configure>", lambda e: self._draw())
        self.strip.bind("<Button-1>", self._on_click)
        self.strip.bind("<Motion>", self._on_motion)
        self.strip.bind("<Leave>", self._on_leave)
        # 页面切换 / 设置卡片折叠后，画布会在下一帧改变尺寸，这里补一次重绘
        self.bind("<Configure>", lambda e: self.after_idle(self._draw))

    # -- 对外接口 ------------------------------------------------------
    def add(self, page: tk.Misc, text: str) -> None:
        self._items.append({"page": page, "text": text, "rect": (0, 0, 0, 0)})
        if len(self._items) == 1:
            self.select(0)
        self._draw()

    def tabs(self) -> List[tk.Misc]:
        return [item["page"] for item in self._items]

    def index(self, page: tk.Misc) -> int:
        for i, item in enumerate(self._items):
            if item["page"] is page:
                return i
        raise tk.TclError("not a page of this TabView")

    def select(self, target: Any = None) -> Optional[tk.Misc]:
        if target is None:
            return self._items[self._active]["page"] if self._items else None
        index = self.index(target) if isinstance(target, tk.Misc) else int(target)
        index = max(0, min(index, len(self._items) - 1))
        changed = index != self._active
        was_mapped = self._items[index]["page"].winfo_ismapped() if self._items else False
        for i, item in enumerate(self._items):
            page = item["page"]
            if i == index:
                if not page.winfo_ismapped():
                    page.pack(fill="both", expand=True)
            else:
                page.pack_forget()
        self._active = index
        self._draw()
        if changed or not was_mapped:
            self.event_generate("<<NotebookTabChanged>>")
        return self._items[index]["page"]

    # -- 事件 ----------------------------------------------------------
    def _hit(self, x: int, y: int) -> int:
        for i, item in enumerate(self._items):
            x1, y1, x2, y2 = item["rect"]
            if x1 <= x <= x2 and y1 <= y <= y2:
                return i
        return -1

    def _on_click(self, event: tk.Event) -> None:
        index = self._hit(event.x, event.y)
        if index >= 0:
            self.select(index)

    def _on_motion(self, event: tk.Event) -> None:
        index = self._hit(event.x, event.y)
        self.strip.configure(cursor="hand2" if index >= 0 else "")
        if index != self._hover:
            self._hover = index
            self._draw()

    def _on_leave(self, _event: tk.Event) -> None:
        if self._hover != -1:
            self._hover = -1
            self._draw()

    # -- 绘制 ----------------------------------------------------------
    def _draw(self) -> None:
        self.strip.delete("all")
        if not self._items:
            return
        x = 0
        top = S(5)
        height = S(34)
        for i, item in enumerate(self._items):
            text = item["text"]
            font = F("body_strong") if i == self._active else F("body")
            width = font.measure(text) + S(34)
            rect = (x, top, x + width, top + height)
            item["rect"] = rect
            active = i == self._active
            if active:
                round_rect(self.strip, x + 0.5, top + 0.5, x + width - 0.5, top + height - 0.5,
                           S(6), fill=C.card, outline=C.border_strong, width=1)
                # 选中项底部的强调色短线
                round_rect(self.strip, x + S(14), top + height - S(6),
                           x + width - S(14), top + height - S(3), S(2),
                           fill=C.accent, outline="")
                fg = C.accent
            else:
                if i == self._hover:
                    round_rect(self.strip, x + 0.5, top + 0.5, x + width - 0.5,
                               top + height - 0.5, S(6), fill="#E7E7E7", outline="")
                fg = C.text2
            self.strip.create_text(x + width / 2, top + height / 2 + 0.5, text=text,
                                   font=font, fill=fg, anchor="center")
            x += width + S(4)


# ---------------------------------------------------------------------------
# 其它小部件
# ---------------------------------------------------------------------------

class BrandTile(tk.Canvas):
    """标题栏左侧的方形 Logo 块。"""

    def __init__(self, master: tk.Misc, glyph: str = "¥", size: int = 38,
                 surface: Optional[str] = None, fill: str = C.accent) -> None:
        self._surface = surface or surface_of(master)
        self._glyph = glyph
        self._fill = fill
        self._size = S(size)
        super().__init__(master, width=self._size, height=self._size,
                         bg=self._surface, highlightthickness=0, bd=0)
        self.bind("<Configure>", lambda e: self._draw())
        self._draw()

    def _draw(self) -> None:
        w = max(self.winfo_width(), self.winfo_reqwidth())
        h = max(self.winfo_height(), self.winfo_reqheight())
        if w < 4 or h < 4:
            return
        self.delete("all")
        round_rect(self, 0.5, 0.5, w - 0.5, h - 0.5, S(9), fill=self._fill, outline="")
        self.create_text(w / 2, h / 2 + S(1), text=self._glyph,
                         font=F("h1"), fill="#FFFFFF", anchor="center")


class StatusDot(tk.Canvas):
    """状态栏左侧的小圆点。"""

    def __init__(self, master: tk.Misc, color: str = C.success,
                 surface: Optional[str] = None) -> None:
        self._surface = surface or surface_of(master)
        self._color = color
        size = S(8)
        super().__init__(master, width=size, height=size, bg=self._surface,
                         highlightthickness=0, bd=0)
        self._draw()

    def set_color(self, color: str) -> None:
        self._color = color
        self._draw()

    def _draw(self) -> None:
        self.delete("all")
        size = S(8)
        self.create_oval(0, 0, size, size, fill=self._color, outline=self._color)


class FluentScrollbar(tk.Canvas):
    """Fluent 细滚动条：无箭头、悬停变色、可拖动。

    接口与 ttk.Scrollbar 一致：把表格 / 文本的 ``yscrollcommand`` 指到 ``set``，
    并把 ``command`` 传进来即可。
    """

    def __init__(self, master: tk.Misc, command: Optional[Callable[..., Any]] = None, *,
                 surface: Optional[str] = None, width: int = 10,
                 orient: str = "vertical") -> None:
        self._surface = surface or surface_of(master)
        self._command = command
        self._orient = orient
        self._first = 0.0
        self._last = 1.0
        self._hover = False
        self._drag_offset: Optional[float] = None
        self._thumb: Tuple[float, float] = (0.0, 0.0)
        self._size = S(width)
        if orient == "vertical":
            super().__init__(master, width=self._size, bg=self._surface,
                             highlightthickness=0, bd=0, takefocus=0)
        else:
            super().__init__(master, height=self._size, bg=self._surface,
                             highlightthickness=0, bd=0, takefocus=0)
        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Enter>", lambda e: self._set_hover(True))
        self.bind("<Leave>", lambda e: self._set_hover(False))
        self.bind("<Button-1>", self._on_press)
        self.bind("<B1-Motion>", self._on_drag)
        self.bind("<ButtonRelease-1>", self._on_release)

    # -- 对外接口 ------------------------------------------------------
    def set(self, first: Any, last: Any) -> None:
        try:
            self._first = max(0.0, min(1.0, float(first)))
            self._last = max(0.0, min(1.0, float(last)))
        except (TypeError, ValueError):
            self._first, self._last = 0.0, 1.0
        self._draw()

    # -- 事件 ----------------------------------------------------------
    def _set_hover(self, value: bool) -> None:
        self._hover = value
        self._draw()

    def _on_press(self, event: tk.Event) -> None:
        y = event.y if self._orient == "vertical" else event.x
        start, end = self._thumb
        if start <= y <= end:
            self._drag_offset = y - start
        else:
            self._drag_offset = None
            if self._command is not None:
                step = -1 if y < start else 1
                self._command("scroll", step, "pages")
        self._draw()

    def _on_drag(self, event: tk.Event) -> None:
        if self._drag_offset is None or self._command is None:
            return
        length = self.winfo_height() if self._orient == "vertical" else self.winfo_width()
        if length <= 0:
            return
        position = event.y if self._orient == "vertical" else event.x
        fraction = (position - self._drag_offset) / length
        self._command("moveto", max(0.0, min(1.0, fraction)))

    def _on_release(self, _event: tk.Event) -> None:
        self._drag_offset = None
        self._draw()

    # -- 绘制 ----------------------------------------------------------
    def _draw(self) -> None:
        self.delete("all")
        if self._orient == "vertical":
            length = self.winfo_height()
        else:
            length = self.winfo_width()
        if length < 20:
            return
        span = self._last - self._first
        if span >= 0.999:
            self._thumb = (0.0, 0.0)
            return
        minimum = S(28)
        start = self._first * length
        end = self._last * length
        if end - start < minimum:
            end = start + minimum
            if end > length:
                end = length
                start = max(0.0, end - minimum)
        self._thumb = (start, end)
        color = "#9A9A9A" if self._drag_offset is not None else (
            "#ADADAD" if self._hover else "#C9C9C9")
        if self._orient == "vertical":
            round_rect(self, S(2), start + S(2), self._size - S(2), end - S(2), S(3),
                       fill=color, outline=color)
        else:
            round_rect(self, start + S(2), S(2), end - S(2), self._size - S(2), S(3),
                       fill=color, outline=color)


class TextArea(tk.Canvas):
    """圆角多行文本框（内嵌细滚动条），用于粘贴名单之类的场景。"""

    def __init__(self, master: tk.Misc, *, height: int = 220,
                 surface: Optional[str] = None, wrap: str = "word",
                 undo: bool = True, font_key: str = "body") -> None:
        self._surface = surface or surface_of(master)
        super().__init__(master, height=S(height), bg=self._surface,
                         highlightthickness=0, bd=0, takefocus=0)
        self._focus = False

        self.text = tk.Text(self, bd=0, relief="flat", bg=C.card, fg=C.text,
                            insertbackground=C.text, font=F(font_key), wrap=wrap,
                            undo=undo, padx=S(12), pady=S(10),
                            highlightthickness=0, selectbackground=C.accent_tint2,
                            selectforeground=C.text)
        self.scroll = FluentScrollbar(self, command=self.text.yview)
        self.text.configure(yscrollcommand=self.scroll.set)
        self._text_win = self.create_window(S(1), S(1), window=self.text, anchor="nw")
        self._scroll_win = self.create_window(S(1), S(1), window=self.scroll, anchor="nw")
        self.bind("<Configure>", self._on_configure)
        self.text.bind("<FocusIn>", lambda e: self._set_focus(True), add="+")
        self.text.bind("<FocusOut>", lambda e: self._set_focus(False), add="+")
        self.create_rectangle(0, 0, 1, 1, fill=C.card, outline="", tags="bg")

    def _set_focus(self, value: bool) -> None:
        self._focus = value
        self._draw()

    def _on_configure(self, _event: tk.Event) -> None:
        w = self.winfo_width()
        h = self.winfo_height()
        if w < 6 or h < 6:
            return
        bar = S(10)
        self.coords(self._text_win, S(3), S(3))
        self.itemconfigure(self._text_win, width=max(w - bar - S(7), S(20)),
                           height=max(h - S(6), S(20)))
        self.coords(self._scroll_win, w - bar - S(3), S(8))
        self.itemconfigure(self._scroll_win, width=bar, height=max(h - S(16), S(20)))
        self._draw()

    def _draw(self) -> None:
        w = self.winfo_width()
        h = self.winfo_height()
        if w < 6 or h < 6:
            return
        self.delete("bg")
        border = C.accent if self._focus else C.border_input
        # 先铺满描边色，再叠一层内缩 1px 的白色，得到干净的 1px 圆角描边
        round_rect(self, 0, 0, w, h, S(6), fill=border, outline="", tags="bg")
        round_rect(self, 1, 1, w - 1, h - 1, S(5), fill=C.card, outline="", tags="bg")
        self.tag_lower("bg")


__all__ = [
    "C", "S", "F", "init", "enable_dpi_awareness", "style_titlebar",
    "round_rect", "surface_of", "Surface", "divider", "text_label",
    "FluentButton", "FluentCheck", "FluentEntry", "FluentSelect",
    "Card", "TableCard", "StatCard", "StatStrip", "TabView",
    "BrandTile", "StatusDot", "TextArea", "FluentScrollbar",
]
