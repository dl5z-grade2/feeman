# -*- coding: utf-8 -*-
"""饭费管理系统 核心模型层。

负责：
  * .cfee 工程文件（JSON）的读写与数据结构定义
  * 身份证号校验（18 位居民身份证标准，GB 11643）
  * 名单 / 请假日 / 整月不订 的管理
  * 费用核算
  * Excel 导入名单、导出核算结果（openpyxl）

本模块不含任何界面代码，便于单元测试。
"""

from __future__ import annotations

import calendar
import json
import os
import re
from dataclasses import dataclass, field, asdict
from datetime import date
from typing import Dict, Iterable, List, Optional, Tuple

FILE_SUFFIX = ".cfee"
FORMAT_VERSION = 1


# ---------------------------------------------------------------------------
# 身份证校验
# ---------------------------------------------------------------------------

_ID_WEIGHTS = (7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2)
_ID_CHECK = "10X98765432"


def validate_id_card(raw: str) -> Tuple[bool, str]:
    """校验 18 位居民身份证号。

    返回 (是否合法, 原因)。允许为空字符串（表示未提供），此时视为合法。
    校验项：长度 18、前 17 位为数字、出生日期真实存在、末位校验码正确。
    """
    s = (raw or "").strip().upper()
    if not s:
        return True, "未填写"
    if len(s) != 18:
        return False, "身份证号必须为 18 位"
    if not re.fullmatch(r"\d{17}[\dX]", s):
        return False, "身份证号前 17 位须为数字，末位为数字或 X"
    # 出生日期合法性
    try:
        year = int(s[6:10])
        month = int(s[10:12])
        day = int(s[12:14])
        if not (1900 <= year <= 2100):
            return False, "出生年份超出合理范围"
        date(year, month, day)
    except ValueError:
        return False, "身份证号中的出生日期不合法"
    # 校验码
    total = sum(int(s[i]) * _ID_WEIGHTS[i] for i in range(17))
    expect = _ID_CHECK[total % 11]
    if expect != s[17]:
        return False, "身份证校验码错误"
    return True, ""


def id_birth_gender(raw: str) -> Tuple[Optional[date], Optional[str]]:
    """从合法身份证号推导出生日期与性别（第 17 位奇男偶女）。"""
    s = (raw or "").strip().upper()
    ok, _ = validate_id_card(s)
    if not s or not ok:
        return None, None
    try:
        b = date(int(s[6:10]), int(s[10:12]), int(s[12:14]))
    except ValueError:
        b = None
    g = "男" if int(s[16]) % 2 == 1 else "女"
    return b, g


# ---------------------------------------------------------------------------
# 排序
# ---------------------------------------------------------------------------

def sid_sort_key(sid: str):
    """学号的"自然排序"键。

    把学号切成数字段与非数字段，数字段按数值大小比较，因此
    ``1 < 2 < 10 < 11``，而不是字典序的 ``1 < 10 < 11 < 2``；
    非数字段（如 "A"、"2024-"）按小写文本比较，可混合使用。
    """
    s = str(sid)
    parts = []
    for chunk in re.split(r"(\d+)", s):
        if chunk.isdigit():
            parts.append((0, int(chunk), ""))
        else:
            parts.append((1, 0, chunk.lower()))
    return (parts, s)


# ---------------------------------------------------------------------------
# 数据模型
# ---------------------------------------------------------------------------

@dataclass
class Person:
    sid: str                       # 学号
    name: str                      # 姓名
    id_card: str = ""              # 身份证（可选，已校验）
    gender: str = ""               # 性别（可选：男/女/""）


@dataclass
class Project:
    """一个 .cfee 工程的全部数据。"""

    title: str = "未命名工程"
    class_name: str = ""           # 班级（导出核算表的“班级”列；空则用 title）
    year: int = 2000
    month: int = 1
    # 全月统一单价；>0 生效。0 表示未设置。
    month_price: float = 0.0
    # 每周固定休息（不供餐）的星期，0=周一 … 6=周日；这些天单价恒为 0
    skip_weekdays: set = field(default_factory=lambda: {5, 6})
    # 每天单价覆盖：{"1": 12.5, ...} 仅当 month_price<=0 或明确覆盖时使用
    day_prices: Dict[str, float] = field(default_factory=dict)
    # 名单：学号 -> Person
    roster: Dict[str, Person] = field(default_factory=dict)
    # 请假/不订记录：学号 -> {"1","5",...} 当月不订的日
    # 没有“整月不订”这个独立标记：整月不订就是把当月每一天都写进这里
    skips: Dict[str, set] = field(default_factory=dict)
    # 导出核算表时是否剔除“费用为 0”的人员（随工程持久化）
    export_skip_zero: bool = True

    # ---- 元信息 ----
    @property
    def days_in_month(self) -> int:
        return calendar.monthrange(self.year, self.month)[1]

    def day_list(self) -> List[int]:
        return list(range(1, self.days_in_month + 1))

    def weekday_of(self, day: int) -> int:
        """某天是星期几：0=周一 … 6=周日。"""
        return calendar.weekday(self.year, self.month, int(day))

    def is_closed_day(self, day: int) -> bool:
        """该天是否属于设置的每周休息日（跳过、不供餐）。"""
        return self.weekday_of(day) in self.skip_weekdays

    def open_days(self) -> List[int]:
        """当月实际供餐的日期（排除每周休息日）。"""
        return [d for d in self.day_list() if not self.is_closed_day(d)]

    def closed_days(self) -> List[int]:
        """当月因每周休息日而停餐的日期。"""
        return [d for d in self.day_list() if self.is_closed_day(d)]

    def price_for_day(self, day: int) -> float:
        """取某天的单价：休息日为 0；否则优先日覆盖，其次全月单价。"""
        if self.is_closed_day(day):
            return 0.0
        v = self.day_prices.get(str(day))
        if v is not None:
            return float(v)
        return float(self.month_price)


    # ---- 名单 ----
    def add_person(self, p: Person) -> None:
        self.roster[p.sid] = p

    def remove_person(self, sid: str) -> None:
        self.roster.pop(sid, None)
        self.skips.pop(sid, None)

    # ---- 不订记录 ----
    def set_skip(self, sid: str, days) -> None:
        self.skips[sid] = set(int(d) for d in days)

    def is_skipping(self, sid: str, day: int) -> bool:
        return int(day) in self.skips.get(sid, set())

    def set_opt_out_month(self, sid: str, value: bool) -> None:
        """整月不订 / 取消整月不订。

        不使用独立标记，而是把当月**每一个供餐日**都写入（或移出）不订记录，
        因此与逐日不订完全等价、可被逐日修改覆盖。
        """
        open_days = set(self.open_days())
        if value:
            self.skips[sid] = set(self.skips.get(sid, set())) | open_days
        else:
            rest = set(self.skips.get(sid, set())) - open_days
            if rest:
                self.skips[sid] = rest
            else:
                self.skips.pop(sid, None)

    def is_opt_out(self, sid: str) -> bool:
        """是否整月不订：当月每一个**供餐日**都在不订记录里。

        只看供餐日，因此每周休息日本来就不供餐，不参与判断。
        """
        open_days = set(self.open_days())
        return bool(open_days) and open_days <= set(self.skips.get(sid, set()))

    # ---- 某天就餐人数 ----
    def eats_on(self, sid: str, day: int) -> bool:
        """某人某天是否订餐。

        该天为每周休息日 -> False；
        该天在其不订记录里 -> False；否则 True。
        """
        day = int(day)
        if day not in self.day_list():
            return False
        if self.is_closed_day(day):
            return False
        return day not in self.skips.get(sid, set())

    def diners_on(self, day: int) -> List[Dict]:
        """某天订餐的人员列表，按学号自然顺序。"""
        return [
            {"sid": sid, "name": self.roster[sid].name, "gender": self.roster[sid].gender}
            for sid in self.sorted_sids()
            if self.eats_on(sid, day)
        ]

    def diner_count(self, day: int) -> int:
        """某天订餐人数。"""
        return len(self.diners_on(day))

    def non_diners_on(self, day: int) -> List[Dict]:
        """某天不订餐的人员列表，按学号自然顺序。"""
        return [
            {"sid": sid, "name": self.roster[sid].name, "gender": self.roster[sid].gender}
            for sid in self.sorted_sids()
            if not self.eats_on(sid, day)
        ]

    # ---- 核算 ----
    def ordered_days(self, sid: str) -> int:
        """某人当月实际订餐天数（不含每周休息日）。"""
        skipped = self.skips.get(sid, set())
        return sum(1 for d in self.open_days() if d not in skipped)

    def person_cost(self, sid: str) -> float:
        """某人当月应缴费用 = 其实际订餐各天单价之和（休息日不计费）。"""
        skipped = self.skips.get(sid, set())
        return sum(self.price_for_day(d) for d in self.open_days() if d not in skipped)

    def sorted_sids(self) -> List[str]:
        """按学号自然顺序（1, 2, 10, 11）返回名单中的学号。"""
        return sorted(self.roster.keys(), key=sid_sort_key)

    def settle(self) -> List[Dict]:
        """核算所有人员，返回按学号自然顺序排序的结果列表。"""
        rows = []
        for sid in self.sorted_sids():
            p = self.roster[sid]
            rows.append({
                "sid": p.sid,
                "name": p.name,
                "gender": p.gender,
                "id_card": p.id_card,
                "days": self.ordered_days(sid),
                "opt_out": self.is_opt_out(sid),
                "cost": round(self.person_cost(sid), 2),
            })
        return rows

    def total_cost(self) -> float:
        return round(sum(r["cost"] for r in self.settle()), 2)

    # ---- 序列化 ----
    def to_dict(self) -> Dict:
        return {
            "format": "cfee",
            "version": FORMAT_VERSION,
            "title": self.title,
            "class_name": self.class_name,
            "year": self.year,
            "month": self.month,
            "month_price": self.month_price,
            "skip_weekdays": sorted(self.skip_weekdays),
            "export_skip_zero": bool(self.export_skip_zero),
            "day_prices": {str(k): v for k, v in self.day_prices.items()},
            "roster": [asdict(p) for p in self.roster.values()],
            "skips": {sid: sorted(list(days)) for sid, days in self.skips.items()},
        }

    @classmethod
    def from_dict(cls, data: Dict) -> "Project":
        proj = cls(
            title=data.get("title", "未命名工程"),
            class_name=str(data.get("class_name", "") or ""),
            year=int(data.get("year", 2000)),
            month=int(data.get("month", 1)),
            month_price=float(data.get("month_price", 0.0)),
            export_skip_zero=bool(data.get("export_skip_zero", True)),
        )
        # 旧版工程没有 skip_weekdays 字段：保留默认（周六、周日）
        if "skip_weekdays" in data:
            proj.skip_weekdays = {int(w) for w in data.get("skip_weekdays") or []}
            proj.skip_weekdays = {w for w in proj.skip_weekdays if 0 <= w <= 6}
        proj.day_prices = {str(k): float(v) for k, v in data.get("day_prices", {}).items()}
        for item in data.get("roster", []):
            p = Person(
                sid=str(item.get("sid", "")),
                name=str(item.get("name", "")),
                id_card=str(item.get("id_card", "") or ""),
                gender=str(item.get("gender", "") or ""),
            )
            if p.sid:
                proj.roster[p.sid] = p
        for sid, days in data.get("skips", {}).items():
            proj.skips[str(sid)] = set(int(d) for d in days)
        # 兼容旧工程：过去的“整月不订”独立标记转换为“每个供餐日都不订”
        for sid in data.get("month_opt_out", []) or []:
            sid = str(sid)
            proj.skips[sid] = set(proj.skips.get(sid, set())) | set(proj.open_days())
        return proj


# ---------------------------------------------------------------------------
# 文件读写
# ---------------------------------------------------------------------------

def save_project(proj: Project, path: str) -> None:
    if not path.lower().endswith(FILE_SUFFIX):
        path += FILE_SUFFIX
    with open(path, "w", encoding="utf-8") as f:
        json.dump(proj.to_dict(), f, ensure_ascii=False, indent=2)


def load_project(path: str) -> Project:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if data.get("format") != "cfee":
        raise ValueError("不是有效的 .cfee 工程文件")
    return Project.from_dict(data)


# ---------------------------------------------------------------------------
# 从自由文本解析“本月不订”
# ---------------------------------------------------------------------------

# 单餐关键词 -> 归类
_MEAL_LUNCH = "午餐"
_MEAL_DINNER = "晚餐"
_MEAL_SATURDAY = "周六餐"
_ALL_MEALS = (_MEAL_LUNCH, _MEAL_DINNER, _MEAL_SATURDAY)

# 出现这些词表示“三餐全不订”
_ALL_MEAL_WORDS = ("所有", "全部", "全都", "三餐", "一日三餐", "全天", "都")
# 出现这些词表示“只订某餐”（取反：未提到的餐才是不订的）
_ONLY_WORDS = ("只", "仅")


def _detect_meals(segment: str) -> Tuple[set, bool, bool]:
    """解析一段描述里提到的餐别。

    返回 (提到的餐别集合, 是否出现“所有/都”, 是否出现“只/仅”)。
    餐别集合的取值为 {"午餐", "晚餐", "周六餐"}。
    """
    text = segment.replace(" ", "")
    meals = set()
    if re.search(r"午餐|午饭|中餐|(?<!下)午(?!餐|间)", text):
        meals.add(_MEAL_LUNCH)
    if re.search(r"晚餐|晚饭|晚(?!上|间)", text):
        meals.add(_MEAL_DINNER)
    if re.search(r"周六|星期六|周末餐|假日餐", text):
        meals.add(_MEAL_SATURDAY)
    has_all = any(w in text for w in _ALL_MEAL_WORDS)
    has_only = any(w in text for w in _ONLY_WORDS)
    return meals, has_all, has_only


def _match_person(proj: "Project", body: str) -> Optional[str]:
    """在一行文本里找出对应的人员学号。

    优先用**人名**穷举匹配（最长优先，避免“张三”抢走“张三丰”），
    因为行首的“7号”“44号”未必是学号（可能是班级序号）；
    人名匹配不到时，才用“N号”尝试按学号匹配。
    """
    compact = body.replace(" ", "")

    # 1) 人名：最长优先
    names = sorted(proj.roster.values(), key=lambda p: len(p.name), reverse=True)
    for p in names:
        if p.name and p.name in compact:
            return p.sid

    # 2) 学号：44号 / 44 号
    for m in re.finditer(r"(\d{1,10})\s*号", body):
        num = m.group(1)
        if num in proj.roster:
            return num
        stripped = num.lstrip("0") or "0"
        for sid in proj.roster:
            if sid.lstrip("0") == stripped:
                return sid
    return None


def _meal_days(proj: "Project") -> Dict[str, List[int]]:
    """把餐别映射到当月具体日期。

    午餐 / 晚餐 -> 周一到周五；周六餐 -> 周六。

    这里按**当月日历**映射，不排除每周休息日：文本里写明“不订周六餐”
    就说明该月周六是供餐日，若因默认休息日是周六而把它整条丢弃，
    用户会觉得功能失效。记到休息日上不会影响费用（休息日单价恒为 0，
    也不计入订餐天数），但能在“每日不订”里看到真实记录。
    """
    weekdays, saturdays = [], []
    for d in proj.day_list():
        wd = proj.weekday_of(d)
        if wd < 5:
            weekdays.append(d)
        elif wd == 5:
            saturdays.append(d)
    return {
        _MEAL_LUNCH: list(weekdays),
        _MEAL_DINNER: list(weekdays),
        _MEAL_SATURDAY: list(saturdays),
    }


def parse_unordered_text(proj: "Project", text: str,
                         allowed_meals: Optional[Iterable[str]] = None,
                         default_unordered: Optional[Iterable[str]] = None
                         ) -> Tuple[List[Dict], List[str]]:
    """从自由格式文本中解析“本月不订”的人员。

    文本形如::

        1. 苏煜博，10月不订午餐
        2. 44号范付璐10月不订午餐、晚餐、周六餐
        4. 41号陈心怡不订所有餐
        7. 43号赵思琪10月不订晚餐
        8. 7号张若衡妈妈 10月不订晚餐，周六餐

    规则：
      * 逐行处理：先用人名穷举匹配（最长优先），匹配不到再用“N号”按学号匹配。
      * 关键词：午餐 / 晚餐 / 周六餐；“所有”“都”“三餐”表示全不订；
        “只订午餐”这类“只/仅”表述取反 —— 只有提到的餐是订的，其余不订。
      * 午餐 / 晚餐对应周一到周五，周六餐对应周六；
        某餐在本月没有对应日期（如本月没有周六）时静默忽略。
      * 未匹配到人员、或未识别到餐别的行，记入警告。

    参数 ``allowed_meals``：**只保留这些餐别的不订记录**。文本里提到但不在其中
    的餐别会被丢弃（视为照常订餐），也不会因为“所有/都”而被补上。
    传 None 表示不过滤（文本提到什么就记什么）。
    参数 ``default_unordered``：文本没写任何餐别时，按这些餐别不订
    （同样受 ``allowed_meals`` 约束）。

    返回 (结果列表, 警告列表)；结果项为
    ``{"sid","name","meals","days","note"}``，``days`` 为要标记不订的日期。
    """
    results: List[Dict] = []
    warnings: List[str] = []
    groups = _meal_days(proj)
    all_open = set(proj.open_days())

    # 只保留这些餐别的不订记录；None 表示不过滤
    allowed = None if allowed_meals is None else {m for m in allowed_meals if m in _ALL_MEALS}
    # 行内没写餐别时的默认不订餐别：未显式指定时取 allowed（没勾就不推断）
    if default_unordered is None:
        default_off = [] if allowed is None else [m for m in _ALL_MEALS if m in allowed]
    else:
        default_off = [m for m in default_unordered if m in _ALL_MEALS]

    def _keep_only_allowed(meals_seq):
        if allowed is None:
            return [m for m in _ALL_MEALS if m in meals_seq and groups.get(m)]
        return [m for m in _ALL_MEALS if m in meals_seq and m in allowed and groups.get(m)]

    for lineno, raw in enumerate(text.splitlines(), start=1):
        line = raw.strip()
        if not line:
            continue
        # 去掉行首序号：1. / 1、/ 1）/ 1:
        body = re.sub(r"^\s*\d{1,3}\s*[.．、,，:：)）]\s*", "", line).strip()
        if not body:
            continue

        sid = _match_person(proj, body)
        if sid is None:
            warnings.append(f"第 {lineno} 行：名单中找不到对应人员，已跳过（{line}）")
            continue
        person = proj.roster[sid]

        meals, has_all, has_only = _detect_meals(body)
        if not meals and not has_all:
            if not default_off:
                warnings.append(f"第 {lineno} 行：{person.name} 未识别到“午餐/晚餐/周六餐”，已跳过")
                continue
            # 整行没写餐别，按默认不订的餐别处理
            meals, has_all, has_only = set(default_off), False, False

        # 先算出“文本认为不订的餐别”，再按 allowed_meals 过滤
        if has_only:
            # 只订提到的餐，其余不订
            unordered = [m for m in _ALL_MEALS if m not in meals]
        elif has_all:
            unordered = [m for m in _ALL_MEALS]
        else:
            unordered = [m for m in _ALL_MEALS if m in meals]

        unordered = _keep_only_allowed(unordered)

        days = set()
        for meal in unordered:
            days |= set(groups.get(meal, []))

        if not days:
            # 提到的餐别都被过滤掉，或本月没有对应日期 -> 静默忽略
            continue

        results.append({
            "sid": sid,
            "name": person.name,
            "meals": unordered,
            "days": sorted(days),
            "note": "、".join(unordered) if unordered else "",
        })

    # 合并同一人的多行
    combined: Dict[str, Dict] = {}
    for item in results:
        cur = combined.get(item["sid"])
        if cur is None:
            combined[item["sid"]] = dict(item)
        else:
            cur["days"] = sorted(set(cur["days"]) | set(item["days"]))
            meals = set(cur["meals"]) | set(item["meals"])
            cur["meals"] = [m for m in _ALL_MEALS if m in meals]
            cur["note"] = "、".join(cur["meals"])
    return list(combined.values()), warnings


def apply_unordered_text(proj: "Project", text: str,
                         allowed_meals: Optional[Iterable[str]] = None,
                         default_unordered: Optional[Iterable[str]] = None,
                         replace: bool = False) -> Tuple[int, List[str]]:
    """把解析结果写入工程的 skips。

    ``replace=True`` 时先清空这些人的原有逐日记录再写入；
    默认 False 表示与原有记录合并。

    返回 (实际应用的人数, 警告列表)。
    """
    results, warnings = parse_unordered_text(
        proj, text, allowed_meals=allowed_meals, default_unordered=default_unordered)
    for item in results:
        sid = item["sid"]
        base = set() if replace else set(proj.skips.get(sid, set()))
        merged = base | set(item["days"])
        if merged:
            proj.set_skip(sid, merged)
        else:
            proj.skips.pop(sid, None)
    return len(results), warnings


# 名单表可识别的表头别名（小写匹配）
_HEADER_ALIASES = {
    "sid": {"学号", "编号", "id", "studentid", "student_id", "sid"},
    "name": {"姓名", "名字", "name"},
    "id_card": {"身份证", "身份证号", "身份证号码", "idcard", "id_card", "证件号"},
    "gender": {"性别", "gender", "sex"},
}


def _norm_header(h) -> str:
    return str(h).strip().lower().replace(" ", "")


def import_roster_from_excel(path: str) -> Tuple[List[Person], List[str]]:
    """从 Excel(.xlsx) 或 CSV 导入名单。

    识别表头：学号/姓名/身份证(可选)/性别(可选)。
    返回 (人员列表, 警告列表)。身份证非法者会被跳过并记入警告。
    """
    import openpyxl

    warnings: List[str] = []
    persons: List[Person] = []

    if path.lower().endswith(".csv"):
        import csv
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            rows = list(csv.reader(f))
    else:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()

    if not rows:
        raise ValueError("表格为空")

    header = [_norm_header(c) for c in rows[0]]
    col_map: Dict[str, int] = {}
    for key, aliases in _HEADER_ALIASES.items():
        for idx, h in enumerate(header):
            if h in aliases or h in {a.replace("_", "") for a in aliases}:
                col_map[key] = idx
                break

    if "sid" not in col_map or "name" not in col_map:
        raise ValueError("表头必须包含 学号 和 姓名 两列（同时支持 身份证/性别 可选列）")

    seen = set()
    for r_i, row in enumerate(rows[1:], start=2):
        def cell(key):
            if key not in col_map:
                return ""
            idx = col_map[key]
            if idx < len(row) and row[idx] is not None:
                return str(row[idx]).strip()
            return ""

        sid = cell("sid")
        name = cell("name")
        idc = cell("id_card")
        gender = cell("gender")
        if not sid and not name:
            continue
        if not sid or not name:
            warnings.append(f"第 {r_i} 行：学号或姓名为空，已跳过")
            continue
        if sid in seen:
            warnings.append(f"第 {r_i} 行：学号 {sid} 重复，已跳过")
            continue
        ok, reason = validate_id_card(idc)
        if not ok:
            warnings.append(f"第 {r_i} 行：{name} 身份证无效（{reason}），已跳过")
            continue
        # 性别缺失时可由身份证推导
        if not gender and idc:
            _, g = id_birth_gender(idc)
            gender = g or ""
        if gender and gender not in ("男", "女"):
            warnings.append(f"第 {r_i} 行：{name} 性别“{gender}”非男/女，已置空")
            gender = ""
        seen.add(sid)
        persons.append(Person(sid=sid, name=name, id_card=idc.upper() if idc else "", gender=gender))

    return persons, warnings


def export_settlement(proj: Project, path: str, skip_zero: Optional[bool] = None) -> None:
    """导出核算结果到 .xlsx。

    列结构：班级 | 姓名 | 身份证号 | 费用；
    首行为表头，末行为费用合计（=SUM 公式），无额外标题、边框或配色。

    skip_zero 为 None 时按工程自身的 ``export_skip_zero`` 设置决定是否
    剔除费用为 0 的人员；显式传入 True/False 可临时覆盖。
    """
    import openpyxl
    from openpyxl.styles import Font

    if skip_zero is None:
        skip_zero = bool(getattr(proj, "export_skip_zero", True))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "饭费核算"

    headers = ["班级", "姓名", "身份证号", "费用"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=1, column=c, value=h)

    class_name = (proj.class_name or proj.title or "").strip()
    rows = proj.settle()
    if skip_zero:
        rows = [r for r in rows if float(r["cost"]) != 0.0]
    r = 2
    for item in rows:
        ws.cell(row=r, column=1, value=class_name)
        ws.cell(row=r, column=2, value=item["name"])
        ws.cell(row=r, column=3, value=item["id_card"])
        ws.cell(row=r, column=4, value=item["cost"])
        r += 1

    # 合计行：与参考文件一样只写 SUM 公式，放在费用列
    if rows:
        ws.cell(row=r, column=4, value=f"=SUM(D2:D{r - 1})")

    # 统一字体：等线 11，无加粗、无边框、无填充
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=4):
        for cell in row:
            cell.font = Font(name="等线", size=11)

    # 身份证号列较宽，其余按内容自适应
    ws.column_dimensions["C"].width = 20.25
    for col, width in (("A", 12), ("B", 10), ("D", 10)):
        ws.column_dimensions[col].width = width

    wb.save(path)



def export_roster_template(path: str) -> None:
    """导出一个名单模板 xlsx，方便填写后导入。"""
    import openpyxl
    from openpyxl.styles import Font
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "名单"
    ws.append(["学号", "姓名", "身份证", "性别"])
    for c in ws[1]:
        c.font = Font(bold=True)
    ws.append(["20240001", "张三", "110101199003071233", "男"])
    ws.append(["20240002", "李四", "", "女"])
    for i, w in enumerate([14, 12, 22, 8], start=1):
        ws.column_dimensions[chr(64 + i)].width = w
    wb.save(path)
