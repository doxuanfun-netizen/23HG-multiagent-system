# -*- coding: utf-8 -*-
"""Bộ tính CPM chuẩn cho 23HG (ngày công).

Quy ước (khớp công thức Excel WORKDAY.INTL):
  * Lịch: làm việc Thứ Hai–Thứ Bảy, nghỉ Chủ nhật và các ngày lễ.
  * Thời lượng tính bằng ngày công; EF = ES + (dur-1) ngày công. Mốc (dur=0): EF = ES.
  * Quan hệ: FS, SS, FF, SF, trễ (lag) tính bằng ngày công, có thể âm.
  * TF (dự trữ tổng) tính bằng ngày công giữa ES và LS.
Không phụ thuộc thư viện ngoài.
"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from typing import Dict, Iterable, List, Optional, Tuple

DAY = dt.timedelta(days=1)


class CpmError(ValueError):
    """Dữ liệu mạng không hợp lệ (tiền nhiệm thiếu, vòng lặp...)."""


class Calendar:
    def __init__(self, holidays: Iterable[dt.date] = (), off_weekdays: Iterable[int] = (6,)):
        self.holidays = frozenset(holidays)
        self.off = frozenset(off_weekdays)  # 0=Thứ Hai ... 6=Chủ nhật

    def is_workday(self, d: dt.date) -> bool:
        return d.weekday() not in self.off and d not in self.holidays

    def on_or_after(self, d: dt.date) -> dt.date:
        while not self.is_workday(d):
            d += DAY
        return d

    def add(self, d: dt.date, n: int) -> dt.date:
        """Tương đương WORKDAY.INTL(d, n): n=0 trả về chính d."""
        step = DAY if n >= 0 else -DAY
        k = abs(n)
        while k > 0:
            d += step
            if self.is_workday(d):
                k -= 1
        return d

    def count(self, a: dt.date, b: dt.date) -> int:
        """Số ngày công trong [a, b] (a <= b)."""
        n, d = 0, a
        while d <= b:
            n += self.is_workday(d)
            d += DAY
        return n

    def float_days(self, es: dt.date, ls: dt.date) -> int:
        if es <= ls:
            return self.count(es, ls) - 1
        return -(self.count(ls, es) - 1)


@dataclass
class Task:
    wbs: str
    duration: int = 0
    kind: str = "Task"  # Task | Milestone | Summary
    preds: List[Tuple[str, str, int]] = field(default_factory=list)  # (wbs, rel, lag)
    snet: Optional[dt.date] = None  # ràng buộc bắt đầu sớm nhất
    name: str = ""
    rain: float = 1.0  # hệ số năng suất mùa mưa K_tt (0<K<1 kéo dài thời lượng)

    def effective_duration(self, use_rain: bool = False) -> int:
        """Thời lượng thực (ngày công). Giống VBA: Round(dur / K) khi 0<K<1 (làm tròn về số chẵn)."""
        if use_rain and 0 < self.rain < 1 and self.duration > 0:
            return int(round(self.duration / self.rain))
        return self.duration


@dataclass
class Result:
    es: dt.date
    ef: dt.date
    ls: dt.date
    lf: dt.date
    tf: int
    critical: bool


def _order(tasks: Dict[str, Task]) -> List[str]:
    state: Dict[str, int] = {}
    out: List[str] = []

    def visit(w: str, trail: Tuple[str, ...]):
        if state.get(w) == 2:
            return
        if state.get(w) == 1:
            raise CpmError("Vòng lặp tiền nhiệm: " + " -> ".join(trail + (w,)))
        state[w] = 1
        for pw, _, _ in tasks[w].preds:
            if pw not in tasks:
                raise CpmError(f"Công tác {w}: tiền nhiệm '{pw}' không tồn tại")
            visit(pw, trail + (w,))
        state[w] = 2
        out.append(w)

    for w in tasks:
        visit(w, ())
    return out


def schedule(tasks_in: Iterable[Task], project_start: dt.date, cal: Calendar,
             deadline: Optional[dt.date] = None, use_rain: bool = False) -> Dict[str, Result]:
    tasks = {t.wbs: t for t in tasks_in}
    if len(tasks) == 0:
        return {}
    leaves = [w for w, t in tasks.items() if t.kind != "Summary"]
    for w in leaves:
        for pw, rel, _ in tasks[w].preds:
            if rel not in ("FS", "SS", "FF", "SF"):
                raise CpmError(f"Công tác {w}: quan hệ '{rel}' không hợp lệ")
            if pw not in tasks:
                raise CpmError(f"Công tác {w}: tiền nhiệm '{pw}' không tồn tại")
            if tasks[pw].kind == "Summary":
                raise CpmError(f"Công tác {w}: không được nối tiền nhiệm vào dòng Summary '{pw}'")
    order = [w for w in _order(tasks) if tasks[w].kind != "Summary"]
    start0 = cal.on_or_after(project_start)

    # ---- tính xuôi ----
    es: Dict[str, dt.date] = {}
    ef: Dict[str, dt.date] = {}
    for w in order:
        t = tasks[w]
        d = max(0, t.effective_duration(use_rain) - 1) if t.kind != "Milestone" else 0
        cands = [start0]
        if t.snet:
            cands.append(cal.on_or_after(t.snet))
        ef_min: Optional[dt.date] = None  # ràng buộc từ FF/SF (xác định EF tối thiểu)
        for pw, rel, lag in t.preds:
            if rel == "FS":
                cands.append(cal.add(ef[pw], 1 + lag))
            elif rel == "SS":
                cands.append(cal.add(es[pw], lag))
            elif rel == "FF":
                e = cal.add(ef[pw], lag)
                ef_min = e if ef_min is None else max(ef_min, e)
            else:  # SF
                e = cal.add(es[pw], lag)
                ef_min = e if ef_min is None else max(ef_min, e)
        es_w = max(cands)
        if ef_min is not None:
            es_w = max(es_w, cal.add(ef_min, -d))
        es[w] = es_w
        ef[w] = es_w if d == 0 else cal.add(es_w, d)

    end = max(ef[w] for w in order)
    if deadline is not None:
        end = deadline  # hạn chót cố định: TF có thể âm nếu lịch vượt hạn

    # ---- tính ngược ----
    succ: Dict[str, List[Tuple[str, str, int]]] = {w: [] for w in order}
    for w in order:
        for pw, rel, lag in tasks[w].preds:
            succ[pw].append((w, rel, lag))
    ls: Dict[str, dt.date] = {}
    lf: Dict[str, dt.date] = {}
    for w in reversed(order):
        t = tasks[w]
        d = max(0, t.effective_duration(use_rain) - 1) if t.kind != "Milestone" else 0
        lf_c: List[dt.date] = [end]
        ls_c: List[dt.date] = []
        for s, rel, lag in succ[w]:
            if rel == "FS":
                lf_c.append(cal.add(ls[s], -(1 + lag)))
            elif rel == "SS":
                ls_c.append(cal.add(ls[s], -lag))
            elif rel == "FF":
                lf_c.append(cal.add(lf[s], -lag))
            else:  # SF
                ls_c.append(cal.add(lf[s], -lag))
        lf_w = min(lf_c)
        ls_w = lf_w if d == 0 else cal.add(lf_w, -d)
        if ls_c:
            ls_w = min(ls_w, min(ls_c))
        ls[w] = ls_w
        lf[w] = ls_w if d == 0 else cal.add(ls_w, d)

    res: Dict[str, Result] = {}
    for w in order:
        tf = cal.float_days(es[w], ls[w])
        res[w] = Result(es[w], ef[w], ls[w], lf[w], tf, tf <= 0)

    # ---- tổng hợp Summary ----
    for w, t in tasks.items():
        if t.kind != "Summary":
            continue
        kids = [k for k in order if k.startswith(w + ".")]
        if not kids:
            raise CpmError(f"Summary '{w}' không có công tác con")
        res[w] = Result(
            min(res[k].es for k in kids), max(res[k].ef for k in kids),
            min(res[k].ls for k in kids), max(res[k].lf for k in kids),
            min(res[k].tf for k in kids), any(res[k].critical for k in kids))
    return res


def parse_pred(text: Optional[str]) -> List[Tuple[str, str, int]]:
    """'1.2.1.2SS+20; 1.1.1FS' -> [('1.2.1.2','SS',20), ('1.1.1','FS',0)]"""
    import re
    out: List[Tuple[str, str, int]] = []
    if not text:
        return out
    for part in str(text).replace(";", ",").split(","):
        part = part.strip()
        if not part:
            continue
        m = re.fullmatch(r"([\d\.]*\d)(FS|SS|FF|SF)?([+-]\d+)?", part)
        if not m:
            raise CpmError(f"Không đọc được tiền nhiệm: '{part}'")
        out.append((m.group(1), m.group(2) or "FS", int(m.group(3) or 0)))
    return out
