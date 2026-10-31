"""予定表をPDFにする。1日1枚と、全期間を1枚にまとめた出力がある。"""

from __future__ import annotations

import os
import re
from datetime import datetime, timedelta

from model import effort_spans, slot_count_for, task_span

FONT_CANDIDATES = (
    r"C:\Windows\Fonts\YuGothR.ttc",
    r"C:\Windows\Fonts\YuGothM.ttc",
    r"C:\Windows\Fonts\meiryo.ttc",
    r"C:\Windows\Fonts\msgothic.ttc",
)
FONT_NAME = "ScheduleJP"


def pdf_file_name(file_name: str) -> str:
    base = re.sub(r"\.(xlsx|xlsm|xls|pdf)$", "", str(file_name or "スケジュール"), flags=re.IGNORECASE)
    return f"{base}_予定.pdf"


def pdf_overview_file_name(file_name: str) -> str:
    base = re.sub(r"\.(xlsx|xlsm|xls|pdf)$", "", str(file_name or "スケジュール"), flags=re.IGNORECASE)
    return f"{base}_全体.pdf"


OVERVIEW_MARGIN = 28
OVERVIEW_GUTTER = 46
OVERVIEW_HEADER = 34
OVERVIEW_AXIS_H = 28
OVERVIEW_EFFORT_H = 18
OVERVIEW_SLOT_H = 44
OVERVIEW_PT_PER_HOUR = 48
MAX_PAGE_PT = 14000


def slot_caption(slot: int, names) -> str:
    if names and slot < len(names) and str(names[slot]).strip():
        return str(names[slot]).strip()
    return str(slot + 1)


def fitted_gutter(names, slot_count: int, minimum: float, font_size: float) -> float:
    from reportlab.pdfbase import pdfmetrics

    width = float(minimum)
    for slot in range(slot_count):
        text = slot_caption(slot, names)
        width = max(width, pdfmetrics.stringWidth(text, FONT_NAME, font_size) + 12)
    return min(width, 160)


def overview_geometry(axis_start, axis_end, slot_count: int, gutter: float = OVERVIEW_GUTTER):
    hours = max((axis_end - axis_start).total_seconds() / 3600, 1 / 60)
    chart_w = hours * OVERVIEW_PT_PER_HOUR
    chart_h = OVERVIEW_HEADER + OVERVIEW_AXIS_H + OVERVIEW_EFFORT_H + slot_count * OVERVIEW_SLOT_H
    natural_w = OVERVIEW_MARGIN * 2 + gutter + chart_w
    natural_h = OVERVIEW_MARGIN * 2 + chart_h
    scale = min(1.0, MAX_PAGE_PT / natural_w, MAX_PAGE_PT / natural_h)
    return natural_w, natural_h, scale


def write_schedule_pdf(path, tasks, axis_start, axis_end, alert_at, file_name="", sheet_name="", slot_count=None, slot_names=None) -> None:
    if not tasks or axis_start is None or axis_end is None:
        raise ValueError("書き出す予定がありません。")
    try:
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.pdfgen import canvas
    except ImportError as exc:
        raise RuntimeError("PDFには reportlab が必要です。次を実行してください: py -3 -m pip install reportlab") from exc
    register_font()
    from app import COLORS, effort_color, format_day, format_hm, style_offsets

    document = canvas.Canvas(str(path), pagesize=landscape(A4))
    page_w, page_h = landscape(A4)
    days = list(iter_days(axis_start, axis_end))
    for day in days:
        draw_day(
            document,
            page_w,
            page_h,
            tasks,
            day,
            alert_at,
            file_name,
            sheet_name,
            COLORS,
            effort_color,
            format_day,
            format_hm,
            style_offsets,
            slot_count if slot_count else slot_count_for(tasks),
            slot_names,
        )
        document.showPage()
    document.save()


def write_overview_pdf(path, tasks, axis_start, axis_end, alert_at, file_name="", sheet_name="", slot_count=None, slot_names=None) -> None:
    if not tasks or axis_start is None or axis_end is None or axis_end <= axis_start:
        raise ValueError("書き出す予定がありません。")
    try:
        from reportlab.pdfgen import canvas
    except ImportError as exc:
        raise RuntimeError("PDFには reportlab が必要です。次を実行してください: py -3 -m pip install reportlab") from exc
    register_font()
    from app import COLORS, effort_color, format_day, format_hm, style_offsets

    count = slot_count if slot_count else slot_count_for(tasks)
    gutter = fitted_gutter(slot_names, count, OVERVIEW_GUTTER, 9)
    natural_w, natural_h, scale = overview_geometry(axis_start, axis_end, count, gutter)
    document = canvas.Canvas(str(path), pagesize=(natural_w * scale, natural_h * scale))
    try:
        document.setViewerPreference("FitWindow", "true")
    except Exception:
        pass
    if scale < 1:
        document.scale(scale, scale)
    draw_overview(
        document,
        natural_w,
        natural_h,
        tasks,
        axis_start,
        axis_end,
        alert_at,
        file_name,
        sheet_name,
        COLORS,
        effort_color,
        format_day,
        format_hm,
        style_offsets,
        count,
        gutter,
        slot_names,
    )
    document.showPage()
    document.save()


def register_font() -> None:
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    if FONT_NAME in pdfmetrics.getRegisteredFontNames():
        return
    for path in FONT_CANDIDATES:
        if not os.path.exists(path):
            continue
        pdfmetrics.registerFont(TTFont(FONT_NAME, path, subfontIndex=0))
        return
    raise RuntimeError("日本語フォントが見つからないため、PDFを作れません。")


def iter_days(start: datetime, end: datetime):
    day = datetime(start.year, start.month, start.day)
    while day < end:
        yield day
        day += timedelta(days=1)


def draw_day(document, page_w, page_h, tasks, day, alert_at, file_name, sheet_name, colors, effort_color, format_day, format_hm, style_offsets, slot_count, slot_names=None):
    margin = 28
    day_end = day + timedelta(days=1)
    document.setFillColor(rl("#1b2428"))
    document.setFont(FONT_NAME, 11)
    title = file_name or "スケジュール"
    if sheet_name:
        title = f"{title} / {sheet_name}"
    document.drawString(margin, page_h - 22, f"{title}    {format_day(day)}")
    if alert_at is not None:
        document.setFillColor(rl("#9f1239"))
        document.setFont(FONT_NAME, 9)
        document.drawRightString(page_w - margin, page_h - 22, f"工数 {trim_number(alert_at)} 以上を赤")

    gutter = fitted_gutter(slot_names, slot_count, 42, 8)
    left = margin + gutter
    right = page_w - margin
    top = page_h - 36
    bottom = margin
    axis_h = 22
    effort_h = 16
    slot_h = (top - bottom - axis_h - effort_h) / slot_count
    width = right - left

    document.setFillColor(rl("#f6fafb"))
    document.rect(left, top - axis_h, width, axis_h, fill=1, stroke=0)
    document.setFillColor(rl("#e7f3f2"))
    document.rect(left, top - axis_h - effort_h, width, effort_h, fill=1, stroke=0)
    document.setFillColor(rl("#f7fbfb"))
    document.rect(margin, bottom, gutter, top - bottom, fill=1, stroke=0)

    document.setFillColor(rl("#5d6c74"))
    document.setFont(FONT_NAME, 7)
    document.drawCentredString(margin + gutter / 2, top - axis_h + 8, "時間")
    document.setFillColor(rl("#0e6b67"))
    document.drawCentredString(margin + gutter / 2, top - axis_h - effort_h + 5, "工数")

    for slot in range(slot_count):
        y = top - axis_h - effort_h - (slot + 1) * slot_h
        if slot % 2:
            document.setFillColor(rl("#f7fbfb"))
            document.rect(left, y, width, slot_h, fill=1, stroke=0)
        document.setStrokeColor(rl("#e1e8ec"))
        document.setLineWidth(0.4)
        document.line(left, y + slot_h, right, y + slot_h)
        document.setFillColor(rl("#1b2428"))
        document.setFont(FONT_NAME, 8)
        document.drawCentredString(margin + gutter / 2, y + slot_h / 2 - 3, slot_caption(slot, slot_names))

    for hour in range(25):
        x = left + hour / 24 * width
        midnight = hour in {0, 24}
        document.setStrokeColor(rl("#c5ced6" if midnight else "#e7eef2"))
        document.setLineWidth(0.6 if midnight else 0.3)
        document.line(x, bottom, x, top if midnight else top - 8)
        if hour < 24:
            document.setStrokeColor(rl("#f1f5f7"))
            document.setLineWidth(0.3)
            document.line(x + width / 48, top - axis_h + 4, x + width / 48, bottom)
            moment = day + timedelta(hours=hour)
            document.setFillColor(rl("#5d6c74"))
            document.setFont(FONT_NAME, 7)
            document.drawString(x + 2, top - 16, format_hm(moment))

    draw_effort(document, tasks, day, day_end, left, width, top - axis_h - effort_h, effort_h, alert_at, effort_color)
    visible = [task for task in tasks if intersects(task, day, day_end)]
    for task in visible:
        draw_task_lines(document, tasks, task, day, day_end, left, width, top, axis_h, effort_h, slot_h, colors, style_offsets)
    for task in visible:
        draw_task_label(document, tasks, task, day, day_end, left, width, top, axis_h, effort_h, slot_h, colors)
    for task in visible:
        draw_task_points(document, tasks, task, day, day_end, left, width, top, axis_h, effort_h, slot_h, colors)


def draw_overview(document, page_w, page_h, tasks, origin, origin_end, alert_at, file_name, sheet_name, colors, effort_color, format_day, format_hm, style_offsets, slot_count, gutter=OVERVIEW_GUTTER, slot_names=None):
    margin = OVERVIEW_MARGIN
    left = margin + gutter
    right = page_w - margin
    top = page_h - margin
    bottom = margin
    axis_h = OVERVIEW_AXIS_H
    effort_h = OVERVIEW_EFFORT_H
    slot_h = OVERVIEW_SLOT_H
    width = right - left
    chart_top = top - OVERVIEW_HEADER

    document.setFillColor(paper())
    document.rect(0, 0, page_w, page_h, fill=1, stroke=0)
    document.setFillColor(rl("#1b2428"))
    document.setFont(FONT_NAME, 12)
    title = file_name or "スケジュール"
    if sheet_name:
        title = f"{title} / {sheet_name}"
    start_label = format_day(origin)
    end_label = format_day(origin_end - timedelta(days=1))
    span = start_label if start_label == end_label else f"{start_label} 〜 {end_label}"
    document.drawString(margin, top - 16, f"{title}    {span}")
    if alert_at is not None:
        document.setFillColor(rl("#9f1239"))
        document.setFont(FONT_NAME, 9)
        document.drawRightString(page_w - margin, top - 16, f"工数 {trim_number(alert_at)} 以上を赤")

    document.setFillColor(rl("#f6fafb"))
    document.rect(left, chart_top - axis_h, width, axis_h, fill=1, stroke=0)
    document.setFillColor(rl("#e7f3f2"))
    document.rect(left, chart_top - axis_h - effort_h, width, effort_h, fill=1, stroke=0)
    document.setFillColor(rl("#f7fbfb"))
    document.rect(margin, bottom, gutter, chart_top - bottom, fill=1, stroke=0)

    document.setFillColor(rl("#5d6c74"))
    document.setFont(FONT_NAME, 8)
    document.drawCentredString(margin + gutter / 2, chart_top - axis_h + 10, "時間")
    document.setFillColor(rl("#0e6b67"))
    document.drawCentredString(margin + gutter / 2, chart_top - axis_h - effort_h + 5, "工数")

    for slot in range(slot_count):
        y = chart_top - axis_h - effort_h - (slot + 1) * slot_h
        if slot % 2:
            document.setFillColor(rl("#f7fbfb"))
            document.rect(left, y, width, slot_h, fill=1, stroke=0)
        document.setStrokeColor(rl("#e1e8ec"))
        document.setLineWidth(0.4)
        document.line(left, y + slot_h, right, y + slot_h)
        document.setFillColor(rl("#1b2428"))
        document.setFont(FONT_NAME, 9)
        document.drawCentredString(margin + gutter / 2, y + slot_h / 2 - 3, slot_caption(slot, slot_names))

    hours = max((origin_end - origin).total_seconds() / 3600, 1 / 60)
    ticks = int(round(hours))
    for index in range(ticks + 1):
        moment = origin + timedelta(hours=index)
        x = left + (index / hours) * width
        midnight = moment.hour == 0 and moment.minute == 0
        document.setStrokeColor(rl("#c5ced6" if midnight else "#e7eef2"))
        document.setLineWidth(0.8 if midnight else 0.35)
        document.line(x, bottom, x, chart_top if midnight else chart_top - 8)
        if index >= ticks:
            continue
        document.setFont(FONT_NAME, 8)
        if midnight:
            document.setFillColor(rl("#1b2428"))
            document.drawString(x + 3, chart_top - 11, format_day(moment))
        document.setFillColor(rl("#5d6c74"))
        document.drawString(x + 3, chart_top - 23, format_hm(moment))

    draw_effort(
        document, tasks, origin, origin_end, left, width,
        chart_top - axis_h - effort_h, effort_h, alert_at, effort_color, origin_end,
    )
    visible = [task for task in tasks if intersects(task, origin, origin_end)]
    for task in visible:
        draw_task_lines(document, tasks, task, origin, origin_end, left, width, chart_top, axis_h, effort_h, slot_h, colors, style_offsets, origin_end)
    for task in visible:
        draw_task_label(document, tasks, task, origin, origin_end, left, width, chart_top, axis_h, effort_h, slot_h, colors, origin_end)
    for task in visible:
        draw_task_points(document, tasks, task, origin, origin_end, left, width, chart_top, axis_h, effort_h, slot_h, colors, origin_end)


def draw_effort(document, tasks, day, day_end, left, width, y, height, alert_at, effort_color, origin_end=None):
    for span in effort_spans(tasks):
        clipped = clip(span.start, span.end, day, day_end)
        if clipped is None:
            continue
        x1 = x_at(clipped[0], day, left, width, origin_end)
        x2 = x_at(clipped[1], day, left, width, origin_end)
        fill, text = effort_color(span.total, alert_at)
        document.setFillColor(rl(fill))
        document.rect(x1, y, max(0.4, x2 - x1), height, fill=1, stroke=0)
        label = trim_number(span.total)
        if x2 - x1 >= 12:
            document.setFillColor(rl(text))
            document.setFont(FONT_NAME, 7)
            document.drawCentredString((x1 + x2) / 2, y + 4, label)
        document.setStrokeColor(rl("#0e6b67"))
        document.setLineWidth(0.8)
        document.line(x1, y + height + 3, x1, y)
        document.line(x2, y + height + 3, x2, y)


def draw_task_lines(document, tasks, task, day, day_end, left, width, top, axis_h, effort_h, slot_h, colors, style_offsets, origin_end=None):
    color = line_color(tasks, task, colors)
    for index, point in enumerate(task.points[:-1]):
        nxt = task.points[index + 1]
        clipped = clip(point.time, nxt.time, day, day_end)
        if clipped is None:
            continue
        style = task.segments[index].style if index < len(task.segments) else "solid"
        y = lane_y(tasks, task, top, axis_h, effort_h, slot_h)
        x1 = x_at(clipped[0], day, left, width, origin_end)
        x2 = x_at(clipped[1], day, left, width, origin_end)
        stroke_segment(document, x1, y, x2, style, color, slot_h, style_offsets)


def draw_task_label(document, tasks, task, day, day_end, left, width, top, axis_h, effort_h, slot_h, colors, origin_end=None):
    pieces = []
    for index, point in enumerate(task.points[:-1]):
        clipped = clip(point.time, task.points[index + 1].time, day, day_end)
        if clipped:
            pieces.append(clipped)
    if not pieces and not any(day <= point.time < day_end for point in task.points):
        return
    if pieces:
        start = min(item[0] for item in pieces)
        end = max(item[1] for item in pieces)
    else:
        start = end = next(point.time for point in task.points if day <= point.time < day_end)
    color = line_color(tasks, task, colors)
    y = lane_y(tasks, task, top, axis_h, effort_h, slot_h)
    document.setFont(FONT_NAME, 7)
    text_w = document.stringWidth(task.name, FONT_NAME, 7) + 8
    x = x_at(start, day, left, width, origin_end)
    limit = left + width - text_w - 2
    if end > start:
        x = min(max(x, x_at(start, day, left, width, origin_end)), max(x_at(start, day, left, width, origin_end), x_at(end, day, left, width, origin_end) - text_w))
    x = min(max(left + 2, x), limit)
    document.setStrokeColor(rl(color))
    document.setFillColor(paper())
    document.setLineWidth(0.8)
    document.roundRect(x, y + 3, text_w, 11, 2, fill=1, stroke=1)
    document.setFillColor(rl("#1b2428"))
    document.drawString(x + 4, y + 6, task.name)


def draw_task_points(document, tasks, task, day, day_end, left, width, top, axis_h, effort_h, slot_h, colors, origin_end=None):
    color = line_color(tasks, task, colors)
    y = lane_y(tasks, task, top, axis_h, effort_h, slot_h)
    document.setFillColor(rl(color))
    document.setStrokeColor(paper())
    document.setLineWidth(0.8)
    for point in task.points:
        if not (day <= point.time < day_end):
            continue
        x = x_at(point.time, day, left, width, origin_end)
        document.circle(x, y, 2.4, fill=1, stroke=1)


def stroke_segment(document, x1, y, x2, style, color, slot_h, style_offsets):
    scale = slot_h / 78
    offsets = tuple(value * scale for value in style_offsets(style))
    document.setLineCap(1)
    if style == "none":
        document.setStrokeColor(rl("#c5ced6"))
        document.setDash(1, 2)
        document.setLineWidth(0.6)
        document.line(x1, y, x2, y)
        document.setDash()
        return
    if style != "dashed":
        halo = {"double": 5, "triple": 7.5, "quad": 10, "quint": 12.5}.get(style, 2.4) * scale
        document.setStrokeColor(paper())
        document.setLineWidth(max(1.2, halo))
        document.line(x1, y, x2, y)
    document.setStrokeColor(rl(color))
    document.setLineWidth(1.1)
    if style == "dashed":
        document.setDash(3, 2)
    else:
        document.setDash()
    for offset in offsets or (0,):
        document.line(x1, y - offset, x2, y - offset)
    document.setDash()


def lane_y(tasks, task, top, axis_h, effort_h, slot_h):
    start, end = task_span(task)
    peers = [item for item in tasks if item.slot == task.slot and ranges_touch(task_span(item), (start, end))]
    peers.sort(key=lambda item: item.id)
    index = next((position for position, item in enumerate(peers) if item.id == task.id), 0)
    center = top - axis_h - effort_h - (task.slot + 0.5) * slot_h
    if len(peers) <= 1:
        return center
    spread = min(10, max(0, (slot_h - 16) / 2))
    step = (spread * 2) / (len(peers) - 1)
    return center + spread - index * step


def task_index(tasks, task) -> int:
    return next((position for position, item in enumerate(tasks) if item.id == task.id), 0)


def line_color(tasks, task, colors) -> str:
    chosen = str(getattr(task, "color", "") or "")
    if chosen in colors:
        return chosen
    return colors[task_index(tasks, task) % len(colors)]


def intersects(task, day, day_end) -> bool:
    start, end = task_span(task)
    return start < day_end and end > day


def ranges_touch(left, right) -> bool:
    return left[0] < right[1] and left[1] > right[0]


def clip(start, end, day, day_end):
    left = start if start > day else day
    right = end if end < day_end else day_end
    if right <= left:
        return None
    return left, right


def x_at(moment, day, left, width, origin_end=None):
    span = 24.0 if origin_end is None else max((origin_end - day).total_seconds() / 3600, 1 / 60)
    hours = (moment - day).total_seconds() / 3600
    hours = 0 if hours < 0 else span if hours > span else hours
    return left + hours / span * width


def rl(hex_color: str):
    from reportlab.lib.colors import Color

    value = hex_color.lstrip("#")
    return Color(int(value[0:2], 16) / 255, int(value[2:4], 16) / 255, int(value[4:6], 16) / 255)


def paper():
    from reportlab.lib.colors import white

    return white


def trim_number(value: float) -> str:
    rounded = round(float(value) * 10) / 10
    if float(rounded).is_integer():
        return str(int(rounded))
    return f"{rounded:.1f}"
