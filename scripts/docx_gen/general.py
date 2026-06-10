"""
日常文书模板

提供：
  - generate_application()     → 申请书
  - generate_meeting_minutes() → 会议纪要
"""

import os
import sys
from datetime import datetime

from .base import ChineseDocument

try:
    from ..chinese_format import (
        PAGE_DEFAULT, FontPreset, ParagraphPreset,
        set_run_font, apply_paragraph_format,
    )
except ImportError:
    from chinese_format import (
        PAGE_DEFAULT, FontPreset, ParagraphPreset,
        set_run_font, apply_paragraph_format,
    )


def generate_application(
    output_path: str,
    title: str = "申请书",
    recipient: str = "",
    body_paragraphs: list = None,
    applicant: str = "",
    date_str: str = "",
    closing: str = "此致",
    salutation: str = "敬礼",
) -> str:
    """生成中文申请书。

    标准结构：
        标题（居中）→ 称呼（顶格）→ 正文（首行缩进2字符）
        → 结束语（此致/敬礼）→ 署名 + 日期（右对齐）

    Args:
        output_path: 输出文件路径（.docx）
        title: 申请书标题（默认"申请书"）
        recipient: 收信人/单位称呼（如"尊敬的XX："）
        body_paragraphs: 正文段落列表
        applicant: 申请人姓名
        date_str: 日期（如"2026年6月10日"，默认今天）
        closing: 结束语（默认"此致"）
        salutation: 敬语（默认"敬礼"）

    Returns:
        输出文件路径
    """
    if body_paragraphs is None:
        body_paragraphs = ["[请在此处填写申请正文内容]"]
    if not date_str:
        date_str = datetime.now().strftime("%Y年%m月%d日")

    doc = ChineseDocument.create(PAGE_DEFAULT)

    # 标题
    ChineseDocument.add_title(doc, title)
    ChineseDocument.add_blank(doc)

    # 称呼（顶格，无缩进）
    if recipient:
        p = doc.add_paragraph()
        p.clear()
        run = p.add_run(recipient)
        set_run_font(run, "宋体", "Times New Roman", 12)
        apply_paragraph_format(p, ParagraphPreset(
            first_line_indent_chars=0, line_spacing=1.25,
        ))

    # 正文
    for para_text in body_paragraphs:
        ChineseDocument.add_body(doc, para_text)

    ChineseDocument.add_blank(doc)

    # 结束语
    if closing:
        p = doc.add_paragraph()
        p.clear()
        run = p.add_run(closing)
        set_run_font(run, "宋体", "Times New Roman", 12)
        apply_paragraph_format(p, ParagraphPreset(
            first_line_indent_chars=0, line_spacing=1.25,
        ))

    if salutation:
        ChineseDocument.add_blank(doc)
        p = doc.add_paragraph()
        p.clear()
        run = p.add_run(salutation)
        set_run_font(run, "宋体", "Times New Roman", 12)
        apply_paragraph_format(p, ParagraphPreset(
            first_line_indent_chars=0, line_spacing=1.25,
        ))

    ChineseDocument.add_blank(doc, 2)

    # 署名和日期（右对齐）
    if applicant:
        ChineseDocument.add_signature(doc, f"申请人：{applicant}")
    ChineseDocument.add_signature(doc, date_str)

    return ChineseDocument.save(doc, output_path)


def generate_meeting_minutes(
    output_path: str,
    title: str = "会议纪要",
    meeting_info: dict = None,
    agenda_items: list = None,
    decisions: list = None,
    recorder: str = "",
    date_str: str = "",
) -> str:
    """生成中文会议纪要。

    标准结构：
        标题 → 会议信息（时间/地点/主持人/参会人员）
        → 会议议题 → 讨论内容 → 会议决议 → 记录人 + 日期

    Args:
        output_path: 输出文件路径（.docx）
        title: 纪要标题（默认"会议纪要"）
        meeting_info: 会议基本信息字典，键可选：
            time / location / host / attendees / absent
        agenda_items: 议题列表
        decisions: 决议列表
        recorder: 记录人
        date_str: 日期（默认今天）

    Returns:
        输出文件路径
    """
    if meeting_info is None:
        meeting_info = {}
    if agenda_items is None:
        agenda_items = ["[请填写会议议题]"]
    if decisions is None:
        decisions = []
    if not date_str:
        date_str = datetime.now().strftime("%Y年%m月%d日")

    doc = ChineseDocument.create(PAGE_DEFAULT)

    # 标题
    ChineseDocument.add_title(doc, title)
    ChineseDocument.add_blank(doc)

    # 会议基本信息
    info_fields = [
        ("会议时间", meeting_info.get("time", "")),
        ("会议地点", meeting_info.get("location", "")),
        ("主持人", meeting_info.get("host", "")),
    ]
    for label, value in info_fields:
        if value:
            p = doc.add_paragraph()
            run = p.add_run(f"{label}：{value}")
            set_run_font(run, "宋体", "Times New Roman", 12)
            apply_paragraph_format(p, ParagraphPreset(
                first_line_indent_chars=0, line_spacing=1.25,
            ))

    # 参会人员
    attendees = meeting_info.get("attendees", "")
    if attendees:
        if isinstance(attendees, list):
            attendees = "、".join(attendees)
        p = doc.add_paragraph()
        run = p.add_run(f"参会人员：{attendees}")
        set_run_font(run, "宋体", "Times New Roman", 12)
        apply_paragraph_format(p, ParagraphPreset(
            first_line_indent_chars=0, line_spacing=1.25,
        ))

    absent = meeting_info.get("absent", "")
    if absent:
        if isinstance(absent, list):
            absent = "、".join(absent)
        p = doc.add_paragraph()
        run = p.add_run(f"缺席人员：{absent}")
        set_run_font(run, "宋体", "Times New Roman", 12)
        apply_paragraph_format(p, ParagraphPreset(
            first_line_indent_chars=0, line_spacing=1.25,
        ))

    ChineseDocument.add_blank(doc)

    # 会议议题
    ChineseDocument.add_heading(doc, "一、会议议题", level=1)
    for item in agenda_items:
        ChineseDocument.add_body(doc, item)

    # 会议决议
    if decisions:
        ChineseDocument.add_blank(doc)
        ChineseDocument.add_heading(doc, "二、会议决议", level=1)
        for i, decision in enumerate(decisions, 1):
            ChineseDocument.add_body(doc, f"{i}. {decision}")

    ChineseDocument.add_blank(doc, 2)

    if recorder:
        ChineseDocument.add_signature(doc, f"记录人：{recorder}")
    ChineseDocument.add_signature(doc, date_str)

    return ChineseDocument.save(doc, output_path)
