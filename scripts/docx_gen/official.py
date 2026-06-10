"""
党政公文模板（GB/T 9704-2012）

提供：
  - generate_tongzhi()  → 通知
  - generate_qingshi()  → 请示
  - generate_baogao()   → 报告
  - generate_han()      → 函
"""

import os
from datetime import datetime

from .base import ChineseDocument

try:
    from ..chinese_format import (
        PAGE_OFFICIAL, FontPreset, ParagraphPreset,
        set_run_font, apply_paragraph_format,
    )
except ImportError:
    from chinese_format import (
        PAGE_OFFICIAL, FontPreset, ParagraphPreset,
        set_run_font, apply_paragraph_format,
    )


# ── 工具函数 ──────────────────────────────────────────────

def _make_meta_line(doc, text: str):
    """添加公文元数据行（发文字号、签发人等），仿宋、无缩进。"""
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, FontPreset.OFFICIAL_META.cn, FontPreset.OFFICIAL_META.en,
                 FontPreset.OFFICIAL_META.size_pt)
    apply_paragraph_format(p, ParagraphPreset(
        first_line_indent_chars=0, line_spacing=28.0 / 12.0,
    ))
    return p


def _make_official_body(doc, text: str):
    """添加公文正文段落（仿宋三号，首行缩进2字符，28磅行距）。"""
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, FontPreset.OFFICIAL_BODY.cn, FontPreset.OFFICIAL_BODY.en,
                 FontPreset.OFFICIAL_BODY.size_pt)
    apply_paragraph_format(p, ParagraphPreset(
        first_line_indent_chars=2, line_spacing=28.0 / 12.0,
        alignment=3,  # JUSTIFY
    ))
    return p


def _make_official_heading(doc, text: str, level: int = 1):
    """添加公文标题。一级=黑体三号，二级=楷体三号加粗。"""
    font = FontPreset.OFFICIAL_H1 if level == 1 else FontPreset.OFFICIAL_H2
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, font.cn, font.en, font.size_pt, font.bold)
    apply_paragraph_format(p, ParagraphPreset(
        first_line_indent_chars=0, line_spacing=28.0 / 12.0,
    ))
    return p


# ── 通知 ──────────────────────────────────────────────────

def generate_tongzhi(
    output_path: str,
    title: str = "通知",
    issuing_authority: str = "",
    doc_number: str = "",
    recipients: str = "",
    body_paragraphs: list = None,
    issuing_date: str = "",
    contact: str = "",
) -> str:
    """生成党政机关公文——通知。

    Args:
        output_path: 输出路径
        title: 通知标题
        issuing_authority: 发文机关（如"XX大学教务处"）
        doc_number: 发文字号（如"XX发〔2026〕1号"）
        recipients: 主送机关（如"各学院："）
        body_paragraphs: 正文段落列表
        issuing_date: 发文日期
        contact: 联系方式（可选）

    Returns:
        输出文件路径
    """
    if body_paragraphs is None:
        body_paragraphs = ["[通知正文内容]"]
    if not issuing_date:
        issuing_date = datetime.now().strftime("%Y年%m月%d日")

    doc = ChineseDocument.create(PAGE_OFFICIAL)

    # 发文机关标志（红头，此处用黑体代替，Word中可设红色）
    if issuing_authority:
        p = doc.add_paragraph()
        run = p.add_run(issuing_authority)
        set_run_font(run, "黑体", "Times New Roman", 26, bold=True, color="FF0000")
        apply_paragraph_format(p, ParagraphPreset(
            first_line_indent_chars=0, line_spacing=1.0,
            alignment=1,  # CENTER
        ))

    # 发文字号
    if doc_number:
        _make_meta_line(doc, doc_number)

    # 标题
    ChineseDocument.add_title(doc, title, large=False)

    # 主送机关
    if recipients:
        p = doc.add_paragraph()
        run = p.add_run(recipients)
        set_run_font(run, FontPreset.OFFICIAL_BODY.cn, FontPreset.OFFICIAL_BODY.en,
                     FontPreset.OFFICIAL_BODY.size_pt)
        apply_paragraph_format(p, ParagraphPreset(
            first_line_indent_chars=0, line_spacing=28.0 / 12.0,
        ))

    # 正文
    for para in body_paragraphs:
        _make_official_body(doc, para)

    # 联系方式
    if contact:
        _make_official_body(doc, contact)

    # 落款
    ChineseDocument.add_blank(doc)
    if issuing_authority:
        ChineseDocument.add_signature(doc, issuing_authority)
    ChineseDocument.add_signature(doc, issuing_date)

    return ChineseDocument.save(doc, output_path)


# ── 请示 ──────────────────────────────────────────────────

def generate_qingshi(
    output_path: str,
    title: str = "请示",
    issuing_authority: str = "",
    doc_number: str = "",
    recipients: str = "",
    reason_paragraphs: list = None,
    matter_paragraphs: list = None,
    conclusion: str = "妥否，请批示。",
    contact: str = "",
    issuing_date: str = "",
) -> str:
    """生成党政机关公文——请示。

    结构：请示缘由 → 请示事项 → 结语 → 落款

    Args:
        output_path: 输出路径
        title: 请示标题（如"关于增加人员编制的请示"）
        issuing_authority: 发文机关
        doc_number: 发文字号
        recipients: 主送机关
        reason_paragraphs: 请示缘由段落
        matter_paragraphs: 请示事项段落
        conclusion: 结语（默认"妥否，请批示。"）
        contact: 联系人及电话
        issuing_date: 发文日期

    Returns:
        输出文件路径
    """
    if reason_paragraphs is None:
        reason_paragraphs = ["[请示缘由]"]
    if matter_paragraphs is None:
        matter_paragraphs = ["[请示事项]"]
    if not issuing_date:
        issuing_date = datetime.now().strftime("%Y年%m月%d日")

    doc = ChineseDocument.create(PAGE_OFFICIAL)

    # 发文机关标志
    if issuing_authority:
        p = doc.add_paragraph()
        run = p.add_run(issuing_authority)
        set_run_font(run, "黑体", "Times New Roman", 26, bold=True)
        apply_paragraph_format(p, ParagraphPreset(
            first_line_indent_chars=0, line_spacing=1.0, alignment=1,
        ))

    if doc_number:
        _make_meta_line(doc, doc_number)

    ChineseDocument.add_title(doc, title, large=False)

    if recipients:
        p = doc.add_paragraph()
        run = p.add_run(recipients)
        set_run_font(run, FontPreset.OFFICIAL_BODY.cn, FontPreset.OFFICIAL_BODY.en,
                     FontPreset.OFFICIAL_BODY.size_pt)
        apply_paragraph_format(p, ParagraphPreset(
            first_line_indent_chars=0, line_spacing=28.0 / 12.0,
        ))

    # 请示缘由
    for para in reason_paragraphs:
        _make_official_body(doc, para)

    # 请示事项
    if matter_paragraphs:
        _make_official_body(doc, "")  # 空行分隔
        for para in matter_paragraphs:
            _make_official_body(doc, para)

    # 结语
    _make_official_body(doc, conclusion)

    # 联系人
    if contact:
        _make_official_body(doc, contact)

    ChineseDocument.add_blank(doc)
    if issuing_authority:
        ChineseDocument.add_signature(doc, issuing_authority)
    ChineseDocument.add_signature(doc, issuing_date)

    return ChineseDocument.save(doc, output_path)


# ── 报告 ──────────────────────────────────────────────────

def generate_baogao(
    output_path: str,
    title: str = "报告",
    issuing_authority: str = "",
    doc_number: str = "",
    recipients: str = "",
    foreword: str = "",
    sections: list = None,
    conclusion: str = "",
    issuing_date: str = "",
) -> str:
    """生成党政机关公文——报告。

    结构：前言 → 主体（分章节）→ 结语 → 落款

    Args:
        output_path: 输出路径
        title: 报告标题
        issuing_authority: 发文机关
        doc_number: 发文字号
        recipients: 主送机关
        foreword: 前言
        sections: 主体章节列表，每项为 {"heading": "...", "body": ["段落1", "段落2"]}
        conclusion: 结语（如"特此报告。"）
        issuing_date: 发文日期

    Returns:
        输出文件路径
    """
    if sections is None:
        sections = [{"heading": "一、基本情况", "body": ["[报告正文]"]}]
    if not issuing_date:
        issuing_date = datetime.now().strftime("%Y年%m月%d日")

    doc = ChineseDocument.create(PAGE_OFFICIAL)

    if issuing_authority:
        p = doc.add_paragraph()
        run = p.add_run(issuing_authority)
        set_run_font(run, "黑体", "Times New Roman", 26, bold=True)
        apply_paragraph_format(p, ParagraphPreset(
            first_line_indent_chars=0, line_spacing=1.0, alignment=1,
        ))

    if doc_number:
        _make_meta_line(doc, doc_number)

    ChineseDocument.add_title(doc, title, large=False)

    if recipients:
        p = doc.add_paragraph()
        run = p.add_run(recipients)
        set_run_font(run, FontPreset.OFFICIAL_BODY.cn, FontPreset.OFFICIAL_BODY.en,
                     FontPreset.OFFICIAL_BODY.size_pt)
        apply_paragraph_format(p, ParagraphPreset(
            first_line_indent_chars=0, line_spacing=28.0 / 12.0,
        ))

    # 前言
    if foreword:
        _make_official_body(doc, foreword)

    # 主体章节
    for sec in sections:
        heading = sec.get("heading", "")
        body = sec.get("body", [])
        if heading:
            _make_official_heading(doc, heading, level=1)
        for para in body:
            _make_official_body(doc, para)

    # 结语
    if conclusion:
        _make_official_body(doc, conclusion)

    ChineseDocument.add_blank(doc)
    if issuing_authority:
        ChineseDocument.add_signature(doc, issuing_authority)
    ChineseDocument.add_signature(doc, issuing_date)

    return ChineseDocument.save(doc, output_path)


# ── 函 ────────────────────────────────────────────────────

def generate_han(
    output_path: str,
    title: str = "函",
    sender: str = "",
    receiver: str = "",
    doc_number: str = "",
    body_paragraphs: list = None,
    contact: str = "",
    issuing_date: str = "",
) -> str:
    """生成党政机关公文——函。

    Args:
        output_path: 输出路径
        title: 函的标题
        sender: 发函机关
        receiver: 收函机关
        doc_number: 发文字号
        body_paragraphs: 正文段落
        contact: 联系方式
        issuing_date: 发文日期

    Returns:
        输出文件路径
    """
    if body_paragraphs is None:
        body_paragraphs = ["[函的正文内容]"]
    if not issuing_date:
        issuing_date = datetime.now().strftime("%Y年%m月%d日")

    doc = ChineseDocument.create(PAGE_OFFICIAL)

    if sender:
        p = doc.add_paragraph()
        run = p.add_run(sender)
        set_run_font(run, "黑体", "Times New Roman", 26, bold=True)
        apply_paragraph_format(p, ParagraphPreset(
            first_line_indent_chars=0, line_spacing=1.0, alignment=1,
        ))

    if doc_number:
        _make_meta_line(doc, doc_number)

    ChineseDocument.add_title(doc, title, large=False)

    if receiver:
        p = doc.add_paragraph()
        run = p.add_run(receiver)
        set_run_font(run, FontPreset.OFFICIAL_BODY.cn, FontPreset.OFFICIAL_BODY.en,
                     FontPreset.OFFICIAL_BODY.size_pt)
        apply_paragraph_format(p, ParagraphPreset(
            first_line_indent_chars=0, line_spacing=28.0 / 12.0,
        ))

    for para in body_paragraphs:
        _make_official_body(doc, para)

    if contact:
        _make_official_body(doc, contact)

    ChineseDocument.add_blank(doc)
    if sender:
        ChineseDocument.add_signature(doc, sender)
    ChineseDocument.add_signature(doc, issuing_date)

    return ChineseDocument.save(doc, output_path)
