"""
学术论文模板（GB/T 7713.2-2022）

提供：
  - generate_course_paper()  → 课程论文
  - generate_thesis()        → 学位论文
  - generate_proposal()      → 开题报告
"""

import os
from datetime import datetime

from .base import ChineseDocument

try:
    from ..chinese_format import (
        PAGE_ACADEMIC, PAGE_DEFAULT, FontPreset, ParagraphPreset,
        set_run_font, apply_paragraph_format,
    )
except ImportError:
    from chinese_format import (
        PAGE_ACADEMIC, PAGE_DEFAULT, FontPreset, ParagraphPreset,
        set_run_font, apply_paragraph_format,
    )


# ── 工具函数 ──────────────────────────────────────────────

def _add_meta_info(doc, lines: list):
    """添加元数据信息行（居中、宋体、小四号）。"""
    for text in lines:
        if text:
            p = doc.add_paragraph()
            run = p.add_run(text)
            set_run_font(run, "宋体", "Times New Roman", 12)
            apply_paragraph_format(p, ParagraphPreset(
                first_line_indent_chars=0, line_spacing=1.5, alignment=1,  # CENTER
            ))


def _add_abstract_section(doc, abstract_text: str, keywords: list = None):
    """添加摘要和关键词。"""
    # 摘要标签
    p = doc.add_paragraph()
    run_label = p.add_run("摘  要：")
    set_run_font(run_label, "黑体", "Times New Roman", 12, bold=True)
    run_text = p.add_run(abstract_text)
    set_run_font(run_text, "宋体", "Times New Roman", 12)
    apply_paragraph_format(p, ParagraphPreset(
        first_line_indent_chars=0, line_spacing=1.25, alignment=3,
    ))

    # 关键词
    if keywords:
        p = doc.add_paragraph()
        run_label = p.add_run("关键词：")
        set_run_font(run_label, "黑体", "Times New Roman", 12, bold=True)
        kw_text = "；".join(keywords)
        run_text = p.add_run(kw_text)
        set_run_font(run_text, "宋体", "Times New Roman", 12)
        apply_paragraph_format(p, ParagraphPreset(
            first_line_indent_chars=0, line_spacing=1.25, alignment=3,
        ))


# ── 课程论文 ──────────────────────────────────────────────

def generate_course_paper(
    output_path: str,
    title: str = "课程论文",
    author: str = "",
    course: str = "",
    instructor: str = "",
    student_id: str = "",
    abstract_text: str = "",
    keywords: list = None,
    sections: list = None,
    references: list = None,
    date_str: str = "",
) -> str:
    """生成课程论文。

    Args:
        output_path: 输出路径
        title: 论文标题
        author: 作者姓名
        course: 课程名称
        instructor: 指导教师
        student_id: 学号
        abstract_text: 摘要
        keywords: 关键词列表
        sections: 章节列表 [{"heading": "...", "body": ["段落1", "段落2"]}]
        references: 参考文献列表 ["[1] ...", "[2] ..."]
        date_str: 日期

    Returns:
        输出文件路径
    """
    if sections is None:
        sections = [{"heading": "一、引言", "body": ["[正文内容]"]}]
    if not date_str:
        date_str = datetime.now().strftime("%Y年%m月%d日")

    doc = ChineseDocument.create(PAGE_ACADEMIC)

    # 标题页
    ChineseDocument.add_title(doc, title, large=True)

    # 作者信息
    meta_lines = []
    if author:
        meta_lines.append(f"作者：{author}")
    if student_id:
        meta_lines.append(f"学号：{student_id}")
    if course:
        meta_lines.append(f"课程：{course}")
    if instructor:
        meta_lines.append(f"指导教师：{instructor}")
    if meta_lines:
        ChineseDocument.add_blank(doc)
        _add_meta_info(doc, meta_lines)

    ChineseDocument.add_blank(doc, 2)

    # 摘要
    if abstract_text:
        ChineseDocument.add_heading(doc, "摘  要", level=1)
        _add_abstract_section(doc, abstract_text, keywords)
        ChineseDocument.add_blank(doc)

    # 正文章节
    for i, sec in enumerate(sections):
        heading = sec.get("heading", "")
        body = sec.get("body", [])

        if heading:
            ChineseDocument.add_heading(doc, heading, level=1)

        for para in body:
            if isinstance(para, dict) and "heading" in para:
                ChineseDocument.add_heading(doc, para["heading"], level=2)
                for sub_para in para.get("body", []):
                    ChineseDocument.add_body(doc, sub_para)
            else:
                ChineseDocument.add_body(doc, para)

        if i < len(sections) - 1:
            ChineseDocument.add_blank(doc)

    # 参考文献
    if references:
        ChineseDocument.add_blank(doc)
        ChineseDocument.add_heading(doc, "参考文献", level=1)
        for ref in references:
            p = doc.add_paragraph()
            run = p.add_run(ref)
            set_run_font(run, "宋体", "Times New Roman", 10.5)
            apply_paragraph_format(p, ParagraphPreset(
                first_line_indent_chars=0, line_spacing=1.25,
            ))

    return ChineseDocument.save(doc, output_path)


# ── 学位论文 ──────────────────────────────────────────────

def generate_thesis(
    output_path: str,
    title: str = "学位论文",
    author: str = "",
    student_id: str = "",
    school: str = "",
    major: str = "",
    instructor: str = "",
    abstract_cn: str = "",
    keywords_cn: list = None,
    abstract_en: str = "",
    keywords_en: list = None,
    chapters: list = None,
    acknowledgments: str = "",
    references: list = None,
    date_str: str = "",
) -> str:
    """生成学位论文。

    Args:
        output_path: 输出路径
        title: 论文题目
        author: 学生姓名
        student_id: 学号
        school: 学院
        major: 专业
        instructor: 指导教师（含职称）
        abstract_cn: 中文摘要
        keywords_cn: 中文关键词
        abstract_en: 英文摘要
        keywords_en: 英文关键词
        chapters: 章节列表 [{"heading": "第一章 xxx", "sections": [...]}]
        acknowledgments: 致谢
        references: 参考文献
        date_str: 日期

    Returns:
        输出文件路径
    """
    if chapters is None:
        chapters = [{"heading": "第一章  绪论", "sections": [
            {"heading": "1.1  研究背景", "body": ["[研究背景内容]"]},
        ]}]
    if not date_str:
        date_str = datetime.now().strftime("%Y年%m月")

    doc = ChineseDocument.create(PAGE_ACADEMIC)

    # ── 封面 ──
    ChineseDocument.add_title(doc, title, large=True)
    ChineseDocument.add_blank(doc, 3)

    cover_info = []
    if school:
        cover_info.append(f"所在学院：{school}")
    if major:
        cover_info.append(f"专    业：{major}")
    if author:
        cover_info.append(f"学生姓名：{author}")
    if student_id:
        cover_info.append(f"学    号：{student_id}")
    if instructor:
        cover_info.append(f"指导教师：{instructor}")
    if date_str:
        cover_info.append(date_str)
    _add_meta_info(doc, cover_info)

    # 分页 → 中文摘要
    p = doc.add_paragraph()
    apply_paragraph_format(p, ParagraphPreset(first_line_indent_chars=0, line_spacing=1.0))

    ChineseDocument.add_heading(doc, "摘  要", level=1)
    if abstract_cn:
        _add_abstract_section(doc, abstract_cn, keywords_cn)
    ChineseDocument.add_blank(doc)

    # 英文摘要
    if abstract_en:
        ChineseDocument.add_heading(doc, "Abstract", level=1)
        p = doc.add_paragraph()
        run = p.add_run(abstract_en)
        set_run_font(run, "Times New Roman", "Times New Roman", 12)
        apply_paragraph_format(p, ParagraphPreset(
            first_line_indent_chars=0, line_spacing=1.25,
        ))
        if keywords_en:
            p = doc.add_paragraph()
            run = p.add_run(f"Keywords: {', '.join(keywords_en)}")
            set_run_font(run, "Times New Roman", "Times New Roman", 12)
            apply_paragraph_format(p, ParagraphPreset(
                first_line_indent_chars=0, line_spacing=1.25,
            ))

    ChineseDocument.add_blank(doc)

    # ── 正文各章 ──
    for chapter in chapters:
        ch_heading = chapter.get("heading", "")
        if ch_heading:
            ChineseDocument.add_heading(doc, ch_heading, level=1)

        for sec in chapter.get("sections", []):
            sec_heading = sec.get("heading", "")
            sec_body = sec.get("body", [])

            if sec_heading:
                ChineseDocument.add_heading(doc, sec_heading, level=2)

            for para in sec_body:
                ChineseDocument.add_body(doc, para)

            ChineseDocument.add_blank(doc)

    # ── 致谢 ──
    if acknowledgments:
        ChineseDocument.add_heading(doc, "致  谢", level=1)
        ChineseDocument.add_body(doc, acknowledgments)
        ChineseDocument.add_blank(doc)

    # ── 参考文献 ──
    if references:
        ChineseDocument.add_heading(doc, "参考文献", level=1)
        for ref in references:
            p = doc.add_paragraph()
            run = p.add_run(ref)
            set_run_font(run, "宋体", "Times New Roman", 10.5)
            apply_paragraph_format(p, ParagraphPreset(
                first_line_indent_chars=0, line_spacing=1.25,
            ))

    return ChineseDocument.save(doc, output_path)


# ── 开题报告 ──────────────────────────────────────────────

def generate_proposal(
    output_path: str,
    title: str = "开题报告",
    author: str = "",
    student_id: str = "",
    school: str = "",
    major: str = "",
    instructor: str = "",
    background: str = "",
    literature_review: str = "",
    research_content: str = "",
    research_methods: str = "",
    expected_results: str = "",
    schedule: list = None,
    references: list = None,
    date_str: str = "",
) -> str:
    """生成开题报告。

    Args:
        output_path: 输出路径
        title: 课题名称
        author: 学生姓名
        student_id: 学号
        school: 学院
        major: 专业
        instructor: 指导教师
        background: 选题背景与意义
        literature_review: 国内外研究现状
        research_content: 研究内容
        research_methods: 研究方法与技术路线
        expected_results: 预期成果
        schedule: 进度安排 [{"stage": "...", "time": "..."}]
        references: 参考文献
        date_str: 日期

    Returns:
        输出文件路径
    """
    if not date_str:
        date_str = datetime.now().strftime("%Y年%m月%d日")

    doc = ChineseDocument.create(PAGE_ACADEMIC)

    # 封面
    ChineseDocument.add_title(doc, title, large=True)
    ChineseDocument.add_blank(doc, 2)

    info = []
    if author:
        info.append(f"学生姓名：{author}")
    if student_id:
        info.append(f"学    号：{student_id}")
    if school:
        info.append(f"学    院：{school}")
    if major:
        info.append(f"专    业：{major}")
    if instructor:
        info.append(f"指导教师：{instructor}")
    if date_str:
        info.append(date_str)
    _add_meta_info(doc, info)
    ChineseDocument.add_blank(doc, 2)

    # 各节
    section_data = [
        ("一、选题背景与意义", background),
        ("二、国内外研究现状", literature_review),
        ("三、研究内容", research_content),
        ("四、研究方法与技术路线", research_methods),
        ("五、预期成果", expected_results),
    ]

    for heading, content in section_data:
        if content:
            ChineseDocument.add_heading(doc, heading, level=1)
            # 将内容按段落分割
            for para in content.split("\n"):
                para = para.strip()
                if para:
                    ChineseDocument.add_body(doc, para)
            ChineseDocument.add_blank(doc)

    # 进度安排
    if schedule:
        ChineseDocument.add_heading(doc, "六、进度安排", level=1)
        for item in schedule:
            stage = item.get("stage", "")
            time = item.get("time", "")
            ChineseDocument.add_body(doc, f"{stage}：{time}")
        ChineseDocument.add_blank(doc)

    # 参考文献
    if references:
        ChineseDocument.add_heading(doc, "七、参考文献", level=1)
        for ref in references:
            p = doc.add_paragraph()
            run = p.add_run(ref)
            set_run_font(run, "宋体", "Times New Roman", 10.5)
            apply_paragraph_format(p, ParagraphPreset(
                first_line_indent_chars=0, line_spacing=1.25,
            ))

    return ChineseDocument.save(doc, output_path)
