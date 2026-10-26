"""格式检查器的 Word 文档格式检查核心。"""

import re
import json
from dataclasses import dataclass, field
from enum import Enum
from docx import Document
from docx.shared import Pt, Emu
from docx.oxml.ns import qn


DEFAULT_ENABLED_CHECKS = {
    "font", "alignment", "line_spacing", "structure", "cover",
    "title_numbering", "figure_table_numbering", "references",
    "page_setup", "header", "abstract", "word_count", "footnotes",
}


class Severity(Enum):
    ERROR = "错误"
    WARNING = "警告"
    INFO = "提示"


@dataclass
class Issue:
    location: str       # 位置描述
    severity: Severity
    category: str       # 问题分类
    description: str    # 问题描述
    expected: str       # 正确格式
    actual: str         # 实际格式
    content: str = ""   # 出错的原文内容


@dataclass
class CheckResult:
    filename: str
    issues: list[Issue] = field(default_factory=list)
    summary: dict = field(default_factory=dict)

    @property
    def error_count(self):
        return sum(1 for i in self.issues if i.severity == Severity.ERROR)

    @property
    def warning_count(self):
        return sum(1 for i in self.issues if i.severity == Severity.WARNING)

    @property
    def info_count(self):
        return sum(1 for i in self.issues if i.severity == Severity.INFO)


# ===== 字体大小常量（pt）=====
# 中文字号对照表
FONT_SIZE_MAP = {
    "初号": 42,
    "小初": 36,
    "一号": 26,
    "小一": 24,
    "二号": 22,
    "小二": 18,
    "三号": 16,
    "小三": 15,
    "四号": 14,
    "小四": 12,
    "五号": 10.5,
    "小五": 9,
    "六号": 7.5,
    "小六": 6.5,
    "七号": 5.5,
    "八号": 5,
}


def emu_to_pt(emu_val):
    """EMU 转换为磅值"""
    if emu_val is None:
        return None
    return round(emu_val / 12700, 1)


def pt_to_size_name(pt_val):
    """pt 值转换为中文字号名称，如 12pt → 小四，15pt → 小三"""
    if pt_val is None:
        return ""
    # 按 pt 值从大到小匹配，允许 0.3pt 误差
    sorted_sizes = sorted(FONT_SIZE_MAP.items(), key=lambda x: -x[1])
    for name, size in sorted_sizes:
        if abs(pt_val - size) <= 0.3:
            return name
    return f"{pt_val}pt"


def format_font_size(font_name, pt_val):
    """格式化为「宋体小四号」这样的可读格式"""
    size_name = pt_to_size_name(pt_val)
    result = ""
    if font_name:
        result = font_name
    if size_name:
        if "号" in size_name:
            result += size_name
        else:
            result += f"{size_name}号"
    elif pt_val:
        result += f"{pt_val}pt"
    return result or f"{pt_val}pt"


def format_expected(font_name, pt_val, bold=None):
    """格式化期望格式，如「宋体小四号加粗」"""
    size_name = pt_to_size_name(pt_val)
    result = ""
    if font_name:
        result = font_name
    if size_name:
        if "号" in size_name:
            result += size_name
        else:
            result += f"{size_name}号"
    elif pt_val:
        result += f"{pt_val}pt"
    if bold is True:
        result += "加粗"
    elif bold is False:
        result += "不加粗"
    return result


def _font_name_from_rpr(rPr, text=""):
    """从一层字符属性中解析适用于当前文本的字体。"""
    if rPr is not None:
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is not None:
            ea = rFonts.get(qn('w:eastAsia'))
            ascii_font = rFonts.get(qn('w:ascii'))
            normal_latin = {"Times New Roman", "Arial", "Calibri", "Cambria",
                            "Georgia", "Verdana", "Tahoma", "Trebuchet MS"}
            # Segoe UI 系列是 Word 的 emoji/符号字体，不应作为文档正文字体
            emoji_fonts = {"Segoe UI Emoji", "Segoe UI Symbol", "Segoe UI"}

            has_chinese = bool(re.search(r'[一-鿿]', text or ''))

            # 处理 eastAsia 或 ascii 为 emoji 字体的情况
            ea_is_emoji = ea in emoji_fonts
            ascii_is_emoji = ascii_font in emoji_fonts

            if ascii_is_emoji and ea_is_emoji:
                # 两个都是 emoji 字体，无法修复，报告
                return ascii_font or ea
            if ascii_is_emoji:
                # ascii 是 emoji，用 eastAsia
                return ea
            if ea_is_emoji:
                # eastAsia 是 emoji（Word 错误设置），用 ascii
                if ascii_font and ascii_font in normal_latin:
                    return ascii_font
                # ascii 也不是正常字体，报告 eastAsia
                return ea

            if ascii_font and ea and ascii_font != ea:
                if ascii_font in normal_latin:
                    # run 含中文 → 用 eastAsia；纯英文 → 用 ascii
                    return ea if has_chinese else ascii_font
                # ascii 异常，报告
                return ascii_font
            if ea:
                return ea
            if ascii_font:
                return ascii_font
    return None


def _style_chain(style):
    """按当前样式到基样式的顺序迭代，避免循环引用。"""
    seen = set()
    while style is not None and id(style) not in seen:
        seen.add(id(style))
        yield style
        style = getattr(style, "base_style", None)


def get_run_font(run):
    """解析直接格式和样式继承后的有效字体名。"""
    direct = _font_name_from_rpr(run._element.rPr, run.text)
    if direct or run.font.name:
        return direct or run.font.name

    paragraph = getattr(run, "_parent", None)
    styles = [getattr(run, "style", None), getattr(paragraph, "style", None)]
    for style in styles:
        for candidate in _style_chain(style):
            font_name = _font_name_from_rpr(candidate.element.rPr, run.text)
            if font_name or candidate.font.name:
                return font_name or candidate.font.name
    return None


def get_run_font_info(run):
    """解析直接格式和样式继承后的字体、字号与粗体。"""
    font_name = get_run_font(run)
    size = run.font.size
    bold = run.font.bold
    paragraph = getattr(run, "_parent", None)
    styles = [getattr(run, "style", None), getattr(paragraph, "style", None)]
    for style in styles:
        for candidate in _style_chain(style):
            if size is None and candidate.font.size is not None:
                size = candidate.font.size
            if bold is None and candidate.font.bold is not None:
                bold = candidate.font.bold
        if size is not None and bold is not None:
            break
    size_pt = emu_to_pt(size) if size else None
    return font_name, size_pt, bold


def get_paragraph_alignment(para):
    """获取段落对齐方式"""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    align = para.paragraph_format.alignment
    if align == WD_ALIGN_PARAGRAPH.CENTER:
        return "居中"
    elif align == WD_ALIGN_PARAGRAPH.LEFT:
        return "左对齐"
    elif align == WD_ALIGN_PARAGRAPH.RIGHT:
        return "右对齐"
    elif align == WD_ALIGN_PARAGRAPH.JUSTIFY:
        return "两端对齐"
    return None


def get_line_spacing(para):
    """获取行间距设置（返回 pt 值和类型）"""
    fmt = para.paragraph_format
    if fmt.line_spacing is not None:
        rule = fmt.line_spacing_rule
        raw = fmt.line_spacing
        # python-docx: line_spacing can be Length (Emu), Pt, or plain int (twips)
        if hasattr(raw, 'emu'):
            spacing_pt = round(raw.emu / 12700, 1)
        elif hasattr(raw, 'pt'):
            spacing_pt = round(raw.pt, 1)
        elif isinstance(raw, (int, float)):
            # plain int = twips (1/20 of a point)
            spacing_pt = round(raw / 20, 1)
        else:
            spacing_pt = None

        if rule is not None:
            from docx.enum.text import WD_LINE_SPACING
            if rule == WD_LINE_SPACING.EXACTLY:
                return spacing_pt, "固定值"
            elif rule == WD_LINE_SPACING.AT_LEAST:
                return spacing_pt, "最小值"
            elif rule == WD_LINE_SPACING.MULTIPLE:
                return spacing_pt, "多倍"
        return spacing_pt, "未知"
    return None, None


def load_rules_from_json(json_path):
    """
    从 JSON 规范文件加载规则，返回与硬编码 rules_dict 同格式的字典
    """
    with open(json_path, "r", encoding="utf-8") as f:
        raw = json.load(f)

    if isinstance(raw, dict):
        raw_rules = raw.get("rules", [])
        profile = raw.get("profile", {})
    else:
        raw_rules = raw
        profile = {}

    rules = {}
    meta = {
        "profile": profile,
        "enabled_checks": profile.get("enabled_checks"),
    }

    for r in raw_rules:
        section = r["section"]

        # 字体字号加粗规则
        if any(k in r for k in ("font", "size_pt", "bold")):
            rules[section] = {
                "font": r.get("font"),
                "size": r.get("size_pt"),
                "bold": r.get("bold"),
                "label": r.get("label", section),
            }

        # 对齐规则
        if r.get("alignment"):
            if section not in rules:
                rules[section] = {"label": r.get("label", section)}
            rules[section]["alignment"] = r["alignment"]

        # 行间距规则
        if r.get("line_spacing_pt"):
            rules.setdefault(section, {"label": r.get("label", section)}).update({
                "line_spacing_pt": r["line_spacing_pt"],
                "line_spacing_type": r.get("line_spacing_type", "固定值"),
                "label": r.get("label", section),
            })

        # 参考文献数量
        if section == "reference_count":
            meta["reference_min"] = r.get("min_count", 20)
        if section == "reference_foreign_count":
            meta["reference_foreign_min"] = r.get("min_count", 3)

        # 正文字数
        if section == "word_count":
            meta["word_count_design"] = r.get("min_words_design", 7000)
            meta["word_count_paper"] = r.get("min_words_paper", 10000)

        # 摘要字数
        if section == "abstract_word_count":
            meta["abstract_min_words"] = r.get("min_words", 400)

        # 页眉
        if section == "page_header":
            if r.get("header_text"):
                meta["header_text"] = r["header_text"]
            if r.get("size_pt"):
                meta["header_size_pt"] = r["size_pt"]
            if r.get("font"):
                meta["header_font"] = r["font"]

        # 脚注字号（供脚注检查使用，避免把论文默认值写死在逻辑里）
        if section == "footnote" and r.get("size_pt"):
            meta["footnote_size_pt"] = r["size_pt"]

        # 页面设置
        if section == "page_setup":
            meta["page_width_mm"] = r.get("width_mm", 210)
            meta["page_height_mm"] = r.get("height_mm", 297)

    return rules, meta


# 连续分类区（如声明）遇到这些特征即视为结束，避免状态粘住整个文档。
SECTION_STOP_KEYWORDS = ("摘要", "关键词", "目录", "致谢", "参考文献",
                         "附录", "abstract", "keywords")

# 声明区最长可延续的段落数，超过后强制按常规规则重新分类。
DECLARATION_MAX_PARAGRAPHS = 10


def normalize_text(text):
    """去掉排版用的空白，让「摘    要」「摘 要」都能可靠命中关键词。"""
    return re.sub(r'\s+', '', text or '')


def looks_like_section_start(text, bold=None, size_pt=None, align=None):
    """判断段落是否像新章节的开头，用于终止声明区这类连续分类。"""
    normalized = normalize_text(text).lower()
    if any(keyword in normalized for keyword in SECTION_STOP_KEYWORDS):
        return True
    # 编号标题：1 xxx / 1.1 xxx / 1.1.1 xxx
    if re.match(r'^\d+(\.\d+)*\s*\S', text or ''):
        return True
    # 图/表题注
    if re.match(r'^(图|表|Figure|Table)\s*\d', text or '', re.IGNORECASE):
        return True
    # 居中加粗的短文本，通常是章节标题
    if align == "居中" and bold and len(text or '') < 30:
        return True
    return False


def classify_paragraph(para, prev_class=None, declaration_streak=0):
    """
    尝试自动分类段落类型
    返回: 'cover_title', 'cover_field', 'abstract_zh_title', 'abstract_zh_body',
          'keywords_zh', 'abstract_en_title', 'abstract_en_body', 'keywords_en',
          'toc_title', 'toc_entry', 'level1_title', 'level2_title', 'level3_title',
          'body_text', 'figure_caption', 'table_caption', 'reference_title',
          'reference_entry', 'acknowledgment_title', 'acknowledgment_body',
          'appendix_title', 'declaration_title', 'declaration_body'
    """
    text = para.text.strip()
    if not text:
        return None
    normalized = normalize_text(text)

    # 获取主要字体信息
    fonts_sizes = []
    for r in para.runs:
        if r.text.strip():
            fn, sz, bold = get_run_font_info(r)
            fonts_sizes.append((fn, sz, bold))

    if not fonts_sizes:
        return None

    main_font = fonts_sizes[0][0]
    main_size = fonts_sizes[0][1]
    main_bold = fonts_sizes[0][2]
    align = get_paragraph_alignment(para)

    # 封面 - 检测封面标题（大字）
    if main_size and main_size >= 36:
        return "cover_title"
    # 封面 - 检测封面字段（题目、姓名、学号等），字号通常在 14-18pt 范围
    cover_keywords = ["题目", "学生姓名", "学号", "班级", "专业",
                      "学院", "指导教师", "完成日期"]
    if main_size and 14 <= main_size <= 18 and any(k in normalized for k in cover_keywords):
        return "cover_field"

    # 声明类（在封面兜底规则之前检测，避免误分类）
    if any(k in normalized for k in ("诚信声明书", "独创性声明", "使用授权声明")):
        return "declaration_title"
    if prev_class in ("declaration_title", "declaration_body"):
        # 命中新章节特征，或声明区已过长，都结束声明区，交回常规规则处理
        if (declaration_streak >= DECLARATION_MAX_PARAGRAPHS
                or looks_like_section_start(text, main_bold, main_size, align)):
            pass
        else:
            return "declaration_body"

    # 封面 - 内容行（接在封面字段后面的值，如 "张三"、"2022级"）
    if prev_class == "cover_field" and main_size and 14 <= main_size <= 18:
        return "cover_field"

    # 摘要标题
    if text.strip() == "摘    要" or text.strip() == "摘要":
        return "abstract_zh_title"
    if re.match(r'^Abstract', text, re.IGNORECASE):
        return "abstract_en_title"

    # 关键词
    if text.startswith("关键词"):
        return "keywords_zh"
    if re.match(r'^Key\s*words', text, re.IGNORECASE):
        return "keywords_en"

    # 目录
    if "目    录" in text or text.strip() == "目录":
        return "toc_title"
    # 目录条目：紧跟目录标题，文本末尾带页码（如 "1 绪论	1"、"参考文献	19"）
    if prev_class in ("toc_title", "toc_entry"):
        if re.search(r'[\t ]+\d+\s*$', text):
            return "toc_entry"

    # 参考文献标题
    if text == "参考文献":
        return "reference_title"

    # 参考文献条目：紧跟在参考文献标题之后、字号为五号(10.5pt)、不是致谢
    if prev_class in ("reference_title", "reference_entry"):
        if main_size and 9.5 <= main_size <= 12:
            # 致谢标题不是文献条目，其余短文本视为条目
            if not re.match(r'^致\s*谢', text) or len(text) > 10:
                return "reference_entry"

    # 也匹配带 [序号] 的格式
    if re.match(r'^\[\d+\]', text):
        return "reference_entry"

    # 致谢
    if re.match(r'^致\s*谢', text) and len(text) < 10:
        return "acknowledgment_title"

    # 附录
    if text.startswith("附录"):
        return "appendix_title"

    # 标题层级检测 (1 xxx, 1.1 xxx, 1.1.1 xxx，允许无空格如"1绪论")
    if main_bold:
        if re.match(r'^\d+\s*\S', text) and main_size and main_size >= 14.5:
            return "level1_title"
        if re.match(r'^\d+\.\d+\s*\S', text) and main_size and main_size >= 13.5:
            return "level2_title"
        if re.match(r'^\d+\.\d+\.\d+\s*\S', text) and main_size and main_size >= 11.5:
            return "level3_title"

    # 无编号标题：加粗 + 字号符合标题特征 + 不是正文长度（短文本更可能是标题）
    if main_bold and len(text) < 30:
        if main_size and 14.5 <= main_size <= 16:
            return "level1_title"
        if main_size and 13.5 <= main_size < 14.5:
            return "level2_title"
        # 12pt 加粗短文本不自动分类为标题，避免误报（正文中加粗关键词太多）

    # 图表标题
    if re.match(r'^(图|Figure)\s*\d+', text, re.IGNORECASE):
        return "figure_caption"
    if re.match(r'^(表|Table)\s*\d+', text, re.IGNORECASE):
        return "table_caption"

    # 摘要正文（在摘要标题和关键词之间）
    if prev_class in ("abstract_zh_title", "abstract_zh_body"):
        if main_font and "楷体" in (main_font or ""):
            return "abstract_zh_body"

    # 英文摘要正文
    if prev_class in ("abstract_en_title", "abstract_en_body"):
        if not text.startswith("Key") and not text.startswith("Abstract"):
            return "abstract_en_body"

    # 致谢正文
    if prev_class in ("acknowledgment_title", "acknowledgment_body"):
        if main_size and 11 <= main_size <= 13:
            return "acknowledgment_body"

    # 正文
    if main_size and 11 <= main_size <= 13:
        return "body_text"

    return "body_text"


def check_font(paragraph_class, para, issues, para_index, rules=None):
    """检测字体是否符合规范"""
    location = f"第 {para_index + 1} 段"

    # 收集所有 run 的字体信息
    run_infos = []
    for r in para.runs:
        if r.text.strip():
            fn, sz, bold = get_run_font_info(r)
            run_infos.append({"font": fn, "size": sz, "bold": bold, "text": r.text[:50]})

    if not run_infos:
        return

    # 截取段落原文
    para_text = para.text.strip()[:80]

    # 定义各部分的规范（默认规则）
    default_rules = {
        "cover_title": {"font": "宋体", "size": 16, "bold": True, "label": "封面题目"},
        "cover_field": {"font": "宋体", "size": 16, "bold": None, "label": "封面信息"},
        "abstract_zh_title": {"font": "黑体", "size": 15, "bold": None, "label": "中文摘要标题"},
        "abstract_zh_body": {"font": "楷体", "size": 12, "bold": None, "label": "中文摘要正文"},
        "keywords_zh": {"font": "黑体", "size": 12, "bold": None, "label": "中文关键词"},
        "abstract_en_title": {"font": "Arial Black", "size": 15, "bold": None, "label": "英文摘要标题"},
        "abstract_en_body": {"font": "Times New Roman", "size": 12, "bold": None, "label": "英文摘要正文"},
        "keywords_en": {"font": "Arial Black", "size": 12, "bold": None, "label": "英文关键词"},
        "level1_title": {"font": "宋体", "size": 15, "bold": True, "label": "一级标题"},
        "level2_title": {"font": "宋体", "size": 14, "bold": True, "label": "二级标题"},
        "level3_title": {"font": "宋体", "size": 12, "bold": True, "label": "三级标题"},
        "body_text": {"font": "宋体", "size": 12, "bold": False, "label": "正文"},
        "figure_caption": {"font": "宋体", "size": 10.5, "bold": False, "label": "图题注"},
        "table_caption": {"font": "宋体", "size": 10.5, "bold": False, "label": "表题注"},
        "reference_title": {"font": "宋体", "size": 15, "bold": True, "label": "参考文献标题"},
        "reference_entry": {"font": "宋体", "size": 10.5, "bold": False, "label": "参考文献条目"},
        "acknowledgment_title": {"font": "宋体", "size": 15, "bold": True, "label": "致谢标题"},
        "acknowledgment_body": {"font": "宋体", "size": 12, "bold": False, "label": "致谢正文"},
        "appendix_title": {"font": "宋体", "size": 15, "bold": True, "label": "附录标题"},
    }
    # 如果传入了外部规则，覆盖默认
    check_rules = default_rules if rules is None else rules

    if paragraph_class not in check_rules:
        return

    rule = check_rules[paragraph_class]
    label = rule["label"]
    expected_font = rule.get("font")
    expected_size = rule.get("size")
    expected_bold = rule.get("bold")

    # 封面题目和声明书为模板固定内容，不检测格式
    if paragraph_class in ("cover_title", "declaration_title", "declaration_body"):
        return

    # 关键词部分：「关键词：」为黑体，后面的具体关键词为楷体，混合字体，跳过整体字体检测
    if paragraph_class == "keywords_zh":
        return

    # 计算每个 run 在段落中的字符起始位置
    para_full_text = para.text
    char_pos = 0
    for info in run_infos:
        info["char_pos"] = char_pos
        char_pos += len(info["text"])

    for info in run_infos:
        pos_label = f"第{info['char_pos'] + 1}字起" if info['char_pos'] > 0 else ""
        loc = f"{location} ({label})" + (f" {pos_label}" if pos_label else "")

        # 字体检测
        # 已尝试解析直接格式和样式继承；仍为 None 时不能可靠判错。
        if expected_font and info["font"]:
            font_match = expected_font in (info["font"] or "")
            if expected_font == "Arial Black":
                font_match = "Arial" in (info["font"] or "") and "Black" in (info["font"] or "")
            if expected_font == "Times New Roman":
                font_match = "Times" in (info["font"] or "")

            if not font_match:
                # 英文摘要部分使用宋体：在中国高校论文中是常见且可接受的做法，跳过不报
                en_classes = {"abstract_en_title", "abstract_en_body", "keywords_en"}
                if paragraph_class in en_classes and info["font"] == "宋体":
                    continue

                # 宋体变体（汉仪宋体、方正宋体等）视为宋体
                song_ti_variants = {"HYShuSongErKW", "FZShuSong", "Founder SS", "SimSun", "NSimSun"}
                if expected_font == "宋体" and any(v in (info["font"] or "") for v in song_ti_variants):
                    continue

                # 参考文献、图/表题注中的英文内容使用 Times New Roman 是正常的
                mixed_classes = {"reference_entry", "figure_caption", "table_caption"}
                if paragraph_class in mixed_classes and "Times" in (info["font"] or ""):
                    # 该 run 主要是英文内容，Times New Roman 可接受
                    has_chinese = bool(re.search(r'[一-鿿]', info["text"] or ""))
                    if not has_chinese:
                        continue

                expected_desc = format_expected(expected_font, expected_size)
                issues.append(Issue(
                    location=loc,
                    severity=Severity.ERROR,
                    category="字体",
                    description=f"字体应为「{expected_desc}」，实际为「{info['font']}」",
                    expected=expected_desc,
                    actual=info["font"] or "未设置",
                    content=info["text"]
                ))

        # 字号检测（允许 0.5pt 误差）
        if expected_size and info["size"]:
            if abs(info["size"] - expected_size) > 0.5:
                expected_desc = format_expected(expected_font, expected_size)
                actual_desc = format_font_size(info["font"], info["size"])
                issues.append(Issue(
                    location=loc,
                    severity=Severity.ERROR,
                    category="字号",
                    description=f"字号应为「{expected_desc}」，实际为「{actual_desc}」",
                    expected=expected_desc,
                    actual=actual_desc,
                    content=info["text"]
                ))

        # 加粗检测（取所有 run 的一致结果）
        if expected_bold is not None:
            # 正文分类的大号粗体文字可能是误分类的标题（如目录后的文档大标题），跳过加粗检查
            if paragraph_class == "body_text" and run_infos[0]["size"] and run_infos[0]["size"] >= 14:
                break
            # 如果任何 run 不符合加粗要求就报
            all_bold = all(i["bold"] for i in run_infos)
            none_bold = all(not i["bold"] for i in run_infos)
            if expected_bold and not all_bold:
                issues.append(Issue(
                    location=f"{location} ({label})",
                    severity=Severity.WARNING,
                    category="加粗",
                    description=f"{label}应加粗",
                    expected="加粗",
                    actual="未加粗" if none_bold else "部分加粗",
                    content=para_text
                ))
                break  # 加粗是段落级属性，只报一次
            elif not expected_bold and not none_bold:
                issues.append(Issue(
                    location=f"{location} ({label})",
                    severity=Severity.WARNING,
                    category="加粗",
                    description=f"{label}不应加粗",
                    expected="不加粗",
                    actual="加粗",
                    content=para_text
                ))
                break  # 加粗是段落级属性，只报一次


def check_alignment(paragraph_class, para, issues, para_index, rules=None):
    """检测对齐方式"""
    location = f"第 {para_index + 1} 段"
    align = get_paragraph_alignment(para)

    # 封面字段通常用 Tab/空格手动排版，段落对齐属性不可靠，跳过
    if paragraph_class in ("cover_title", "cover_field"):
        return

    default_alignments = {
        "abstract_zh_title": "居中",
        "abstract_en_title": "居中",
        "toc_title": "居中",
        "figure_caption": "居中",
        "reference_title": "居中",
        "acknowledgment_title": "居中",
        "appendix_title": "居中",
    }

    if rules is None:
        expected_alignment = default_alignments.get(paragraph_class)
    else:
        expected_alignment = (rules.get(paragraph_class) or {}).get("alignment")

    if expected_alignment and align and align != expected_alignment:
        label_map = {
            "abstract_zh_title": "中文摘要标题",
            "abstract_en_title": "英文摘要标题",
            "toc_title": "目录标题",
            "figure_caption": "图题注",
            "reference_title": "参考文献标题",
            "acknowledgment_title": "致谢标题",
        }
        issues.append(Issue(
            location=f"{location} ({label_map.get(paragraph_class, paragraph_class)})",
            severity=Severity.ERROR,
            category="对齐",
            description=f"对齐方式应为「{expected_alignment}」，实际为「{align or '未设置'}」",
            expected=expected_alignment,
            actual=align or "未设置",
            content=para.text.strip()[:80]
        ))


def check_line_spacing(paragraph_class, para, issues, para_index, rules=None):
    """检测行间距"""
    location = f"第 {para_index + 1} 段"
    spacing_pt, spacing_type = get_line_spacing(para)

    if spacing_pt is None:
        return

    para_text = para.text.strip()[:80]

    # 正文和标题
    body_classes = {"body_text", "level1_title", "level2_title", "level3_title",
                    "abstract_zh_body", "abstract_en_body", "keywords_zh", "keywords_en",
                    "acknowledgment_body"}

    if paragraph_class in body_classes:
        # 外部模板只检查明确写入的行距，不继承论文默认值。
        expected_spacing, expected_type = None, None
        if rules is None:
            expected_spacing, expected_type = 22, "固定值"
        else:
            ls_rule = (rules.get(paragraph_class)
                       or rules.get("line_spacing_body") or {})
            if ls_rule.get("line_spacing_pt"):
                expected_spacing = ls_rule["line_spacing_pt"]
                expected_type = ls_rule.get("line_spacing_type", "固定值")

        if expected_spacing:
            if spacing_type != expected_type:
                issues.append(Issue(
                    location=f"{location} (行间距)",
                    severity=Severity.WARNING,
                    category="行间距",
                    description=f"行间距类型应为「{expected_type}」，实际为「{spacing_type}」",
                    expected=f"{expected_type} {expected_spacing}磅",
                    actual=f"{spacing_type} {spacing_pt}磅",
                    content=para_text,
                ))
            elif abs(spacing_pt - expected_spacing) > 1:
                issues.append(Issue(
                    location=f"{location} (行间距)",
                    severity=Severity.ERROR,
                    category="行间距",
                    description=(f"行间距应为「{expected_type} {expected_spacing}磅」，"
                                 f"实际为「{spacing_pt}磅」"),
                    expected=f"{expected_type} {expected_spacing}磅",
                    actual=f"{spacing_type} {spacing_pt}磅",
                    content=para_text,
                ))
        return

    # 参考文献条目
    if paragraph_class == "reference_entry":
        ref_spacing, ref_type = None, None
        if rules is None:
            ref_spacing, ref_type = 18, "固定值"
        else:
            ref_rule = rules.get("reference_entry") or {}
            if ref_rule.get("line_spacing_pt"):
                ref_spacing = ref_rule["line_spacing_pt"]
                ref_type = ref_rule.get("line_spacing_type", "固定值")

        if ref_spacing and spacing_type == ref_type and abs(spacing_pt - ref_spacing) > 1:
            issues.append(Issue(
                location=f"{location} (行间距)",
                severity=Severity.ERROR,
                category="行间距",
                description=(f"参考文献行间距应为「{ref_type} {ref_spacing}磅」，"
                             f"实际为「{spacing_pt}磅」"),
                expected=f"{ref_type} {ref_spacing}磅",
                actual=f"{spacing_type} {spacing_pt}磅",
                content=para_text,
            ))


def check_structure(paragraphs, classified, issues):
    """检测文档结构是否完整"""
    required_sections = [
        ("abstract_zh_title", "中文摘要"),
        ("abstract_en_title", "英文摘要"),
        ("keywords_zh", "中文关键词"),
        ("level1_title", "一级标题（正文）"),
        ("reference_title", "参考文献"),
    ]

    found = set()
    for cls in classified:
        if cls:
            found.add(cls)

    for cls_key, label in required_sections:
        if cls_key not in found:
            issues.append(Issue(
                location="文档结构",
                severity=Severity.ERROR,
                category="结构",
                description=f"缺少必要部分：{label}",
                expected=f"应包含「{label}」",
                actual="未找到"
            ))

    # 检查标题层级是否从 1 开始
    level1_count = sum(1 for c in classified if c == "level1_title")
    if level1_count == 0:
        issues.append(Issue(
            location="文档结构",
            severity=Severity.WARNING,
            category="结构",
            description="未检测到一级标题（格式：1 ×××）",
            expected="应有编号为 1, 2, 3... 的一级标题",
            actual="未找到"
        ))


def check_title_numbering(paragraphs, classified, issues):
    """检测标题编号是否规范"""
    level1_pattern = re.compile(r'^(\d+)\s+\S')
    level2_pattern = re.compile(r'^(\d+)\.(\d+)\s*\S')
    level3_pattern = re.compile(r'^(\d+)\.(\d+)\.(\d+)\s*\S')

    expected_l1 = 1
    expected_l2 = {}
    expected_l3 = {}

    for i, (para, cls) in enumerate(zip(paragraphs, classified)):
        text = para.text.strip()
        location = f"第 {i + 1} 段"

        if cls == "level1_title":
            m = level1_pattern.match(text)
            if m:
                num = int(m.group(1))
                if num != expected_l1:
                    issues.append(Issue(
                        location=f"{location} (一级标题)",
                        severity=Severity.WARNING,
                        category="编号",
                        description=f"一级标题编号应为 {expected_l1}，实际为 {num}",
                        expected=str(expected_l1),
                        actual=str(num)
                    ))
                expected_l1 = num + 1
                expected_l2[num] = 1

        elif cls == "level2_title":
            m = level2_pattern.match(text)
            if m:
                l1 = int(m.group(1))
                l2 = int(m.group(2))
                exp = expected_l2.get(l1, 1)
                if l2 != exp:
                    issues.append(Issue(
                        location=f"{location} (二级标题)",
                        severity=Severity.WARNING,
                        category="编号",
                        description=f"二级标题编号应为 {l1}.{exp}，实际为 {l1}.{l2}",
                        expected=f"{l1}.{exp}",
                        actual=f"{l1}.{l2}"
                    ))
                expected_l2[l1] = l2 + 1
                expected_l3[f"{l1}.{l2}"] = 1

        elif cls == "level3_title":
            m = level3_pattern.match(text)
            if m:
                l1 = int(m.group(1))
                l2 = int(m.group(2))
                l3 = int(m.group(3))
                key = f"{l1}.{l2}"
                exp = expected_l3.get(key, 1)
                if l3 != exp:
                    issues.append(Issue(
                        location=f"{location} (三级标题)",
                        severity=Severity.INFO,
                        category="编号",
                        description=f"三级标题编号应为 {l1}.{l2}.{exp}，实际为 {l1}.{l2}.{l3}",
                        expected=f"{l1}.{l2}.{exp}",
                        actual=f"{l1}.{l2}.{l3}"
                    ))
                expected_l3[key] = l3 + 1


def check_figure_table_numbering(paragraphs, classified, issues):
    """检测图表编号，含子图编号（a）（b）"""
    fig_pattern = re.compile(r'(图|Figure)\s*(\d+)[\-.](\d+)', re.IGNORECASE)
    # 子图：编号后有标题，末尾带（a）（b），如 "图2-1 短视频封面（a）"
    subfig_pattern = re.compile(r'(图|Figure)\s*(\d+)[\-.](\d+)\s+.+[（(]\s*([a-z]+)\s*[）)]\s*$', re.IGNORECASE)
    table_pattern = re.compile(r'(表|Table)\s*(\d+)[\-.](\d+)', re.IGNORECASE)

    # 每章下一个期望的主序号；子图与主图共用同一主编号
    expected_fig = {}
    expected_table = {}
    seen_fig = set()
    seen_table = set()
    # 子图编号收集：key="ch-num", value=[(位置, 子编号字母, 原文)]
    subfigures = {}

    for i, (para, cls) in enumerate(zip(paragraphs, classified)):
        text = para.text.strip()
        location = f"第 {i + 1} 段"

        if cls == "figure_caption":
            # 先尝试匹配子图格式：标题末尾带（a）（b）
            sm = subfig_pattern.search(text)
            m = sm or fig_pattern.search(text)
            if not m:
                issues.append(Issue(
                    location=f"{location} (图题注)",
                    severity=Severity.WARNING,
                    category="编号",
                    description="图编号格式应为「图X-Y 标题」，子图格式为「图X-Y 标题（a）」「图X-Y 标题（b）」",
                    expected="图X-Y 或 图X-Y 标题（a）",
                    actual=text[:30],
                    content=text[:80],
                ))
                continue

            ch = int(m.group(2))
            num = int(m.group(3))
            key = f"{ch}-{num}"

            if sm:
                subfigures.setdefault(key, []).append((i, m.group(4).lower(), text[:60]))

            # 同一个主编号只占位一次，重复的子图不再推进序号
            if key not in seen_fig:
                expected = expected_fig.setdefault(ch, 1)
                if num != expected:
                    issues.append(Issue(
                        location=f"{location} (图题注)",
                        severity=Severity.ERROR,
                        category="编号",
                        description=f"图编号不连续：应为「图{ch}-{expected}」，实际为「图{ch}-{num}」",
                        expected=f"图{ch}-{expected}",
                        actual=f"图{ch}-{num}",
                        content=text[:80],
                    ))
                expected_fig[ch] = num + 1
                seen_fig.add(key)

        elif cls == "table_caption":
            m = table_pattern.search(text)
            if not m:
                issues.append(Issue(
                    location=f"{location} (表题注)",
                    severity=Severity.WARNING,
                    category="编号",
                    description="表编号格式应为「表X-Y」（X 为章号，Y 为表序号）",
                    expected="表X-Y",
                    actual=text[:30],
                    content=text[:80],
                ))
                continue

            ch = int(m.group(2))
            num = int(m.group(3))
            key = f"{ch}-{num}"
            if key not in seen_table:
                expected = expected_table.setdefault(ch, 1)
                if num != expected:
                    issues.append(Issue(
                        location=f"{location} (表题注)",
                        severity=Severity.ERROR,
                        category="编号",
                        description=f"表编号不连续：应为「表{ch}-{expected}」，实际为「表{ch}-{num}」",
                        expected=f"表{ch}-{expected}",
                        actual=f"表{ch}-{num}",
                        content=text[:80],
                    ))
                expected_table[ch] = num + 1
                seen_table.add(key)

    # 子图编号连续性检查
    for fig_key, subs in subfigures.items():
        if len(subs) < 2:
            issues.append(Issue(
                location=f"图 {fig_key}",
                severity=Severity.WARNING,
                category="编号",
                description=f"子图编号「{fig_key}」只有 {len(subs)} 个子图，至少需要 2 个才有必要使用子编号（a）（b）",
                expected=f"图{fig_key} 标题（a）+ 图{fig_key} 标题（b）",
                actual=f"图{fig_key} 标题（{subs[0][1]}）",
                content=subs[0][2]
            ))
            continue
        # 验证从 (a) 开始连续递增
        labels = [s[1] for s in subs]
        expected_labels = [chr(ord('a') + j) for j in range(len(labels))]
        for j, (actual_label, (para_idx, _, orig_text)) in enumerate(zip(labels, subs)):
            if actual_label != expected_labels[j]:
                issues.append(Issue(
                    location=f"第 {para_idx + 1} 段 (图 {fig_key} 子图)",
                    severity=Severity.ERROR,
                    category="编号",
                    description=f"子图编号应为「（{expected_labels[j]}）」，实际为「（{actual_label}）」",
                    expected=f"图{fig_key} 标题（{expected_labels[j]}）",
                    actual=f"图{fig_key} 标题（{actual_label}）",
                    content=orig_text
                ))


def check_reference_format(paragraphs, classified, issues, meta=None):
    """检测参考文献格式（GB/T 7714）"""
    ref_count = 0
    foreign_count = 0
    bracket_count = 0
    non_bracket_count = 0

    in_refs = False
    for i, (para, cls) in enumerate(zip(paragraphs, classified)):
        if cls == "reference_title":
            in_refs = True
            continue

        if in_refs and cls == "reference_entry":
            ref_count += 1
            text = para.text.strip()

            # 检查是否为外文文献
            has_chinese = bool(re.search(r'[一-鿿]', text))
            if not has_chinese:
                foreign_count += 1

            # 检查格式：是否有 [序号]
            if re.match(r'^\[\d+\]', text):
                bracket_count += 1
            else:
                non_bracket_count += 1

    # 格式一致性检查
    if bracket_count > 0 and non_bracket_count > 0:
        issues.append(Issue(
            location="参考文献",
            severity=Severity.WARNING,
            category="参考文献",
            description=f"参考文献编号格式不统一：{bracket_count} 条使用 [序号]，{non_bracket_count} 条未使用",
            expected="统一使用 [序号] 或统一不用",
            actual=f"混用（{bracket_count} 条带 [序号]，{non_bracket_count} 条不带）"
        ))

    min_refs = (meta or {}).get("reference_min", 20)
    min_foreign = (meta or {}).get("reference_foreign_min", 3)

    if ref_count < min_refs:
        issues.append(Issue(
            location="参考文献",
            severity=Severity.ERROR,
            category="参考文献",
            description=f"参考文献数量不足，要求不少于 {min_refs} 篇，实际 {ref_count} 篇",
            expected=f"≥{min_refs} 篇",
            actual=f"{ref_count} 篇"
        ))

    if foreign_count < min_foreign:
        issues.append(Issue(
            location="参考文献",
            severity=Severity.ERROR,
            category="参考文献",
            description=f"外文文献不足，要求不少于 {min_foreign} 篇，实际 {foreign_count} 篇",
            expected=f"≥{min_foreign} 篇",
            actual=f"{foreign_count} 篇"
        ))


def check_reference_entry_format(paragraphs, classified, issues):
    """逐条检查参考文献条目格式（GB/T 7714）"""
    # GB/T 7714 文献类型标识
    # [M] 专著  [J] 期刊  [D] 学位论文  [C] 论文集  [N] 报纸  [EB/OL] 网络文献
    type_markers = re.compile(r'\[[A-Z](/[A-Z]+)*\]')

    in_refs = False
    ref_index = 0
    for i, (para, cls) in enumerate(zip(paragraphs, classified)):
        if cls == "reference_title":
            in_refs = True
            continue

        if in_refs and cls == "reference_entry":
            ref_index += 1
            text = para.text.strip()
            location = f"参考文献 [{ref_index}]"
            has_chinese = bool(re.search(r'[一-鿿]', text))

            # 去掉序号前缀
            content = re.sub(r'^\[\d+\]\s*', '', text)

            # 检查 1：是否有文献类型标识 [M] [J] [D] [EB/OL] 等
            if not type_markers.search(content):
                issues.append(Issue(
                    location=location,
                    severity=Severity.WARNING,
                    category="参考文献",
                    description="缺少文献类型标识，如 [M]（专著）、[J]（期刊）、[EB/OL]（网络文献）等",
                    expected="示例：作者.题名[J].刊名,年,卷(期):页码.",
                    actual=content[:60],
                    content=text[:80]
                ))

            if has_chinese:
                # 中文文献基本检查
                # GB/T 7714 中文格式：作者.题名[类型].出版信息.
                # 应包含 "." 分隔符
                dot_count = content.count('.')
                if dot_count < 2:
                    issues.append(Issue(
                        location=location,
                        severity=Severity.INFO,
                        category="参考文献",
                        description="中文参考文献格式可能不完整，GB/T 7714 要求用「.」分隔作者、题名、出版信息等要素",
                        expected="作者.题名[J].刊名,年,卷(期):页码.",
                        actual=content[:60],
                        content=text[:80]
                    ))

                # 检查是否包含年份信息
                if not re.search(r'[\(（]\s*\d{4}\s*[\)）]|\b20\d{2}\b|\b19\d{2}\b', content):
                    issues.append(Issue(
                        location=location,
                        severity=Severity.INFO,
                        category="参考文献",
                        description="中文参考文献可能缺少出版年份",
                        expected="应包含年份信息，如 2024 或 (2024)",
                        actual=content[:60],
                        content=text[:80]
                    ))
            else:
                # 英文文献基本检查
                # GB/T 7714 英文格式：Author A, Author B. Title[J]. Journal, Year, Vol(No): Pages.
                # 应包含年份
                if not re.search(r'\b(19|20)\d{2}\b', content):
                    issues.append(Issue(
                        location=location,
                        severity=Severity.INFO,
                        category="参考文献",
                        description="英文参考文献可能缺少出版年份",
                        expected="应包含年份，如 2024",
                        actual=content[:60],
                        content=text[:80]
                    ))

        elif in_refs and cls and cls != "reference_entry":
            break


def check_page_setup(doc, issues, meta=None):
    """检测页面设置"""
    meta = meta or {}
    expected_w = meta.get("page_width_mm", 210)
    expected_h = meta.get("page_height_mm", 297)

    for i, section in enumerate(doc.sections):
        w_mm = round(section.page_width / 36000, 1)
        h_mm = round(section.page_height / 36000, 1)

        if abs(w_mm - expected_w) > 5 or abs(h_mm - expected_h) > 5:
            issues.append(Issue(
                location=f"页面设置 (第 {i+1} 节)",
                severity=Severity.ERROR,
                category="页面",
                description=f"纸张应为 {expected_w}×{expected_h}mm (A4)，实际为 {w_mm}×{h_mm}mm",
                expected=f"{expected_w}×{expected_h}mm (A4)",
                actual=f"{w_mm}×{h_mm}mm"
            ))


def check_header(doc, issues, meta=None, classified=None):
    """检测页眉（仅检查正文部分）。

    规则没有指定页眉文字时，不猜测内容，只保留「非正文节不应有页眉」
    这类结构判断。
    """
    meta = meta or {}
    expected = meta.get("header_text")
    expected_size = meta.get("header_size_pt")
    expected_font = meta.get("header_font") or "宋体"

    non_header_classes = {"reference_title", "reference_entry",
                          "acknowledgment_title", "acknowledgment_body",
                          "appendix_title"}

    # 找出每个节的起始段落索引
    section_starts = [0]
    for j, para in enumerate(doc.paragraphs):
        pPr = para._element.pPr
        if pPr is not None and pPr.find(qn('w:sectPr')) is not None:
            section_starts.append(j + 1)

    for i, section in enumerate(doc.sections):
        header = section.header
        if header.is_linked_to_previous:
            continue

        # 非正文节（参考文献/致谢/附录）不应有页眉
        if classified and i < len(section_starts):
            start_idx = section_starts[i]
            if start_idx < len(classified) and classified[start_idx] in non_header_classes:
                # 检查这些节是否有页眉内容，有则报错
                header_texts = [p.text.strip() for p in header.paragraphs if p.text.strip()]
                if header_texts:
                    section_label = {"reference_title": "参考文献", "reference_entry": "参考文献",
                                     "acknowledgment_title": "致谢", "acknowledgment_body": "致谢",
                                     "appendix_title": "附录"}.get(classified[start_idx], "非正文部分")
                    issues.append(Issue(
                        location=f"页眉 (第 {i+1} 节 - {section_label})",
                        severity=Severity.WARNING,
                        category="页眉",
                        description=f"{section_label}部分不应有页眉",
                        expected="无页眉",
                        actual=header_texts[0],
                        content=header_texts[0]
                    ))
                continue

        header_texts = [p.text.strip() for p in header.paragraphs if p.text.strip()]

        # 规则未指定页眉文字时，不做内容比对
        if not expected:
            continue

        found = any(ht == expected for ht in header_texts)
        if not found:
            actual_text = header_texts[0] if header_texts else "（空）"
            issues.append(Issue(
                location=f"页眉 (第 {i+1} 节)",
                severity=Severity.ERROR,
                category="页眉",
                description=f"页眉内容应为「{expected}」",
                expected=expected,
                actual=actual_text,
                content=actual_text
            ))
            continue

        # 找到规定文字后，再核对页眉字号（规则未指定字号时不检查）
        if not expected_size:
            continue

        for p in header.paragraphs:
            for r in p.runs:
                if not r.text.strip():
                    continue
                fn, sz, bold = get_run_font_info(r)
                if sz and abs(sz - expected_size) > 0.5:
                    actual_desc = format_font_size(fn, sz)
                    expected_desc = format_font_size(expected_font, expected_size)
                    issues.append(Issue(
                        location=f"页眉 (第 {i+1} 节)",
                        severity=Severity.WARNING,
                        category="页眉",
                        description=(f"页眉应为「{expected_desc}」，"
                                     f"实际为「{actual_desc}」"),
                        expected=expected_desc,
                        actual=actual_desc,
                        content=r.text.strip()[:80]
                    ))


def check_cover_completeness(paragraphs, classified, issues):
    """检测封面信息是否完整（指导教师、完成日期等字段后面是否有值）"""
    # 需要检查的封面字段关键词
    required_fields = {
        "指导教师": "指导教师姓名",
        "完成日期": "完成日期",
    }

    for i, (para, cls) in enumerate(zip(paragraphs, classified)):
        if cls != "cover_field":
            continue
        text = para.text.strip()
        for keyword, label in required_fields.items():
            if keyword in text:
                # 检查该段落中关键词后面是否有实际内容
                # 如 "指导教师       钟志炫" 中 "钟志炫" 是值
                after_keyword = text.split(keyword, 1)[1].strip()
                if not after_keyword:
                    # 也检查下一个段落是否是该字段的值（有时值在下一段）
                    has_value = False
                    if i + 1 < len(paragraphs) and classified[i + 1] == "cover_field":
                        next_text = paragraphs[i + 1].text.strip()
                        # 下一段不是另一个字段标签
                        if next_text and not any(k in next_text for k in required_fields.keys()):
                            has_value = True
                    if not has_value:
                        issues.append(Issue(
                            location=f"第 {i + 1} 段 (封面)",
                            severity=Severity.WARNING,
                            category="结构",
                            description=f"封面「{label}」未填写",
                            expected=f"应填写{label}",
                            actual="空白",
                            content=text
                        ))


def check_abstract_content(paragraphs, classified, issues, meta=None):
    """检测摘要字数"""
    meta = meta or {}
    min_words = meta.get("abstract_min_words", 400)
    min_en_words = meta.get("abstract_en_min_words", 150)

    # 中文摘要
    zh_abstract_text = []
    in_zh = False
    for para, cls in zip(paragraphs, classified):
        if cls == "abstract_zh_title":
            in_zh = True
            continue
        if cls == "keywords_zh":
            break
        if in_zh and cls == "abstract_zh_body":
            zh_abstract_text.append(para.text.strip())

    zh_text = "".join(zh_abstract_text)
    zh_count = len(re.sub(r'\s+', '', zh_text))
    if 0 < zh_count < min_words:
        issues.append(Issue(
            location="中文摘要",
            severity=Severity.ERROR,
            category="字数",
            description=f"中文摘要字数不足，要求不少于 {min_words} 字，实际约 {zh_count} 字",
            expected=f"≥{min_words} 字",
            actual=f"约 {zh_count} 字"
        ))

    # 英文摘要：按英文单词数统计，仅作提示，避免与中文摘要标准混用
    en_abstract_text = []
    in_en = False
    for para, cls in zip(paragraphs, classified):
        if cls == "abstract_en_title":
            in_en = True
            continue
        if cls == "keywords_en":
            break
        if in_en and cls == "abstract_en_body":
            en_abstract_text.append(para.text.strip())

    en_count = len(re.findall(r'[A-Za-z]+', " ".join(en_abstract_text)))
    if 0 < en_count < min_en_words:
        issues.append(Issue(
            location="英文摘要",
            severity=Severity.INFO,
            category="字数",
            description=(f"英文摘要偏短，约 {en_count} 个英文单词，"
                         f"通常与中文摘要内容对应"),
            expected=f"≥{min_en_words} 词（若模板另有要求，以模板为准）",
            actual=f"约 {en_count} 词"
        ))


def check_word_count(doc, paragraphs, classified, issues, meta=None):
    """检测正文字数"""
    meta = meta or {}
    min_design = meta.get("word_count_design", 7000)
    min_paper = meta.get("word_count_paper", 10000)

    body_text = []
    in_body = False
    for para, cls in zip(paragraphs, classified):
        if cls in ("level1_title", "level2_title", "level3_title", "body_text"):
            in_body = True
        if cls in ("reference_title", "acknowledgment_title"):
            break
        if in_body and para.text.strip():
            body_text.append(para.text.strip())

    total_text = "".join(body_text)
    char_count = len(re.sub(r'\s+', '', total_text))

    if 0 < char_count < min_design:
        issues.append(Issue(
            location="正文",
            severity=Severity.WARNING,
            category="字数",
            description=(f"正文字数偏少：当前约 {char_count} 字（不含图表），"
                         f"低于所选规则的最低要求。"),
            expected=(f"≥{min_design} 字（按所选规则）；"
                      f"若适用较长材料标准，要求 ≥{min_paper} 字"),
            actual=f"约 {char_count} 字"
        ))


def check_footnotes(doc, paragraphs, classified, issues, meta=None):
    """检测脚注：编号顺序、格式、与参考文献对应关系"""
    from lxml import etree

    # 提取正文中的脚注引用顺序
    body = doc.element.body
    ref_elements = body.findall('.//' + qn('w:footnoteReference'))
    ref_ids = []
    for ref in ref_elements:
        fn_id = ref.get(qn('w:id'))
        if fn_id and fn_id not in ('0', '-1'):
            ref_ids.append(fn_id)

    if not ref_ids:
        return

    # 提取脚注内容和字体信息
    fn_contents = {}
    fn_has_font_issue = False
    for rel in doc.part.rels.values():
        if 'footnote' in rel.reltype.lower():
            fn_part = rel.target_part
            fn_xml = etree.fromstring(fn_part.blob)
            for fn in fn_xml.findall(qn('w:footnote')):
                fn_id = fn.get(qn('w:id'))
                if fn_id in ('0', '-1'):
                    continue
                texts = fn.findall('.//' + qn('w:t'))
                content = ''.join(t.text or '' for t in texts).strip()
                fn_contents[fn_id] = content

                # 检查脚注字体格式
                for run in fn.findall(qn('w:r')):
                    rPr = run.find(qn('w:rPr'))
                    if rPr is None:
                        continue
                    rFonts = rPr.find(qn('w:rFonts'))
                    sz = rPr.find(qn('w:sz'))
                    sz_val = None
                    if sz is not None:
                        sz_half = sz.get(qn('w:val'))
                        if sz_half:
                            try:
                                sz_val = round(int(sz_half) / 2, 1)
                            except ValueError:
                                pass

                    # 提取字体
                    font_name = None
                    if rFonts is not None:
                        font_name = rFonts.get(qn('w:eastAsia')) or rFonts.get(qn('w:ascii'))

                    # 获取该 run 的文本
                    run_texts = run.findall(qn('w:t'))
                    run_content = ''.join(t.text or '' for t in run_texts).strip()
                    if not run_content:
                        continue

                    # 检查字号（取规则中的脚注字号，未指定则默认小五 9pt）
                    expected_fn_size = (meta or {}).get("footnote_size_pt") or 9
                    if sz_val and abs(sz_val - expected_fn_size) > 0.5:
                        if not fn_has_font_issue:
                            actual_desc = format_font_size(font_name, sz_val)
                            expected_desc = format_font_size("宋体", expected_fn_size)
                            issues.append(Issue(
                                location=f"脚注 [{int(fn_id) if fn_id.isdigit() else fn_id}]",
                                severity=Severity.WARNING,
                                category="脚注",
                                description=(f"脚注字号应为「{expected_desc}」，"
                                             f"实际为「{actual_desc}」"),
                                expected=expected_desc,
                                actual=actual_desc,
                                content=content[:60]
                            ))
                            fn_has_font_issue = True

                    # 检查字体：中文部分应为宋体，英文部分应为 Times New Roman
                    has_english = bool(re.search(r'[A-Za-z]', run_content))
                    if font_name:
                        if has_english and not re.search(r'[一-鿿]', run_content):
                            # 纯英文部分应为 Times New Roman
                            if 'Times' not in font_name and font_name != '宋体':
                                actual_desc = format_font_size(font_name, sz_val)
                                issues.append(Issue(
                                    location=f"脚注 [{int(fn_id)}]",
                                    severity=Severity.WARNING,
                                    category="脚注",
                                    description=f"脚注中英文字体应为「Times New Roman小五号」，实际为「{actual_desc}」",
                                    expected="Times New Roman小五号",
                                    actual=actual_desc,
                                    content=run_content[:40]
                                ))
                        elif not has_english:
                            # 中文部分应为宋体
                            if '宋体' not in (font_name or ''):
                                actual_desc = format_font_size(font_name, sz_val)
                                issues.append(Issue(
                                    location=f"脚注 [{int(fn_id)}]",
                                    severity=Severity.WARNING,
                                    category="脚注",
                                    description=f"脚注中文字体应为「宋体小五号」，实际为「{actual_desc}」",
                                    expected="宋体小五号",
                                    actual=actual_desc,
                                    content=run_content[:40]
                                ))
            break

    # 检查编号是否连续。
    # 注意：Word 的脚注 id 起始值并不固定（0、1 通常预留给分隔符），
    # 因此以第一处引用为基准检查是否依次递增，而不是硬性要求从 1 开始。
    ref_int_ids = []
    for rid in ref_ids:
        try:
            ref_int_ids.append(int(rid))
        except (TypeError, ValueError):
            ref_int_ids.append(None)

    base_id = None
    for j, (rid, int_id) in enumerate(zip(ref_ids, ref_int_ids)):
        location = f"脚注引用 (第{j + 1}处)"
        if int_id is None:
            issues.append(Issue(
                location=location,
                severity=Severity.WARNING,
                category="脚注",
                description=f"脚注编号格式异常，无法判断顺序：{rid}",
                expected="数字编号",
                actual=str(rid),
                content=fn_contents.get(rid, '')[:60],
            ))
            continue

        if base_id is None:
            base_id = int_id
            continue

        if int_id != base_id + j:
            issues.append(Issue(
                location=location,
                severity=Severity.ERROR,
                category="脚注",
                description=(f"脚注编号不连续：期望 [{j + 1}]，"
                             f"实际 [{int_id - base_id + 1}]"),
                expected=f"[{j + 1}]",
                actual=f"[{int_id - base_id + 1}]",
                content=fn_contents.get(rid, '')[:60],
            ))

    # 检查重复引用：同一文献多次引用时是否都给了新编号
    # （GB/T 7714 要求顺序编码制，每次引用按出现顺序编号）
    display_index = {}
    for j, rid in enumerate(ref_ids):
        display_index[rid] = j + 1

    seen_contents = {}
    for rid in ref_ids:
        fn_text = fn_contents.get(rid, '')
        if not fn_text:
            continue
        if fn_text in seen_contents:
            issues.append(Issue(
                location=f"脚注引用 (第{display_index[rid]}处)",
                severity=Severity.WARNING,
                category="脚注",
                description=(f"脚注内容与第 {display_index[seen_contents[fn_text]]} 处重复，"
                             f"同一文献应使用同一编号"),
                expected="重复引用使用相同编号",
                actual=f"给了新编号 [{display_index[rid]}]",
                content=fn_text[:60],
            ))
        else:
            seen_contents[fn_text] = rid


def check_document(filepath, rules_json_path=None):
    """
    主检测函数
    :param filepath: Word 文件路径
    :param rules_json_path: 可选，规范 JSON 文件路径。不传则用内置规则。
    """
    doc = Document(filepath)
    paragraphs = doc.paragraphs

    # 加载规则
    rules = None
    meta = {}
    if rules_json_path:
        rules, meta = load_rules_from_json(rules_json_path)

    enabled_checks = meta.get("enabled_checks")
    if enabled_checks is None:
        enabled_checks = DEFAULT_ENABLED_CHECKS
    else:
        enabled_checks = set(enabled_checks)

    # 分类所有段落
    classified = []
    prev = None
    declaration_streak = 0
    for para in paragraphs:
        cls = classify_paragraph(para, prev, declaration_streak)
        classified.append(cls)
        if cls:
            prev = cls
        if cls in ("declaration_title", "declaration_body"):
            declaration_streak += 1
        else:
            declaration_streak = 0

    result = CheckResult(filename=filepath.split("/")[-1])

    # 逐段检测
    for i, (para, cls) in enumerate(zip(paragraphs, classified)):
        if cls is None or not para.text.strip():
            continue
        if "font" in enabled_checks:
            check_font(cls, para, result.issues, i, rules)
        if "alignment" in enabled_checks:
            check_alignment(cls, para, result.issues, i, rules)
        if "line_spacing" in enabled_checks:
            check_line_spacing(cls, para, result.issues, i, rules)

    # 全局检测
    if "structure" in enabled_checks:
        check_structure(paragraphs, classified, result.issues)
    if "cover" in enabled_checks:
        check_cover_completeness(paragraphs, classified, result.issues)
    if "title_numbering" in enabled_checks:
        check_title_numbering(paragraphs, classified, result.issues)
    if "figure_table_numbering" in enabled_checks:
        check_figure_table_numbering(paragraphs, classified, result.issues)
    if "references" in enabled_checks:
        check_reference_format(paragraphs, classified, result.issues, meta)
        check_reference_entry_format(paragraphs, classified, result.issues)
    if "page_setup" in enabled_checks:
        check_page_setup(doc, result.issues, meta)
    if "header" in enabled_checks:
        check_header(doc, result.issues, meta, classified)
    if "abstract" in enabled_checks:
        check_abstract_content(paragraphs, classified, result.issues, meta)
    if "word_count" in enabled_checks:
        check_word_count(doc, paragraphs, classified, result.issues, meta)
    if "footnotes" in enabled_checks:
        check_footnotes(doc, paragraphs, classified, result.issues, meta)

    # 统计
    result.summary = {
        "total_paragraphs": len(paragraphs),
        "classified_paragraphs": sum(1 for c in classified if c),
        "errors": result.error_count,
        "warnings": result.warning_count,
        "infos": result.info_count,
    }

    return result
