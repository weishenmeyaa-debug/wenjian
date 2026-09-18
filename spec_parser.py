"""
规范文件解析模块
从 Word 格式规范中有限提取规则草稿，输出 JSON
"""

import re
import json
from docx import Document
from docx.oxml.ns import qn


# 中文字号 -> pt 对照表
SIZE_MAP = {
    "初号": 42, "小初": 36, "一号": 26, "小一": 24,
    "二号": 22, "小二": 18, "三号": 16, "小三": 15,
    "四号": 14, "小四": 12, "五号": 10.5, "小五": 9,
    "六号": 7.5, "小六": 6.5, "七号": 5.5, "八号": 5,
}


def parse_size(text):
    """从文本中提取字号"""
    # 匹配 "小三号" "三号" "小四号" "四号" "五号" 等
    for name, pt in SIZE_MAP.items():
        if name + "号" in text:
            return pt, name + "号"
    # 匹配 "5号" "小3号" 等变体
    m = re.search(r'小\s*(\d)\s*号', text)
    if m:
        n = int(m.group(1))
        name = f"小{'零一二三四五六七八九'[n]}号"
        return SIZE_MAP.get(name, None), name
    m = re.search(r'(\d)\s*号', text)
    if m:
        n = int(m.group(1))
        name = f"{'零一二三四五六七八九'[n]}号"
        return SIZE_MAP.get(name, None), name
    return None, None


def parse_font(text):
    """从文本中提取字体名"""
    fonts = []
    if "Arial Black" in text:
        fonts.append("Arial Black")
    elif "Times New Roman" in text or "Time New Roman" in text:
        fonts.append("Times New Roman")
    if "宋体" in text:
        fonts.append("宋体")
    if "黑体" in text:
        fonts.append("黑体")
    if "楷体" in text:
        fonts.append("楷体")
    return fonts


def parse_bold(text):
    """检查是否要求加粗"""
    return "加粗" in text


def parse_alignment(text):
    """检查对齐方式"""
    if "居中" in text:
        return "居中"
    if "两端对齐" in text:
        return "两端对齐"
    if "左对齐" in text:
        return "左对齐"
    return None


def parse_line_spacing(text):
    """提取行间距"""
    m = re.search(r'固定行[距间]?[距]?\s*(\d+)\s*pt', text, re.IGNORECASE)
    if m:
        return int(m.group(1)), "固定值"
    m = re.search(r'行间距?固定值\s*(\d+)\s*pt', text, re.IGNORECASE)
    if m:
        return int(m.group(1)), "固定值"
    return None, None


def parse_spec_docx(filepath):
    """
    解析规范文档，返回结构化规则列表
    """
    doc = Document(filepath)
    paragraphs = doc.paragraphs

    # ===== 第一部分：解析正文中的格式规则 =====
    prose_rules = []

    for i, p in enumerate(paragraphs):
        text = p.text.strip()
        if not text:
            continue

        # --- 封面规则 ---
        if "三号宋体加粗" in text and ("题目" in text or "中文题目" in text):
            prose_rules.append({
                "section": "cover_title",
                "label": "封面题目",
                "font": "宋体", "size_pt": 16, "size_name": "三号",
                "bold": True, "alignment": "居中",
                "source": f"P{i}"
            })
        if "三号宋体" in text and ("作者姓名" in text or "指导教师" in text or "日期" in text):
            prose_rules.append({
                "section": "cover_field",
                "label": "封面信息（姓名/指导/日期等）",
                "font": "宋体", "size_pt": 16, "size_name": "三号",
                "bold": None, "alignment": "居中",
                "source": f"P{i}"
            })

        # --- 摘要规则 ---
        if "摘要" in text and "小三号黑体" in text:
            prose_rules.append({
                "section": "abstract_zh_title",
                "label": "中文摘要标题",
                "font": "黑体", "size_pt": 15, "size_name": "小三号",
                "bold": None, "alignment": "居中",
                "source": f"P{i}"
            })
        if "关键词" in text and "小四号黑体" in text:
            prose_rules.append({
                "section": "keywords_zh",
                "label": "中文关键词",
                "font": "黑体", "size_pt": 12, "size_name": "小四号",
                "bold": None, "alignment": None,
                "source": f"P{i}"
            })
        if "楷体" in text and "小四号" in text and "摘要" in text:
            prose_rules.append({
                "section": "abstract_zh_body",
                "label": "中文摘要正文",
                "font": "楷体", "size_pt": 12, "size_name": "小四号",
                "bold": None, "alignment": None,
                "source": f"P{i}"
            })
        if "Arial Black" in text and "Abstract" in text and "小三" in text:
            prose_rules.append({
                "section": "abstract_en_title",
                "label": "英文摘要标题 Abstract",
                "font": "Arial Black", "size_pt": 15, "size_name": "小三号",
                "bold": None, "alignment": "居中",
                "source": f"P{i}"
            })
        if "Arial Black" in text and "Key words" in text and "小四" in text:
            prose_rules.append({
                "section": "keywords_en",
                "label": "英文关键词 Key words",
                "font": "Arial Black", "size_pt": 12, "size_name": "小四号",
                "bold": None, "alignment": None,
                "source": f"P{i}"
            })
        if "Times New Roman" in text and "小四" in text and "英文摘要" in text:
            prose_rules.append({
                "section": "abstract_en_body",
                "label": "英文摘要正文",
                "font": "Times New Roman", "size_pt": 12, "size_name": "小四号",
                "bold": None, "alignment": None,
                "source": f"P{i}"
            })

        # --- 目录规则 ---
        if "目录标题小三号黑体" in text:
            prose_rules.append({
                "section": "toc_title",
                "label": "目录标题",
                "font": "黑体", "size_pt": 15, "size_name": "小三号",
                "bold": None, "alignment": "居中",
                "source": f"P{i}"
            })

        # --- 正文标题规则 ---
        if re.match(r'一级标题.*小三号宋体', text.replace(" ", "")):
            prose_rules.append({
                "section": "level1_title",
                "label": "一级标题",
                "font": "宋体", "size_pt": 15, "size_name": "小三号",
                "bold": True, "alignment": None,
                "source": f"P{i}"
            })
        if re.match(r'二级标题.*四号宋体', text.replace(" ", "")):
            prose_rules.append({
                "section": "level2_title",
                "label": "二级标题",
                "font": "宋体", "size_pt": 14, "size_name": "四号",
                "bold": True, "alignment": None,
                "source": f"P{i}"
            })
        if re.match(r'三级标题.*小四号宋体', text.replace(" ", "")):
            prose_rules.append({
                "section": "level3_title",
                "label": "三级标题",
                "font": "宋体", "size_pt": 12, "size_name": "小四号",
                "bold": True, "alignment": None,
                "source": f"P{i}"
            })
        if re.match(r'正文.*小四号宋体', text.replace(" ", "")):
            prose_rules.append({
                "section": "body_text",
                "label": "正文",
                "font": "宋体", "size_pt": 12, "size_name": "小四号",
                "bold": False, "alignment": None,
                "source": f"P{i}"
            })

        # --- 行间距 ---
        if "固定行间距22pt" in text.replace(" ", ""):
            prose_rules.append({
                "section": "line_spacing_body",
                "label": "正文及标题行间距",
                "line_spacing_pt": 22,
                "line_spacing_type": "固定值",
                "source": f"P{i}"
            })

        # --- 图表 ---
        if "图正下方居中" in text and "宋体五号" in text:
            prose_rules.append({
                "section": "figure_caption",
                "label": "图题注",
                "font": "宋体", "size_pt": 10.5, "size_name": "五号",
                "bold": False, "alignment": "居中",
                "source": f"P{i}"
            })
        if "表正上方居中" in text and "宋体五号" in text:
            prose_rules.append({
                "section": "table_caption",
                "label": "表题注",
                "font": "宋体", "size_pt": 10.5, "size_name": "五号",
                "bold": False, "alignment": "居中",
                "source": f"P{i}"
            })

        # --- 参考文献 ---
        if "参考文献" in text and ("五号宋体" in text or "五号宋" in text):
            prose_rules.append({
                "section": "reference_title",
                "label": "参考文献标题",
                "font": "宋体", "size_pt": 15, "size_name": "小三号",
                "bold": True, "alignment": "居中",
                "source": f"P{i}"
            })
        if "参考文献" in text and "固定行距18pt" in text.replace(" ", ""):
            prose_rules.append({
                "section": "reference_entry",
                "label": "参考文献条目",
                "font": "宋体", "size_pt": 10.5, "size_name": "五号",
                "bold": False, "alignment": None,
                "line_spacing_pt": 18, "line_spacing_type": "固定值",
                "source": f"P{i}"
            })

        # --- 参考文献数量 ---
        if "不少于20篇" in text or "不少于 20 篇" in text:
            prose_rules.append({
                "section": "reference_count",
                "label": "参考文献数量",
                "min_count": 20,
                "source": f"P{i}"
            })
        if "外文文献" in text and ("不少于3篇" in text or "不少于 3 篇" in text):
            prose_rules.append({
                "section": "reference_foreign_count",
                "label": "外文文献数量",
                "min_count": 3,
                "source": f"P{i}"
            })

        # --- 脚注 ---
        if "脚注" in text and "小五号宋体" in text:
            prose_rules.append({
                "section": "footnote",
                "label": "脚注",
                "font": "宋体", "size_pt": 9, "size_name": "小五号",
                "bold": False, "alignment": None,
                "source": f"P{i}"
            })

        # --- 页码 ---
        if "页码位于页面底端居中" in text and "小五号" in text:
            prose_rules.append({
                "section": "page_number",
                "label": "页码",
                "font": None, "size_pt": 9, "size_name": "小五号",
                "alignment": "居中",
                "source": f"P{i}"
            })

        # --- 页眉 ---
        if "页眉字样为" in text:
            m = re.search(r'[「“”"](.*?)[」“”"]', text)
            header_val = m.group(1) if m else "XXX大学毕业设计（论文）"
            prose_rules.append({
                "section": "page_header",
                "label": "页眉",
                "font": "宋体", "size_pt": 9, "size_name": "小五号",
                "alignment": "居中",
                "header_text": header_val,
                "source": f"P{i}"
            })

        # --- 页面设置 ---
        if "A4" in text and "210" in text and "297" in text:
            prose_rules.append({
                "section": "page_setup",
                "label": "页面设置",
                "paper": "A4",
                "width_mm": 210,
                "height_mm": 297,
                "source": f"P{i}"
            })

        # --- 正文字数 ---
        if "设计说明文" in text and "7000" in text:
            prose_rules.append({
                "section": "word_count",
                "label": "正文字数",
                "min_words_design": 7000,
                "min_words_paper": 10000,
                "source": f"P{i}"
            })

        # --- 摘要字数 ---
        if "摘要" in text and "不少于400字" in text.replace(" ", ""):
            prose_rules.append({
                "section": "abstract_word_count",
                "label": "摘要字数",
                "min_words": 400,
                "source": f"P{i}"
            })

    # ===== 第二部分：解析末尾的注释段落（格式标注） =====
    # P738 开始有一组简洁的格式标注
    annotation_rules = []
    prev_section = None
    section_context = {}

    for i, p in enumerate(paragraphs):
        text = p.text.strip()
        if not text:
            continue

        # 根据前面的段落推断这些注释对应的文档部分
        # 查找附近的模板内容来确定上下文
        for j in range(max(0, i - 5), i):
            prev_text = paragraphs[j].text.strip()
            if "摘要" in prev_text and len(prev_text) < 20:
                section_context[i] = "abstract_area"
            elif "目录" in prev_text and len(prev_text) < 10:
                section_context[i] = "toc_area"
            elif "参考文献" in prev_text and len(prev_text) < 15:
                section_context[i] = "reference_area"
            elif "致" in prev_text and "谢" in prev_text and len(prev_text) < 10:
                section_context[i] = "ack_area"

    # ===== 合并规则 =====
    # 按 section 去重，保留最详细的
    seen = {}
    for rule in prose_rules:
        key = rule["section"]
        if key not in seen or len(str(rule)) > len(str(seen[key])):
            seen[key] = rule

    return list(seen.values())


def rules_to_json(rules, output_path):
    """将规则列表输出为 JSON 文件"""
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(rules, f, ensure_ascii=False, indent=2)


def load_rules(json_path):
    """从 JSON 文件加载规则"""
    with open(json_path, "r", encoding="utf-8") as f:
        return json.load(f)


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("用法: python spec_parser.py <规范文件路径> [输出JSON路径]")
        sys.exit(1)

    filepath = sys.argv[1]
    output = sys.argv[2] if len(sys.argv) > 2 else "format_rules.json"

    rules = parse_spec_docx(filepath)
    rules_to_json(rules, output)

    print(f"解析完成，共提取 {len(rules)} 条规则:")
    for r in rules:
        print(f"  [{r['section']}] {r['label']}")
        parts = []
        if r.get("font"): parts.append(f"字体={r['font']}")
        if r.get("size_name"): parts.append(f"字号={r['size_name']}")
        if r.get("bold") is not None: parts.append(f"加粗={'是' if r['bold'] else '否'}")
        if r.get("alignment"): parts.append(f"对齐={r['alignment']}")
        if r.get("line_spacing_pt"): parts.append(f"行距={r['line_spacing_pt']}pt")
        print(f"    {', '.join(parts)}")
    print(f"\n规则已保存到: {output}")
