"""格式检查器 - 本地 Word 文档格式检查应用。"""

import io
import os
import uuid
import json
import subprocess
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

from dataclasses import dataclass, field

from flask import Flask, render_template, request, redirect, url_for, send_file, flash, jsonify
from checker import check_document
from report import generate_report_html


@dataclass
class FailedResult:
    """单个文件解析失败时的占位结果，字段与 CheckResult 保持一致。"""
    filename: str
    error: str
    issues: list = field(default_factory=list)
    summary: dict = field(default_factory=lambda: {"errors": 0, "warnings": 0, "infos": 0})

app = Flask(__name__)
app.secret_key = os.urandom(24)
app.config["SHOW_FOOTER"] = True
app.config["MAX_CONTENT_LENGTH"] = 96 * 1024 * 1024
# 本地工具，改界面文案后刷新页面即可看到，无需重启。
app.config["TEMPLATES_AUTO_RELOAD"] = True
app.jinja_env.auto_reload = True

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
SPEC_DIR = os.path.join(BASE_DIR, "specs")
# 本机偏好（目前只记录主模板）。放在 BASE_DIR 而不是 specs/，
# 避免被模板扫描当成一套模板读进去。
SETTINGS_PATH = os.path.join(BASE_DIR, "settings.json")

ALLOWED_EXTENSIONS = {"docx", "doc"}
MAX_FILES_PER_CHECK = 12
REPORT_TTL_SECONDS = 60 * 60
BUILTIN_PROFILE_ID = "__builtin__"
# 默认只监听本机。要让同一网络的其他人访问，启动时显式加 --host 0.0.0.0。
DEFAULT_HOST = "127.0.0.1"

# 报告只在当前进程内短期保存，不写入磁盘。
REPORT_CACHE = {}

# 规则里的中文字体名 → 浏览器字族，让「格式参考」的样例显示成对的字体。
FONT_CSS = {
    "宋体": "'SimSun','Songti SC',serif",
    "黑体": "'SimHei','Heiti SC',sans-serif",
    "楷体": "'KaiTi','Kaiti SC',serif",
    "仿宋": "'FangSong','STFangsong',serif",
    "微软雅黑": "'Microsoft YaHei','PingFang SC',sans-serif",
    "等线": "'DengXian','PingFang SC',sans-serif",
}

# 「格式参考」页要告诉使用者这套模板会查哪些项。顺序按检查流程排，
# 不按字母序，否则读起来是乱的。
CHECK_ORDER = (
    "structure", "cover", "abstract", "word_count", "font", "alignment",
    "line_spacing", "title_numbering", "figure_table_numbering",
    "references", "footnotes", "page_setup", "header",
)
CHECK_LABELS = {
    "structure": "章节结构",
    "cover": "封面",
    "abstract": "摘要与关键词",
    "word_count": "字数",
    "font": "字体字号",
    "alignment": "对齐",
    "line_spacing": "行距",
    "title_numbering": "标题编号",
    "figure_table_numbering": "图表编号",
    "references": "参考文献",
    "footnotes": "脚注",
    "page_setup": "页面设置",
    "header": "页眉",
}

def purge_directory(directory):
    """清空临时目录里的历史残留，避免旧上传件长期留在本机。"""
    try:
        entries = os.listdir(directory)
    except OSError:
        return 0

    removed = 0
    for name in entries:
        path = os.path.join(directory, name)
        try:
            if os.path.isfile(path) or os.path.islink(path):
                os.remove(path)
                removed += 1
        except OSError:
            pass
    return removed


os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(SPEC_DIR, exist_ok=True)

# 上次异常退出可能留下副本，启动时先清一遍。
purge_directory(UPLOAD_DIR)


@app.context_processor
def inject_footer():
    return dict(show_footer=app.config.get("SHOW_FOOTER", False))


@app.errorhandler(413)
def request_entity_too_large(error):
    """上传超限时不暴露 Flask 默认错误页，回到首页给出可执行的提示。"""
    limit_mb = (app.config.get("MAX_CONTENT_LENGTH") or 0) // (1024 * 1024)
    flash(f"这一批文件合计超过 {limit_mb} MB，请减少文件数量或分批检查。", "error")
    return redirect(url_for("index"))


def _safe_display_name(filename):
    """仅保留客户端文件名的最后一段，用于页面显示。"""
    return os.path.basename((filename or "").replace("\\", "/")) or "未命名文档.docx"


def _extension(filename):
    name = _safe_display_name(filename)
    return name.rsplit(".", 1)[-1].lower() if "." in name else ""


def _profile_payload(path):
    """读取新旧两种规则格式，并补齐可展示的模板信息。"""
    with open(path, "r", encoding="utf-8") as f:
        raw = json.load(f)

    if isinstance(raw, dict):
        rules = raw.get("rules", [])
        profile = raw.get("profile", {})
    else:
        rules = raw
        profile = {}

    return {
        "profile": {
            "name": profile.get("name", "未命名模板"),
            "document_type": profile.get("document_type", "Word 文档"),
            "description": profile.get("description", "这套模板没有写说明"),
            "source": profile.get("source", "未标注来源"),
            "version": profile.get("version", "1.0"),
            "updated_at": profile.get("updated_at", "未标注"),
            "enabled_checks": profile.get("enabled_checks"),
        },
        "rules": rules,
    }


def load_settings():
    """读取本机偏好。文件缺失或损坏时当作空设置，不影响检查功能。"""
    try:
        with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def save_settings(settings):
    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except OSError:
        return False
    return True


def list_profiles():
    settings = load_settings()
    default_id = settings.get("default_profile_id") or ""
    profiles = []
    builtin_path = os.path.join(BASE_DIR, "format_rules.json")
    if os.path.exists(builtin_path):
        payload = _profile_payload(builtin_path)
        profiles.append({
            "id": BUILTIN_PROFILE_ID,
            "path": builtin_path,
            "is_builtin": True,
            **payload["profile"],
            "rule_count": len(payload["rules"]),
        })

    for filename in sorted(os.listdir(SPEC_DIR)):
        if not filename.endswith(".json"):
            continue
        path = os.path.join(SPEC_DIR, filename)
        try:
            payload = _profile_payload(path)
        except (OSError, ValueError, TypeError):
            continue
        profiles.append({
            "id": filename,
            "path": path,
            "is_builtin": False,
            **payload["profile"],
            "rule_count": len(payload["rules"]),
        })

    for profile in profiles:
        profile["is_default"] = profile["id"] == default_id
        # 卡片上那行说明。示例模板不参与「主模板」这套说法。
        if profile["is_builtin"]:
            profile["meta_text"] = f"示例模板 · {profile['rule_count']} 条规则"
        elif profile["is_default"]:
            profile["meta_text"] = f"主要格式模板 · {profile['rule_count']} 条已保存规则"
        else:
            profile["meta_text"] = f"格式模板 · {profile['rule_count']} 条已保存规则"

    # 主模板固定排第一，其余按名称排，保证每次打开顺序一致。
    profiles.sort(key=lambda p: (not p["is_default"], p["is_builtin"], p["name"]))
    return profiles


def resolve_profile(profile_id):
    """把外部传入的模板标识解析为已知模板，找不到就返回 None。

    不做「返回第一个」的兜底：任何无法识别的标识都必须显式失败，
    否则上层路由可能拿到内置模板继续处理，掩盖非法输入。
    """
    if not profile_id:
        return None
    for profile in list_profiles():
        if profile["id"] == profile_id:
            return profile
    return None


def get_profile(profile_id=None):
    """返回要展示的模板：未指定时用主模板，主模板缺失再退回内置示例。

    仅用于「选择默认值」这类展示场景；涉及文件删除、读取的路由
    必须改用 resolve_profile，避免对非法标识静默兜底。
    """
    profile = resolve_profile(profile_id)
    if profile:
        return profile

    default_profile = resolve_profile(load_settings().get("default_profile_id"))
    if default_profile:
        return default_profile

    builtin = resolve_profile(BUILTIN_PROFILE_ID)
    if builtin:
        return builtin

    profiles = list_profiles()
    return profiles[0] if profiles else None


def sync_default_profile(new_profile_id=None):
    """让主模板始终指向一个真实存在的模板。

    - 主模板指向的文件被删掉时，重选一个自建模板，没有就退回内置示例。
    - 刚保存或刚上传的模板（new_profile_id）出现时，如果主模板还停在
      内置示例上，就把主模板交给新模板——用户建了自己的模板，默认值
      不该继续指向示例。
    """
    settings = load_settings()
    current = resolve_profile(settings.get("default_profile_id"))
    candidate = resolve_profile(new_profile_id) if new_profile_id else None

    needs_switch = current is None or (current["is_builtin"] and candidate is not None)
    if not needs_switch:
        return current

    chosen = candidate
    if chosen is None:
        owned = [p for p in list_profiles() if not p["is_builtin"]]
        chosen = max(owned, key=lambda p: os.path.getmtime(p["path"])) if owned else resolve_profile(BUILTIN_PROFILE_ID)

    if chosen and chosen["id"] != settings.get("default_profile_id"):
        settings["default_profile_id"] = chosen["id"]
        save_settings(settings)
    return chosen


def _cleanup_report_cache():
    now = time.time()
    expired = [key for key, value in REPORT_CACHE.items()
               if now - value["created_at"] > REPORT_TTL_SECONDS]
    for key in expired:
        REPORT_CACHE.pop(key, None)


def convert_doc_to_docx(doc_path):
    """仅在 macOS 且存在 textutil 时兼容转换旧 .doc。"""
    if sys.platform != "darwin" or not shutil.which("textutil"):
        return None
    docx_path = doc_path.rsplit(".", 1)[0] + ".docx"
    try:
        subprocess.run(
            ["textutil", "-convert", "docx", doc_path, "-output", docx_path],
            check=True, capture_output=True, timeout=30
        )
        return docx_path
    except Exception:
        return None


def get_spec_rules(profile_id=None):
    """根据不透明模板标识获取规则路径。"""
    profile = get_profile(profile_id)
    return profile["path"] if profile else None


def infer_enabled_checks(rules):
    """依据已出现的规则字段启用检查，避免自定义模板继承论文专用检查。"""
    enabled = set()
    sections = {rule.get("section", "") for rule in rules}
    if any(any(key in rule for key in ("font", "size_pt", "bold")) for rule in rules):
        enabled.add("font")
    if any(rule.get("alignment") for rule in rules):
        enabled.add("alignment")
    if any(rule.get("line_spacing_pt") for rule in rules):
        enabled.add("line_spacing")
    if sections & {"cover_title", "cover_field"}:
        enabled.add("cover")
    if sections & {"level1_title", "level2_title", "level3_title"}:
        enabled.add("title_numbering")
    if sections & {"figure_caption", "table_caption"}:
        enabled.add("figure_table_numbering")
    if any(section.startswith("reference_") for section in sections):
        enabled.add("references")
    if "page_setup" in sections:
        enabled.add("page_setup")
    if "page_header" in sections:
        enabled.add("header")
    if any(section.startswith("abstract_") or section.startswith("keywords_") for section in sections):
        enabled.add("abstract")
    if "word_count" in sections:
        enabled.add("word_count")
    if "footnote" in sections:
        enabled.add("footnotes")
    if "document_structure" in sections:
        enabled.add("structure")
    return sorted(enabled)


@app.context_processor
def inject_nav_profile():
    """导航链接要带上当前模板，否则点过去会跳回主模板。

    模板规则页在看 A 模板、点「格式参考」却看到主模板的数字，
    是很容易误导人的错位，所以统一由这里提供。
    """
    profile = get_profile(request.args.get("profile"))
    return {"nav_profile_id": profile["id"] if profile else None}


@app.route("/")
def index():
    profiles = list_profiles()
    selected_profile = get_profile(request.args.get("profile"))
    return render_template(
        "index.html",
        profiles=profiles,
        selected_profile=selected_profile,
        has_own_profile=any(not p["is_builtin"] for p in profiles),
        max_files=MAX_FILES_PER_CHECK,
        max_total_mb=(app.config.get("MAX_CONTENT_LENGTH") or 0) // (1024 * 1024),
        old_doc_supported=sys.platform == "darwin" and bool(shutil.which("textutil")),
    )


@app.route("/examples")
def examples():
    # 加载当前规范规则，传给模板动态渲染示例
    profile = get_profile(request.args.get("profile"))
    rules_json_path = profile["path"] if profile else None
    rules_dict = {}
    raw_rules = []
    if rules_json_path:
        try:
            with open(rules_json_path, "r", encoding="utf-8") as f:
                raw = json.load(f)
            if isinstance(raw, dict):
                raw = raw.get("rules", [])
            raw_rules = [r for r in raw if isinstance(r, dict)]
            for r in raw_rules:
                section = r.get("section", "")
                rules_dict[section] = r
        except Exception:
            pass

    # 提取常用参数，给模板用
    def rule_val(section, key, default=""):
        return rules_dict.get(section, {}).get(key, default)

    def css_font(section, default_font="宋体"):
        """把规则里的字体名翻成浏览器能用的字族，样例才会显示成对的字体。"""
        name = rule_val(section, "font", default_font)
        return FONT_CSS.get(name) or FONT_CSS[default_font]

    enabled = infer_enabled_checks(raw_rules)

    params = {
        "body_font": rule_val("body_text", "font", "宋体"),
        "body_size": rule_val("body_text", "size_pt", 12),
        "body_size_name": rule_val("body_text", "size_name", "小四"),
        "body_css": css_font("body_text"),
        "caption_font": rule_val("figure_caption", "font", "宋体"),
        "caption_size": rule_val("figure_caption", "size_pt", 10.5),
        "caption_size_name": rule_val("figure_caption", "size_name", "五号"),
        "footnote_font": rule_val("footnote", "font", "宋体"),
        "footnote_size": rule_val("footnote", "size_pt", 9),
        "footnote_size_name": rule_val("footnote", "size_name", "小五"),
        "ref_entry_font": rule_val("reference_entry", "font", "宋体"),
        "ref_entry_size": rule_val("reference_entry", "size_pt", 10.5),
        "ref_entry_size_name": rule_val("reference_entry", "size_name", "五号"),
        "ref_css": css_font("reference_entry"),
        "line_spacing": rule_val("line_spacing_body", "line_spacing_pt", 22),
        "line_spacing_type": rule_val("line_spacing_body", "line_spacing_type", "固定值"),
        "ref_spacing": rule_val("reference_entry", "line_spacing_pt", 18),
        "ref_min": rule_val("reference_count", "min_count", 20),
        "ref_foreign_min": rule_val("reference_foreign_count", "min_count", 3),
        "fig_notes": rule_val("figure_caption", "notes", ""),
        "fn_notes": rule_val("footnote", "notes", ""),
        "title1_font": rule_val("level1_title", "font", "宋体"),
        "title1_size": rule_val("level1_title", "size_pt", 15),
        "title1_size_name": rule_val("level1_title", "size_name", "小三"),
        "title1_css": css_font("level1_title"),
        "title1_bold": rule_val("level1_title", "bold", True),
        "title2_font": rule_val("level2_title", "font", "宋体"),
        "title2_size": rule_val("level2_title", "size_pt", 14),
        "title2_size_name": rule_val("level2_title", "size_name", "四号"),
        "title2_css": css_font("level2_title"),
        "title2_bold": rule_val("level2_title", "bold", True),
        "title3_font": rule_val("level3_title", "font", "宋体"),
        "title3_size": rule_val("level3_title", "size_pt", 12),
        "title3_size_name": rule_val("level3_title", "size_name", "小四"),
        "title3_css": css_font("level3_title"),
        "title3_bold": rule_val("level3_title", "bold", True),
        "abstract_body_font": rule_val("abstract_zh_body", "font", "楷体"),
        "abstract_body_size": rule_val("abstract_zh_body", "size_pt", 12),
        "enabled_checks": [(key, CHECK_LABELS[key]) for key in CHECK_ORDER if key in enabled],
        "has_figure_rule": "figure_table_numbering" in enabled,
    }

    # 参考文献这张卡片的最后一行只说数量要求。
    # 条目排版是不是模板规定的，由卡片左上角的标签（来自模板 / 默认排版）表示，
    # 不再用一句话重复，也不把默认值说成模板要求。
    has_ref_entry = "reference_entry" in rules_dict
    has_ref_count = "reference_count" in rules_dict
    has_ref_foreign = "reference_foreign_count" in rules_dict
    params["has_ref_rule"] = has_ref_entry or has_ref_count or has_ref_foreign
    params["has_ref_entry"] = has_ref_entry
    if has_ref_count and has_ref_foreign:
        params["ref_note"] = "当前模板参考文献数量要求：参考文献不少于 %s 篇，其中外文不少于 %s 篇。" % (
            params["ref_min"], params["ref_foreign_min"])
    elif has_ref_count:
        params["ref_note"] = "当前模板参考文献数量要求：参考文献不少于 %s 篇。" % params["ref_min"]
    elif has_ref_foreign:
        params["ref_note"] = "当前模板参考文献数量要求：外文参考文献不少于 %s 篇。" % params["ref_foreign_min"]
    else:
        params["ref_note"] = "当前模板没规定参考文献数量。"

    return render_template("examples.html", profile=profile, **params)


@app.route("/editor")
def editor():
    profiles = list_profiles()
    selected_profile = get_profile(request.args.get("profile"))
    return render_template(
        "editor.html",
        profiles=profiles,
        selected_profile=selected_profile,
        old_doc_supported=sys.platform == "darwin" and bool(shutil.which("textutil")),
    )


@app.route("/editor/active")
def editor_active():
    """返回当前生效的规则及来源信息"""
    profile = get_profile(request.args.get("profile"))
    if not profile:
        return jsonify(ok=True, source="none", source_label="无规则文件", rules=[])

    try:
        payload = _profile_payload(profile["path"])
        rules = payload["rules"]
    except Exception as e:
        return jsonify(ok=False, error=str(e))

    return jsonify(
        ok=True,
        source="builtin" if profile["is_builtin"] else profile["id"],
        source_label=profile["name"],
        profile={key: profile.get(key) for key in (
            "id", "name", "document_type", "description", "source",
            "version", "updated_at", "enabled_checks", "is_builtin", "is_default",
            "rule_count", "meta_text"
        )},
        rules=rules,
    )


@app.route("/editor/rename", methods=["POST"])
def editor_rename():
    """改模板显示名称。内置示例也允许改，改名不动规则内容。"""
    data = request.get_json(silent=True) or {}
    profile = resolve_profile(data.get("profile_id"))
    if not profile:
        return jsonify(ok=False, error="模板不存在"), 404

    name = str(data.get("name") or "").strip()[:60]
    if not name:
        return jsonify(ok=False, error="名称不能为空"), 400

    for other in list_profiles():
        if other["id"] != profile["id"] and other["name"] == name:
            return jsonify(ok=False, error="已有同名模板，换个名字"), 400

    try:
        with open(profile["path"], "r", encoding="utf-8") as f:
            raw = json.load(f)
        if isinstance(raw, dict):
            raw.setdefault("profile", {})["name"] = name
            payload = raw
        else:
            payload = {"profile": {"name": name}, "rules": raw}
        with open(profile["path"], "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except (OSError, ValueError) as e:
        return jsonify(ok=False, error=f"写入失败：{e}"), 500

    return jsonify(ok=True, name=name)


@app.route("/editor/set-default", methods=["POST"])
def editor_set_default():
    """把某套模板设为主模板：首页和检查时默认用它。"""
    data = request.get_json(silent=True) or {}
    profile = resolve_profile(data.get("profile_id"))
    if not profile:
        return jsonify(ok=False, error="模板不存在"), 404

    settings = load_settings()
    settings["default_profile_id"] = profile["id"]
    if not save_settings(settings):
        return jsonify(ok=False, error="无法写入设置文件"), 500
    return jsonify(ok=True)


@app.route("/editor/save", methods=["POST"])
def editor_save():
    """保存编辑后的规则为新规范文件"""
    data = request.get_json()
    if not data or not data.get("name") or not data.get("rules"):
        return jsonify(ok=False, error="缺少名称或规则数据")

    rules = data["rules"]
    errors = []
    for i, r in enumerate(rules):
        if not r.get("section", "").strip():
            errors.append(f"第{i+1}条：section 不能为空")
        if not r.get("label", "").strip():
            errors.append(f"第{i+1}条：label 不能为空")
        if r.get("size_pt") is not None and r["size_pt"] != "":
            try:
                v = float(r["size_pt"])
                if v <= 0 or v > 100:
                    errors.append(f"第{i+1}条：字号 {v} 不合理")
            except (ValueError, TypeError):
                errors.append(f"第{i+1}条：字号格式错误")
        if r.get("line_spacing_pt") is not None and r["line_spacing_pt"] != "":
            try:
                v = float(r["line_spacing_pt"])
                if v <= 0 or v > 100:
                    errors.append(f"第{i+1}条：行距 {v} 不合理")
            except (ValueError, TypeError):
                errors.append(f"第{i+1}条：行距格式错误")
    if errors:
        return jsonify(ok=False, error="；".join(errors))

    profile_data = data.get("profile") or {}
    profile_name = str(data["name"]).strip()[:80]
    filename = f"profile_{uuid.uuid4().hex[:10]}.json"
    filepath = os.path.join(SPEC_DIR, filename)

    allowed_checks = {
        "font", "alignment", "line_spacing", "structure", "cover",
        "title_numbering", "figure_table_numbering", "references",
        "page_setup", "header", "abstract", "word_count", "footnotes",
    }
    enabled_checks = profile_data.get("enabled_checks")
    if enabled_checks is not None:
        enabled_checks = [item for item in enabled_checks if item in allowed_checks]
    else:
        enabled_checks = infer_enabled_checks(rules)

    payload = {
        "profile": {
            "name": profile_name,
            "document_type": str(profile_data.get("document_type") or "自定义 Word 文档")[:80],
            "description": str(profile_data.get("description") or "使用者确认的本地格式规则")[:200],
            "source": str(profile_data.get("source") or "使用者手动配置")[:200],
            "version": str(profile_data.get("version") or "1.0")[:30],
            "updated_at": datetime.now().strftime("%Y-%m-%d"),
            "enabled_checks": enabled_checks,
        },
        "rules": rules,
    }

    try:
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        promoted = sync_default_profile(filename)
        return jsonify(ok=True, profile_id=filename,
                       is_default=bool(promoted and promoted["id"] == filename))
    except Exception as e:
        return jsonify(ok=False, error=str(e))


@app.route("/editor/load/<spec_name>")
def editor_load(spec_name):
    """加载已有规范文件，返回 JSON"""
    profile = resolve_profile(spec_name)
    if not profile:
        return jsonify(ok=False, error="规范文件不存在"), 404

    try:
        payload = _profile_payload(profile["path"])
        return jsonify(ok=True, profile=payload["profile"], rules=payload["rules"])
    except (OSError, ValueError, TypeError) as e:
        return jsonify(ok=False, error=str(e)), 500


@app.route("/editor/delete/<spec_name>", methods=["POST"])
def editor_delete(spec_name):
    """删除已有规范"""
    profile = resolve_profile(spec_name)
    if not profile or profile["is_builtin"]:
        return jsonify(ok=False, error="不能删除该规则模板"), 400
    os.remove(profile["path"])
    # 删掉的如果是主模板，需要把默认值挪到别的模板上。
    sync_default_profile()
    return jsonify(ok=True)


@app.route("/check", methods=["POST"])
def check():
    files = request.files.getlist("files")
    if not files or all(f.filename == "" for f in files):
        flash("请上传至少一个 Word 文件", "error")
        return redirect(url_for("index"))

    if len(files) > MAX_FILES_PER_CHECK:
        flash(f"一次最多检查 {MAX_FILES_PER_CHECK} 个文件", "error")
        return redirect(url_for("index"))

    profile = get_profile(request.form.get("profile_id"))
    rules_json_path = profile["path"] if profile else None

    results = []
    for f in files:
        filename = _safe_display_name(f.filename)
        ext = _extension(filename)
        if ext not in ALLOWED_EXTENSIONS:
            continue

        # 保存上传文件
        task_id = uuid.uuid4().hex[:8]
        filepath = os.path.join(UPLOAD_DIR, f"{task_id}.{ext}")
        docx_path = None
        try:
            # 保存本身也可能失败（磁盘、权限、流中断），必须一并纳入清理范围。
            f.save(filepath)
            docx_path = filepath

            if ext == "doc":
                docx_path = convert_doc_to_docx(filepath)
                if not docx_path:
                    results.append(FailedResult(
                        filename=filename,
                        error=("当前系统无法安全转换旧 .doc，"
                               "请在 Word 或 WPS 中另存为 .docx 后重试"),
                    ))
                    continue

            result = check_document(docx_path, rules_json_path=rules_json_path)
            result.filename = filename
            results.append(result)
        except Exception as e:
            error_msg = str(e)
            if "Package not found" in error_msg or "not a zip file" in error_msg:
                error_msg = "文件格式损坏，请确认文件未损坏且为有效的 Word 文档"
            elif "encrypted" in error_msg.lower() or "password" in error_msg.lower():
                error_msg = "文件已加密或设置了密码，请先解密后再上传"
            elif "PermissionError" in error_msg:
                error_msg = "文件被占用，请关闭 Word 后重试"
            else:
                error_msg = "文件解析失败，请确认为有效的 Word 文档（.doc 或 .docx）"
            results.append(FailedResult(filename=filename, error=error_msg))

        finally:
            for cleanup_path in {filepath, docx_path}:
                if cleanup_path and os.path.exists(cleanup_path):
                    try:
                        os.remove(cleanup_path)
                    except OSError:
                        pass

    if not results:
        flash("没有有效的 Word 文件（.doc 或 .docx）", "error")
        return redirect(url_for("index"))

    # 生成报告
    _cleanup_report_cache()
    report_id = uuid.uuid4().hex[:16]
    report_html = generate_report_html(results, profile=profile)
    REPORT_CACHE[report_id] = {
        "created_at": time.time(),
        "html": report_html,
    }

    return render_template("report.html", results=results, report_id=report_id, profile=profile)


@app.route("/download/<report_id>")
def download_report(report_id):
    _cleanup_report_cache()
    cached = REPORT_CACHE.get(report_id)
    if cached:
        return send_file(
            io.BytesIO(cached["html"].encode("utf-8")),
            as_attachment=True,
            mimetype="text/html; charset=utf-8",
            download_name=f"文档格式检查报告_{report_id}.html",
        )
    return "报告不存在", 404


# 启动时校正一次主模板指向，避免设置文件里留着已删除的模板。
sync_default_profile()


def local_ipv4_addresses():
    """列出本机在局域网里的地址，方便把网址发给同一网络的人。"""
    import socket
    ips = set()
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ips.add(info[4][0])
    except OSError:
        pass
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.connect(("8.8.8.8", 80))
            ips.add(s.getsockname()[0])
    except OSError:
        pass
    return sorted(ip for ip in ips if not ip.startswith("127."))


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=5001)
    parser.add_argument(
        "--host",
        default=DEFAULT_HOST,
        help="默认只监听本机；改成 0.0.0.0 后，同一网络的其他人也能打开（用完请关闭）",
    )
    args = parser.parse_args()
    if args.host not in ("127.0.0.1", "localhost"):
        print("\n" + "=" * 52)
        print("已对局域网开放：同一网络内任何人都能打开下面的地址、上传文档。")
        print("试完请关闭这个窗口，或按 Ctrl+C 停止。")
        for ip in local_ipv4_addresses():
            print("    http://%s:%s" % (ip, args.port))
        print("=" * 52 + "\n")
    app.run(host=args.host, debug=False, port=args.port)
