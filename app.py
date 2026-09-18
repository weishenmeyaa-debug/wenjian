"""文检 - 本地 Word 文档格式检查应用。"""

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

from flask import Flask, render_template, request, redirect, url_for, send_file, flash, jsonify
from checker import check_document
from spec_parser import parse_spec_docx
from report import generate_report_html

app = Flask(__name__)
app.secret_key = os.urandom(24)
app.config["SHOW_FOOTER"] = True
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
SPEC_DIR = os.path.join(BASE_DIR, "specs")

ALLOWED_EXTENSIONS = {"docx", "doc"}
MAX_FILES_PER_CHECK = 12
REPORT_TTL_SECONDS = 60 * 60
BUILTIN_PROFILE_ID = "__builtin__"

# 报告只在当前进程内短期保存，不写入磁盘。
REPORT_CACHE = {}

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(SPEC_DIR, exist_ok=True)


@app.context_processor
def inject_footer():
    return dict(show_footer=app.config.get("SHOW_FOOTER", False))


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
            "name": profile.get("name", "本地规则模板"),
            "document_type": profile.get("document_type", "Word 文档"),
            "description": profile.get("description", "由本地规则文件提供的格式检查要求"),
            "source": profile.get("source", "内置规则或经使用者确认的本地规则"),
            "version": profile.get("version", "1.0"),
            "updated_at": profile.get("updated_at", "未标注"),
            "enabled_checks": profile.get("enabled_checks"),
        },
        "rules": rules,
    }


def list_profiles():
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
    return profiles


def get_profile(profile_id=None):
    profiles = list_profiles()
    if not profiles:
        return None

    requested = profile_id or BUILTIN_PROFILE_ID
    for profile in profiles:
        if profile["id"] == requested:
            return profile
    return profiles[0]


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


@app.route("/")
def index():
    profiles = list_profiles()
    selected_profile = get_profile(request.args.get("profile"))
    return render_template(
        "index.html",
        profiles=profiles,
        selected_profile=selected_profile,
        max_files=MAX_FILES_PER_CHECK,
        old_doc_supported=sys.platform == "darwin" and bool(shutil.which("textutil")),
    )


@app.route("/examples")
def examples():
    # 加载当前规范规则，传给模板动态渲染示例
    profile = get_profile(request.args.get("profile"))
    rules_json_path = profile["path"] if profile else None
    rules_dict = {}
    if rules_json_path:
        try:
            with open(rules_json_path, "r", encoding="utf-8") as f:
                raw = json.load(f)
            if isinstance(raw, dict):
                raw = raw.get("rules", [])
            for r in raw:
                section = r.get("section", "")
                rules_dict[section] = r
        except Exception:
            pass

    # 提取常用参数，给模板用
    def rule_val(section, key, default=""):
        return rules_dict.get(section, {}).get(key, default)

    params = {
        "body_font": rule_val("body_text", "font", "宋体"),
        "body_size": rule_val("body_text", "size_pt", 12),
        "body_size_name": rule_val("body_text", "size_name", "小四"),
        "caption_font": rule_val("figure_caption", "font", "宋体"),
        "caption_size": rule_val("figure_caption", "size_pt", 10.5),
        "caption_size_name": rule_val("figure_caption", "size_name", "五号"),
        "footnote_font": rule_val("footnote", "font", "宋体"),
        "footnote_size": rule_val("footnote", "size_pt", 9),
        "footnote_size_name": rule_val("footnote", "size_name", "小五"),
        "ref_entry_font": rule_val("reference_entry", "font", "宋体"),
        "ref_entry_size": rule_val("reference_entry", "size_pt", 10.5),
        "ref_entry_size_name": rule_val("reference_entry", "size_name", "五号"),
        "line_spacing": rule_val("line_spacing_body", "line_spacing_pt", 22),
        "ref_spacing": rule_val("reference_entry", "line_spacing_pt", 18),
        "ref_min": rule_val("reference_count", "min_count", 20),
        "ref_foreign_min": rule_val("reference_foreign_count", "min_count", 3),
        "fig_notes": rule_val("figure_caption", "notes", ""),
        "fn_notes": rule_val("footnote", "notes", ""),
        "title1_font": rule_val("level1_title", "font", "宋体"),
        "title1_size": rule_val("level1_title", "size_pt", 15),
        "title1_size_name": rule_val("level1_title", "size_name", "小三"),
        "title2_font": rule_val("level2_title", "font", "宋体"),
        "title2_size": rule_val("level2_title", "size_pt", 14),
        "title2_size_name": rule_val("level2_title", "size_name", "四号"),
        "abstract_body_font": rule_val("abstract_zh_body", "font", "楷体"),
        "abstract_body_size": rule_val("abstract_zh_body", "size_pt", 12),
    }

    return render_template("examples.html", profile=profile, **params)


@app.route("/editor")
def editor():
    profiles = list_profiles()
    selected_profile = get_profile(request.args.get("profile"))
    return render_template("editor.html", profiles=profiles, selected_profile=selected_profile)


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
            "version", "updated_at", "enabled_checks", "is_builtin"
        )},
        rules=rules,
    )


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
        return jsonify(ok=True, profile_id=filename)
    except Exception as e:
        return jsonify(ok=False, error=str(e))


@app.route("/editor/load/<spec_name>")
def editor_load(spec_name):
    """加载已有规范文件，返回 JSON"""
    profile = get_profile(spec_name)
    if not profile or profile["id"] != spec_name:
        return jsonify(ok=False, error="规范文件不存在")

    try:
        payload = _profile_payload(profile["path"])
        return jsonify(ok=True, profile=payload["profile"], rules=payload["rules"])
    except Exception as e:
        return jsonify(ok=False, error=str(e))


@app.route("/editor/delete/<spec_name>", methods=["POST"])
def editor_delete(spec_name):
    """删除已有规范"""
    profile = get_profile(spec_name)
    if not profile or profile["id"] != spec_name or profile["is_builtin"]:
        return jsonify(ok=False, error="不能删除该规则模板"), 400
    os.remove(profile["path"])
    return jsonify(ok=True)


@app.route("/upload_spec", methods=["POST"])
def upload_spec():
    """上传规范文件，解析并保存规则"""
    spec_file = request.files.get("spec_file")
    if not spec_file or spec_file.filename == "":
        flash("请选择规范文件", "error")
        return redirect(url_for("index"))

    filename = _safe_display_name(spec_file.filename)
    ext = _extension(filename)

    if ext not in ("doc", "docx"):
        flash("规范文件仅支持 .doc 或 .docx 格式", "error")
        return redirect(url_for("index"))

    # 保存到随机命名的临时文件；原始文件名只用于页面说明。
    task_id = uuid.uuid4().hex[:8]
    filepath = os.path.join(SPEC_DIR, f"{task_id}_source.{ext}")
    spec_file.save(filepath)
    docx_path = filepath

    try:
        if ext == "doc":
            docx_path = convert_doc_to_docx(filepath)
            if not docx_path:
                flash("无法转换 .doc 文件，请另存为 .docx 后重试", "error")
                return redirect(url_for("editor"))

        rules = parse_spec_docx(docx_path)
        rules_json_path = os.path.join(SPEC_DIR, f"profile_{task_id}.json")
        payload = {
            "profile": {
                "name": f"待确认规则 · {Path(filename).stem[:40]}",
                "document_type": "待确认 Word 文档",
                "description": "由规范文件有限提取，使用前必须人工核验",
                "source": filename,
                "version": "草稿",
                "updated_at": datetime.now().strftime("%Y-%m-%d"),
                "enabled_checks": infer_enabled_checks(rules),
            },
            "rules": rules,
        }
        with open(rules_json_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        flash(f"规范文件解析成功，提取 {len(rules)} 条规则", "success")
    except Exception:
        flash("规范文件解析失败。请确认文件未损坏，并先另存为 .docx 后重试。", "error")
        return redirect(url_for("editor"))

    finally:
        for cleanup_path in {filepath, docx_path}:
            if cleanup_path and os.path.exists(cleanup_path):
                try:
                    os.remove(cleanup_path)
                except OSError:
                    pass

    return redirect(url_for("editor", profile=f"profile_{task_id}.json"))


@app.route("/delete_spec/<spec_name>", methods=["POST"])
def delete_spec(spec_name):
    """删除已上传的规范"""
    profile = get_profile(spec_name)
    if profile and profile["id"] == spec_name and not profile["is_builtin"]:
        os.remove(profile["path"])
    return redirect(url_for("index"))


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
        f.save(filepath)

        docx_path = filepath
        try:
            if ext == "doc":
                docx_path = convert_doc_to_docx(filepath)
                if not docx_path:
                    results.append(type("Err", (), {
                        "filename": filename,
                        "issues": [],
                        "summary": {"errors": 0, "warnings": 0, "infos": 0},
                        "error": "当前系统无法安全转换旧 .doc，请在 Word 或 WPS 中另存为 .docx 后重试"
                    })())
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
                error_msg = f"文件解析失败，请确认为有效的 Word 文档（.doc 或 .docx）"
            results.append(type("Err", (), {
                "filename": filename,
                "issues": [],
                "summary": {"errors": 0, "warnings": 0, "infos": 0},
                "error": error_msg
            })())

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


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=5001)
    args = parser.parse_args()
    app.run(host="127.0.0.1", debug=False, port=args.port)
