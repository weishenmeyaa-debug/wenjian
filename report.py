"""生成可离线打开的 HTML 格式检查报告。"""

from html import escape

from checker import Severity


def generate_report(results, output_path, profile=None):
    """生成 HTML 报告"""
    html = _build_html(results, profile=profile)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)


def generate_report_html(results, profile=None):
    """返回 HTML 字符串（用于页面直接展示）"""
    return _build_html(results, profile=profile)


def _build_html(results, profile=None):
    total_errors = sum(r.summary.get("errors", 0) for r in results if hasattr(r, 'summary'))
    total_warnings = sum(r.summary.get("warnings", 0) for r in results if hasattr(r, 'summary'))
    total_infos = sum(r.summary.get("infos", 0) for r in results if hasattr(r, 'summary'))

    sections = []
    for result in results:
        sections.append(_build_document_section(result))

    profile_name = escape((profile or {}).get("name", "未命名规则"))
    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>文档格式检查报告</title>
<style>
{_get_css()}
</style>
</head>
<body>
<div class="report">
  <header>
    <p class="eyebrow">WENJIAN · LOCAL REPORT</p>
    <h1>文档格式检查报告</h1>
    <p class="meta">规则：{profile_name} · 生成时间：{_import_datetime()}</p>
    <div class="summary-bar">
      <span class="badge error">必须修改 {total_errors}</span>
      <span class="badge warning">建议复核 {total_warnings}</span>
      <span class="badge info">人工确认 {total_infos}</span>
      <span class="badge count">共 {len(results)} 个文件</span>
    </div>
    <p class="notice">本报告是规则辅助检查结果，不替代正式规范、教师或评审人员的最终判断。</p>
  </header>
  {"".join(sections)}
</div>
</body>
</html>"""


def _import_datetime():
    from datetime import datetime
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _build_document_section(result):
    errors = [i for i in result.issues if i.severity == Severity.ERROR]
    warnings = [i for i in result.issues if i.severity == Severity.WARNING]
    infos = [i for i in result.issues if i.severity == Severity.INFO]

    error_rows = "".join(_build_issue_row(i) for i in errors)
    warning_rows = "".join(_build_issue_row(i) for i in warnings)
    info_rows = "".join(_build_issue_row(i) for i in infos)

    has_error = hasattr(result, 'error') and getattr(result, 'error', None)

    return f"""
  <section class="document">
    <h2>{escape(str(result.filename))}</h2>
    {"<p class='parse-error'>解析失败：" + escape(str(getattr(result, 'error', ''))) + "</p>" if has_error else ""}
    <div class="doc-summary">
      <span class="badge error">必须修改 {len(errors)}</span>
      <span class="badge warning">建议复核 {len(warnings)}</span>
      <span class="badge info">人工确认 {len(infos)}</span>
    </div>

    {_build_group("必须修改", "error", error_rows) if errors else ""}
    {_build_group("建议复核", "warning", warning_rows) if warnings else ""}
    {_build_group("需要人工确认", "info", info_rows) if infos else ""}

    {"" if result.issues or has_error else "<p class='all-good'>在当前规则和可识别范围内未发现格式问题，仍需人工复核。</p>"}
  </section>"""


def _build_group(title, level, rows):
    return f"""
    <section class="issue-group {level}">
      <h3>{title}</h3>
      <div class="issue-list">{rows}</div>
    </section>"""


def _build_issue_row(issue):
    content = getattr(issue, 'content', '') or ''
    content_display = f'<p class="content-preview">{escape(str(content))}</p>' if content else ''
    return f"""
      <article class="issue-card">
        <div class="issue-meta"><span class="tag">{escape(str(issue.category))}</span><span>{escape(str(issue.location))}</span></div>
        <h4>{escape(str(issue.description))}</h4>
        <div class="compare">
          <div><span>规则要求</span><strong class="expected">{escape(str(issue.expected))}</strong></div>
          <div><span>当前情况</span><strong class="actual">{escape(str(issue.actual))}</strong></div>
        </div>
        {content_display}
      </article>"""


def _get_css():
    return """
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: "PingFang SC", "Microsoft YaHei", sans-serif; background: #f3efe6; color: #24251f; line-height: 1.65; }
.report { max-width: 980px; margin: 0 auto; padding: 40px 22px 64px; }
header, .document { background: #fffdf8; border: 1px solid #ddd5c7; box-shadow: 0 14px 36px rgba(58,49,36,.08); }
header { padding: 34px; margin-bottom: 22px; border-radius: 18px; }
.eyebrow { color: #7b2d26; font-size: 12px; font-weight: 800; letter-spacing: .14em; margin-bottom: 8px; }
header h1 { font-family: "Songti SC", "SimSun", serif; font-size: 30px; margin-bottom: 8px; }
.meta { color: #6d6b62; font-size: 14px; margin-bottom: 20px; }
.summary-bar, .doc-summary { display: flex; gap: 9px; flex-wrap: wrap; }
.badge { padding: 6px 11px; border-radius: 999px; font-size: 13px; font-weight: 700; }
.badge.error { background: #f8e4e0; color: #8f2820; }
.badge.warning { background: #f8edcf; color: #805515; }
.badge.info { background: #e5ece8; color: #315e4c; }
.badge.count { background: #eeece5; color: #55564f; }
.notice { margin-top: 18px; padding-top: 16px; border-top: 1px solid #e7e0d5; color: #77736a; font-size: 13px; }
.document { padding: 30px; border-radius: 16px; margin-bottom: 22px; }
.document h2 { font-family: "Songti SC", "SimSun", serif; font-size: 22px; margin-bottom: 12px; word-break: break-word; }
.doc-summary { margin-bottom: 22px; }
.issue-group { margin-top: 26px; }
.issue-group h3 { font-size: 16px; margin-bottom: 11px; }
.issue-list { display: grid; gap: 10px; }
.issue-card { border: 1px solid #e3ddd2; border-left: 4px solid #7b2d26; background: #fff; border-radius: 10px; padding: 16px 18px; }
.warning .issue-card { border-left-color: #ad7626; }
.info .issue-card { border-left-color: #517665; }
.issue-meta { display: flex; align-items: center; gap: 9px; color: #737067; font-size: 12px; margin-bottom: 7px; }
.tag { background: #efede7; padding: 2px 7px; border-radius: 999px; color: #4d4e48; }
.issue-card h4 { font-size: 16px; margin-bottom: 12px; }
.compare { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
.compare div { background: #f7f4ed; border-radius: 8px; padding: 10px 12px; }
.compare span { display: block; color: #77736a; font-size: 12px; margin-bottom: 3px; }
.compare strong { font-size: 14px; }
.expected { color: #315e4c; }
.actual { color: #8f2820; }
.content-preview { margin-top: 10px; color: #625d52; background: #fbf0d6; padding: 9px 12px; border-radius: 7px; word-break: break-word; font-size: 13px; }
.parse-error { color: #8f2820; background: #f8e4e0; padding: 12px; border-radius: 8px; margin: 10px 0; }
.all-good { color: #315e4c; background: #e5ece8; padding: 16px; border-radius: 8px; text-align: center; font-size: 15px; }
@media (max-width: 640px) { .report { padding: 18px 12px 40px; } header, .document { padding: 20px; } .compare { grid-template-columns: 1fr; } }
@media print { body { background: #fff; } .report { max-width: none; padding: 0; } header, .document { box-shadow: none; break-inside: avoid; } }
"""
