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
    <p class="eyebrow">GUIJIAN · LOCAL REPORT</p>
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
    # 报告会被下载为独立文件，样式必须自包含；这里沿用与主界面
    # 一致的「编辑校对台」配色，并跟随系统深浅色。
    return """
:root {
  color-scheme: light dark;
  --paper: #f3efe6;
  --surface: #fffdf8;
  --surface-raised: #fffefb;
  --surface-sunken: #f7f4ed;
  --ink: #24251f;
  --ink-soft: #4c4d45;
  --muted: #706f67;
  --line: #ddd5c7;
  --accent: #7b2d26;
  --accent-soft: #f5e5e0;
  --accent-ink: #8f2820;
  --success: #315e4c;
  --success-soft: #e5eee9;
  --warning: #8a5d18;
  --warning-soft: #f7ecd1;
  --warning-ink: #7a5316;
  --highlight-bg: #fbf0d6;
  --highlight-ink: #625d52;
  --shadow: 0 16px 42px rgba(58,49,36,.10);
}
@media (prefers-color-scheme: dark) {
  :root {
    --paper: #191a17;
    --surface: #212320;
    --surface-raised: #282b27;
    --surface-sunken: #1b1d1a;
    --ink: #ece8de;
    --ink-soft: #cfcac0;
    --muted: #9b968b;
    --line: #383b35;
    --accent: #e08d7f;
    --accent-soft: #3a2420;
    --accent-ink: #eda294;
    --success: #8fc4a8;
    --success-soft: #1e2c26;
    --warning: #d9ab5c;
    --warning-soft: #322818;
    --warning-ink: #d9ab5c;
    --highlight-bg: #2c271d;
    --highlight-ink: #d9cfba;
    --shadow: 0 18px 44px rgba(0,0,0,.40);
  }
}

* { margin: 0; padding: 0; box-sizing: border-box; }
body {
  font-family: "PingFang SC", "Microsoft YaHei", sans-serif;
  background: var(--paper);
  color: var(--ink);
  line-height: 1.65;
  font-size: 16px;
}
.report { max-width: 980px; margin: 0 auto; padding: 40px 22px 64px; }

header, .document {
  background: var(--surface);
  border: 1px solid var(--line);
  box-shadow: var(--shadow);
}
header { padding: 34px; margin-bottom: 22px; border-radius: 18px; }
.eyebrow { color: var(--accent); font-size: 12px; font-weight: 800; letter-spacing: .14em; margin-bottom: 8px; }
header h1 { font-family: "Source Han Serif SC", "Songti SC", "SimSun", serif; font-size: 30px; margin-bottom: 8px; }
.meta { color: var(--muted); font-size: 14px; margin-bottom: 20px; }

.summary-bar, .doc-summary { display: flex; gap: 9px; flex-wrap: wrap; }
.badge { padding: 6px 11px; border-radius: 999px; font-size: 13px; font-weight: 700; }
.badge.error { background: var(--accent-soft); color: var(--accent-ink); }
.badge.warning { background: var(--warning-soft); color: var(--warning-ink); }
.badge.info { background: var(--success-soft); color: var(--success); }
.badge.count { background: var(--surface-sunken); color: var(--ink-soft); }

.notice { margin-top: 18px; padding-top: 16px; border-top: 1px solid var(--line); color: var(--muted); font-size: 13px; }

.document { padding: 30px; border-radius: 16px; margin-bottom: 22px; }
.document h2 { font-family: "Source Han Serif SC", "Songti SC", "SimSun", serif; font-size: 22px; margin-bottom: 12px; word-break: break-word; }
.doc-summary { margin-bottom: 22px; }

.issue-group { margin-top: 26px; }
.issue-group h3 { font-size: 16px; margin-bottom: 11px; }
.issue-list { display: grid; gap: 10px; }

.issue-card {
  border: 1px solid var(--line);
  border-left: 4px solid var(--accent);
  background: var(--surface-raised);
  border-radius: 10px;
  padding: 16px 18px;
}
.warning .issue-card { border-left-color: var(--warning); }
.info .issue-card { border-left-color: var(--success); }

.issue-meta { display: flex; flex-wrap: wrap; align-items: center; gap: 9px; color: var(--muted); font-size: 12px; margin-bottom: 7px; }
.tag { background: var(--surface-sunken); padding: 2px 7px; border-radius: 999px; color: var(--ink-soft); }
.issue-card h4 { font-size: 16px; margin-bottom: 12px; }

.compare { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
.compare div { background: var(--surface-sunken); border-radius: 8px; padding: 10px 12px; }
.compare span { display: block; color: var(--muted); font-size: 12px; margin-bottom: 3px; }
.compare strong { font-size: 14px; }
.expected { color: var(--success); }
.actual { color: var(--accent-ink); }

.content-preview { margin-top: 10px; color: var(--highlight-ink); background: var(--highlight-bg); padding: 9px 12px; border-radius: 7px; word-break: break-word; font-size: 13px; }
.parse-error { color: var(--accent-ink); background: var(--accent-soft); padding: 12px; border-radius: 8px; margin: 10px 0; }
.all-good { color: var(--success); background: var(--success-soft); padding: 16px; border-radius: 8px; text-align: center; font-size: 15px; }

@media (max-width: 640px) { .report { padding: 18px 12px 40px; } header, .document { padding: 20px; } .compare { grid-template-columns: 1fr; } }
@media print {
  :root { color-scheme: light; }
  body { background: #fff; color: #000; }
  .report { max-width: none; padding: 0; }
  header, .document { box-shadow: none; break-inside: avoid; }
}
"""
