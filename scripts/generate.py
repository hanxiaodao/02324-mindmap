"""
02324 离散数学 · 思维导图 生成器
从 .docx 笔记直接生成交互式 HTML 思维导图，零 LLM 依赖。

用法:
  python generate.py [--docx 02324_笔记.docx]

依赖:
  pip install python-docx
"""

import argparse
import os
import re
import sys
from collections import OrderedDict

from docx import Document
from docx.oxml.ns import qn

ROOT = os.path.dirname(os.path.abspath(__file__))        # scripts/
PROJECT_ROOT = os.path.dirname(ROOT)                      # 项目根目录

# ── HTML 模板（复用已有样式） ──

CSS = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>德克离散数学 · 考点笔记</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Noto+Serif+SC:wght@600;700&family=Noto+Sans+SC:wght@400;500;600&display=swap" rel="stylesheet">
<script>
(function(){
  var t;
  try{ t=localStorage.getItem('theme'); }catch(e){}
  if(t!=='light'&&t!=='dark'){ t=window.matchMedia&&window.matchMedia('(prefers-color-scheme: light)').matches?'light':'dark'; }
  document.documentElement.setAttribute('data-theme',t);
})();
</script>
<style>
:root {
  /* 九章 accent 色——饱和、可辨认 */
  --c1:#60a5fa; --c2:#34d399; --c3:#fbbf24;
  --c4:#f472b6; --c5:#a78bfa; --c6:#22d3ee;
  --c7:#fb923c; --c8:#a3e635; --c9:#fb7185;
  /* 页面基础色 */
  --bg:#0d1117; --surface:#161b22; --card:#1c2230;
  --border:#2a3344; --border-strong:#3d4f66;
  --text:#c9d2dd; --muted:#7d8fa8; --faint:#4a5568;
  /* 派生色（暗色） */
  --heading:#e2eaf2;
  --hover-bg:rgba(255,255,255,.03);
  --detail-bg:rgba(0,0,0,.15);
  --tag-bg:rgba(255,255,255,.07); --tag-text:#9ab;
  --formula-bg:#111c2e; --formula-text:#8fb5e8;
  --th-bg:rgba(255,255,255,.05);
  --row-hover-bg:rgba(255,255,255,.02);
  --tip:#fbbf24; --ok:#34d399; --err:#f87171; --em:#fbbf24;
  --detail-text:var(--muted); --detail-size:13px;
  color-scheme:dark;
}
/* ── 亮色模式 ── */
:root[data-theme="light"] {
  /* 暖纸色：去纯白、降一档亮度，减轻长时间阅读的刺眼感 */
  --bg:#f3f1ec; --surface:#faf9f5; --card:#faf9f5;
  --border:#e2dfd7; --border-strong:#c9c5ba;
  --text:#33393f; --muted:#626b7a; --faint:#9aa1ab;
  --heading:#242930;
  --hover-bg:rgba(45,45,55,.045);
  --detail-bg:#f5f3ee;
  --tag-bg:rgba(45,45,55,.06); --tag-text:#5d6570;
  --formula-bg:#eef2fa; --formula-text:#2c5cb0;
  --th-bg:rgba(45,45,55,.05);
  --row-hover-bg:rgba(45,45,55,.03);
  --tip:#a16207; --ok:#047857; --err:#dc2626; --em:#a16207;
  --detail-text:#4c5868; --detail-size:14px;
  /* 亮色下加深 accent，保证可读性 */
  --c1:#2563eb; --c2:#059669; --c3:#d97706;
  --c4:#db2777; --c5:#7c3aed; --c6:#0891b2;
  --c7:#ea580c; --c8:#65a30d; --c9:#e11d48;
  color-scheme:light;
}
* { box-sizing:border-box; margin:0; padding:0; }
body {
  background:var(--bg);
  color:var(--text);
  font-family:'Noto Sans SC','PingFang SC','Microsoft YaHei',sans-serif;
  font-size:15px;
  line-height:1.6;
  -webkit-font-smoothing:antialiased;
}

/* ── Header ── */
header {
  padding:40px 20px 32px;
  text-align:center;
  border-bottom:1px solid var(--border);
  margin-bottom:20px;
}
header h1 {
  font-family:'Noto Serif SC',serif;
  font-size:26px;
  font-weight:700;
  letter-spacing:2px;
  color:var(--heading);
  margin-bottom:10px;
}
.header-meta {
  display:flex;
  align-items:center;
  justify-content:center;
  gap:20px;
  flex-wrap:wrap;
}
.header-stat {
  font-size:12px;
  color:var(--muted);
  letter-spacing:.5px;
}
.header-stat strong {
  font-size:18px;
  font-weight:600;
  color:var(--text);
  display:block;
  font-family:'Noto Sans SC',sans-serif;
  line-height:1.2;
  margin-bottom:2px;
}
.header-divider {
  width:1px; height:28px;
  background:var(--border-strong);
}

/* ── Chapter list ── */
.chapters { padding:0 16px 60px; display:flex; flex-direction:column; gap:8px; max-width:860px; margin:0 auto; }

/* 章节卡片：用左侧彩色边框替代整片渐变背景 */
.ch {
  border-radius:10px;
  overflow:hidden;
  border:1px solid var(--border);
  background:var(--surface);
  transition:border-color .2s;
}
.ch:hover { border-color:var(--border-strong); }

.ch-title {
  display:flex;
  align-items:center;
  justify-content:space-between;
  padding:14px 18px 14px 20px;
  cursor:pointer;
  user-select:none;
  font-weight:600;
  font-size:15px;
  color:var(--text);
  position:relative;
  transition:background .15s;
}
.ch-title:hover { background:var(--hover-bg); }
.ch-title::before {
  content:'';
  position:absolute;
  left:0; top:0; bottom:0;
  width:3px;
  border-radius:0;
  transition:width .2s;
}
.ch.open .ch-title { background:var(--hover-bg); }

/* 章节色彩：通过 accent 变量控制，只改左边框 */
.ch1 { --accent:var(--c1); } .ch1 .ch-title::before { background:var(--c1); }
.ch2 { --accent:var(--c2); } .ch2 .ch-title::before { background:var(--c2); }
.ch3 { --accent:var(--c3); } .ch3 .ch-title::before { background:var(--c3); }
.ch4 { --accent:var(--c4); } .ch4 .ch-title::before { background:var(--c4); }
.ch5 { --accent:var(--c5); } .ch5 .ch-title::before { background:var(--c5); }
.ch6 { --accent:var(--c6); } .ch6 .ch-title::before { background:var(--c6); }
.ch7 { --accent:var(--c7); } .ch7 .ch-title::before { background:var(--c7); }
.ch8 { --accent:var(--c8); } .ch8 .ch-title::before { background:var(--c8); }
.ch9 { --accent:var(--c9); } .ch9 .ch-title::before { background:var(--c9); }

.ch-title-text { flex:1; }
.ch-title .badge {
  font-size:11px;
  font-weight:400;
  color:var(--muted);
  margin-left:8px;
  font-family:'Noto Sans SC',sans-serif;
}
.ch-title .arrow {
  font-size:10px;
  color:var(--faint);
  transition:transform .25s;
  flex-shrink:0;
  margin-left:12px;
}
.ch.open .arrow { transform:rotate(180deg); }

/* ── 展开后：章节内容区 ── */
.ch-body {
  background:var(--card);
  border-top:1px solid var(--border);
  display:none;
}
.ch.open .ch-body { display:block; }

/* ── 考点行 ── */
.kp { border-bottom:1px solid var(--border); }
.kp:last-child { border-bottom:none; }

.kp-title {
  display:flex;
  align-items:center;
  gap:12px;
  padding:11px 18px 11px 20px;
  cursor:pointer;
  user-select:none;
  transition:background .15s;
}
.kp-title:hover { background:var(--hover-bg); }

/* 考点编号：彩色胶囊，与章节同色 */
.kp-num {
  font-size:10px;
  font-weight:600;
  color:var(--accent);
  min-width:36px;
  white-space:nowrap;
  letter-spacing:.3px;
  font-family:'Noto Sans SC',sans-serif;
}
.kp-name {
  flex:1;
  font-size:14px;
  font-weight:500;
  color:var(--text);
}
.kp-arrow {
  font-size:9px;
  color:var(--faint);
  transition:transform .2s;
  flex-shrink:0;
}
.kp.open .kp-arrow { transform:rotate(180deg); }

/* ── 考点详情 ── */
.kp-detail {
  display:none;
  padding:14px 20px 18px 68px;
  font-size:var(--detail-size);
  color:var(--detail-text);
  line-height:1.75;
  border-top:1px solid var(--border);
  background:var(--detail-bg);
}
.kp.open .kp-detail { display:block; }
.kp-detail b { color:var(--text); font-weight:600; }

.kp-detail .tag {
  display:inline-block;
  background:var(--tag-bg);
  border:1px solid var(--border);
  border-radius:4px;
  padding:1px 7px;
  font-size:11px;
  margin:1px 3px;
  color:var(--tag-text);
}

/* 公式：微蓝底 + 左侧 accent 边线 */
.kp-detail .formula {
  display:block;
  background:var(--formula-bg);
  border-left:2px solid var(--accent, #60a5fa);
  border-radius:0 6px 6px 0;
  padding:8px 14px;
  margin:8px 0;
  font-size:13px;
  color:var(--formula-text);
  letter-spacing:.6px;
  line-height:1.7;
  overflow-x:auto;
  white-space:pre-wrap;
  font-family:'Consolas','Courier New',monospace;
}

/* 表格 */
.kp-detail table {
  width:100%;
  border-collapse:collapse;
  margin:10px 0;
  font-size:12px;
}
.kp-detail td, .kp-detail th {
  border:1px solid var(--border);
  padding:6px 10px;
  text-align:center;
  vertical-align:middle;
}
.kp-detail th {
  background:var(--th-bg);
  color:var(--text);
  font-weight:600;
  font-size:11px;
  letter-spacing:.3px;
}
.kp-detail tr:hover td { background:var(--row-hover-bg); }
.kp-detail ul { padding-left:16px; }
.kp-detail li { margin:3px 0; }

/* 语义色 */
.kp-detail .tip { color:var(--tip); }
.kp-detail .ok  { color:var(--ok); }
.kp-detail .err { color:var(--err); }
.kp-detail .em  { color:var(--em); font-weight:600; }

/* 子标题 */
.kp-detail .kp-sub {
  display:block;
  color:var(--text);
  font-size:12px;
  font-weight:600;
  margin-top:14px;
  margin-bottom:4px;
  letter-spacing:.3px;
  padding-bottom:4px;
  border-bottom:1px solid var(--border);
}
.kp-detail .kp-sub:first-child { margin-top:2px; }

/* ── 主题切换按钮 ── */
.theme-toggle {
  position:fixed;
  top:16px; right:16px;
  width:38px; height:38px;
  border-radius:50%;
  border:1px solid var(--border);
  background:var(--surface);
  color:var(--muted);
  cursor:pointer;
  display:flex;
  align-items:center;
  justify-content:center;
  padding:0;
  z-index:100;
  box-shadow:0 2px 8px rgba(0,0,0,.12);
  transition:border-color .2s, color .2s, transform .15s;
}
.theme-toggle:hover {
  border-color:var(--border-strong);
  color:var(--text);
  transform:scale(1.06);
}
.theme-toggle .icon-sun { display:none; }
:root[data-theme="light"] .theme-toggle .icon-sun { display:block; }
:root[data-theme="light"] .theme-toggle .icon-moon { display:none; }

/* 键盘焦点 */
.ch-title:focus-visible, .kp-title:focus-visible { outline:2px solid var(--c1); outline-offset:-2px; }
.theme-toggle:focus-visible { outline:2px solid var(--c1); outline-offset:2px; }

/* ── 响应式 ── */
@media (max-width:600px) {
  header { padding:28px 16px 24px; }
  header h1 { font-size:21px; }
  .chapters { padding:0 10px 48px; }
  .kp-detail { padding-left:20px; }
  .header-divider { display:none; }
  .theme-toggle { top:10px; right:10px; }
}
@media (prefers-reduced-motion:reduce) {
  .arrow, .kp-arrow, .ch-title::before, .theme-toggle { transition:none; }
}
</style>
</head>
<body>
<button class="theme-toggle" onclick="toggleTheme()" aria-label="切换亮色/暗色模式" title="切换亮色/暗色模式">
  <svg class="icon-sun" width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M4.93 19.07l1.41-1.41M17.66 6.34l1.41-1.41"/></svg>
  <svg class="icon-moon" width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></svg>
</button>
<header>
  <h1>离散数学 · 考点笔记</h1>
  <div class="header-meta">
    <div class="header-stat"><strong>{ch_count}</strong>章</div>
    <div class="header-divider"></div>
    <div class="header-stat"><strong>全国自考 02324</strong>考点覆盖</div>
  </div>
</header>
<div class="chapters">
"""

FOOTER = """
<script>
function toggleTheme(){var r=document.documentElement,next=r.getAttribute('data-theme')==='light'?'dark':'light';r.setAttribute('data-theme',next);try{localStorage.setItem('theme',next)}catch(e){}}
document.addEventListener('keydown',function(e){if(e.key!=='Enter'&&e.key!==' ')return;var el=e.target;if(el.classList&&el.classList.contains('ch-title')){e.preventDefault();toggleCh(el)}else if(el.classList&&el.classList.contains('kp-title')){e.preventDefault();toggleKp(el)}});
function toggleCh(el){var ch=el.closest('.ch'),isOpen=ch.classList.contains('open');ch.classList.toggle('open');el.setAttribute('aria-expanded',String(!isOpen));if(!isOpen){ch.querySelectorAll('.kp-detail').forEach(function(d){d.style.display='none'});ch.querySelectorAll('.kp').forEach(function(k){k.classList.remove('open')});ch.querySelectorAll('.kp-title').forEach(function(t){t.setAttribute('aria-expanded','false')})}}
function toggleKp(el){var kp=el.closest('.kp'),isOpen=kp.classList.contains('open');kp.classList.toggle('open');el.setAttribute('aria-expanded',String(!isOpen));kp.querySelector('.kp-detail').style.display=isOpen?'none':'block'}
</script>
</div></body></html>
"""

# ── docx 解析 ──


def iter_block_items(doc):
    """按文档顺序遍历段落和表格。"""
    from docx.oxml.ns import qn
    body = doc.element.body
    para_idx = 0
    table_idx = 0
    for child in body:
        if child.tag == qn('w:p'):
            yield ('paragraph', doc.paragraphs[para_idx])
            para_idx += 1
        elif child.tag == qn('w:tbl'):
            if table_idx < len(doc.tables):
                yield ('table', doc.tables[table_idx])
                table_idx += 1


MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'


def para_full_text(para):
    """按段落内元素的原始顺序，拼接普通文字和 OMML 公式的纯文本。
    加粗 run 用 **...** 包裹，供 format_content 转换为 <b>。"""
    parts = []
    for child in para._element:
        tag = child.tag
        if tag == qn('w:r'):
            # 判断该 run 是否加粗（w:rPr/w:b 存在且不是 w:b val="0"）
            bold = False
            rpr = child.find(qn('w:rPr'))
            if rpr is not None:
                b_el = rpr.find(qn('w:b'))
                if b_el is not None:
                    val = b_el.get(qn('w:val'))
                    bold = val not in ('0', 'false', 'off')
            # 取 w:t 文本
            text = ''.join(t.text or '' for t in child.findall(qn('w:t')))
            if text:
                parts.append(f'**{text}**' if bold else text)
        elif tag == f'{{{MATH_NS}}}oMathPara':
            for om in child.findall(f'{{{MATH_NS}}}oMath'):
                parts.append(''.join(om.itertext()))
        elif tag == f'{{{MATH_NS}}}oMath':
            parts.append(''.join(child.itertext()))
        elif tag == qn('w:hyperlink'):
            for t in child.findall('.//' + qn('w:t')):
                parts.append(t.text or '')
    return ''.join(parts)


def table_to_html(table):
    """docx 表格 → HTML table 字符串。"""
    rows = []
    for ri, row in enumerate(table.rows):
        cells = []
        for ci, cell in enumerate(row.cells):
            tag = 'th' if ri == 0 else 'td'
            text = cell.text.strip()
            if not text and ri == 0:
                tag = 'th'
            cells.append(f'<{tag}>{h(text)}</{tag}>')
        rows.append(f'<tr>{"".join(cells)}</tr>')
    return f'<table>{"".join(rows)}</table>'


def h(text):
    """HTML 转义。"""
    return (text
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;'))


def format_content(text):
    """将纯文本转换为 HTML 格式。保留基本标记。"""
    text = h(text)
    # 保留常见符号
    text = text.replace('¬', '¬')
    text = text.replace('∧', '∧')
    text = text.replace('∨', '∨')
    text = text.replace('→', '→')
    text = text.replace('↔', '↔')
    text = text.replace('⇔', '⇔')
    text = text.replace('¬', '¬')
    text = text.replace('∈', '∈')
    text = text.replace('⊆', '⊆')
    text = text.replace('∪', '∪')
    text = text.replace('∩', '∩')
    text = text.replace('∅', '∅')
    text = text.replace('∀', '∀')
    text = text.replace('∃', '∃')
    # 加粗标记 **text** → <b>text</b>
    text = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text)
    # 换行
    text = text.replace('\n', '<br>')
    return text


def docx_to_html(docx_path):
    """主解析函数：docx → HTML 字符串。"""
    doc = Document(docx_path)

    chapters = []  # [(chapter_title, [(kp_num, kp_title, [(sub_title, [content_lines])])])]
    current_chapter = None
    current_kp = None
    current_sub = None
    chapter_kp_count = 0
    kp_num = 0

    # 跳过标题和目录的状态
    skip_title = True  # 跳过第一个 Heading 1 (德克离散数学_课程笔记)
    seen_first_h1 = False

    # 临时缓冲：当一个 Heading 3 出现前的 Normal 段落归入上一个 Heading 3，
    # 在 Heading 2 前的 Normal 段落归入上个 Heading 2
    # 我们用累积模式
    pending_text = []  # 不属于任何 H3 的累积文本（归入当前 Kp）

    # pending_text 归入当前考点（作为无标题子主题）
    # 在切换考点或保存章节前刷新
    def flush_pending():
        nonlocal pending_text
        if pending_text and current_kp is not None:
            current_kp[1].append(('', pending_text.copy()))
            pending_text = []

    for kind, item in iter_block_items(doc):
        if kind == 'paragraph':
            p = item
            style_name = p.style.name
            text = para_full_text(p).strip()

            if not text:
                continue

            if style_name.startswith('Heading'):
                level = int(style_name.split()[-1])

                if level == 1:
                    if skip_title:
                        skip_title = False
                        continue
                    # 新章节
                    if current_chapter is not None:
                        flush_pending()
                        if current_kp is not None:
                            current_chapter.append((kp_num, current_kp[0], current_kp[1]))
                            current_kp = None
                        chapters.append((current_chapter_title, chapter_kp_count, current_chapter))
                    current_chapter = []
                    current_chapter_title = text
                    chapter_kp_count = 0
                    current_kp = None
                    current_sub = None
                    kp_num = 0
                    pending_text = []

                elif level == 2:
                    flush_pending()
                    # 考点
                    if current_kp is not None:
                        current_chapter.append((kp_num, current_kp[0], current_kp[1]))
                    kp_num += 1
                    chapter_kp_count += 1
                    current_kp = [text, []]  # [title, [(sub_title, content_lines)]]
                    current_sub = None
                    pending_text = []

                elif level == 3:
                    flush_pending()
                    # 子主题
                    if current_kp is None:
                        # 没有 H2 直接出现 H3，创建一个无名考点
                        kp_num += 1
                        chapter_kp_count += 1
                        current_kp = ['', []]
                    current_sub = text
                    current_kp[1].append((text, []))
                    pending_text = []

            else:
                # Normal 段落 → 内容
                if current_sub is not None and current_kp is not None and current_kp[1]:
                    # 属于当前 H3
                    current_kp[1][-1][1].append(text)
                elif current_kp is not None:
                    # 属于当前考点（没有 H3 的内容）
                    pending_text.append(text)
                # 如果既无考点也无子主题，忽略（如目录文本）

        elif kind == 'table':
            tbl_html = table_to_html(item)
            if current_sub is not None and current_kp is not None and current_kp[1]:
                current_kp[1][-1][1].append(('__table__', tbl_html))
            elif current_kp is not None:
                pending_text.append(('__table__', tbl_html))

    # 保存最后一章
    flush_pending()
    if current_kp is not None:
        current_chapter.append((kp_num, current_kp[0], current_kp[1]))
    if current_chapter:
        chapters.append((current_chapter_title, chapter_kp_count, current_chapter))

    # ── 渲染 HTML ──
    ch_index = 0
    ch_html_parts = []

    for ch_title, kp_count, kps in chapters:
        ch_index += 1
        ch_class = f'ch{ch_index}'
        # 从标题提取序号（第一章 → 1）
        seq_match = re.search(r'第[一二三四五六七八九十]+章', ch_title)
        if seq_match:
            cn = seq_match.group()
            num_map = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
                       '六': 6, '七': 7, '八': 8, '九': 9, '十': 10}
            for k, v in num_map.items():
                if k in cn:
                    ch_class = f'ch{v}'
                    break

        parts = [f'<div class="ch {ch_class}">']
        title_display = ch_title.replace('：', ' · ', 1) if '：' in ch_title else ch_title
        title_display = title_display.replace('**', '')
        # 去掉可能的数字序号前缀
        title_display = re.sub(r'^\d+[\.、]\s*', '', title_display)
        parts.append(
            f'<div class="ch-title" onclick="toggleCh(this)" role="button" tabindex="0" aria-expanded="false">'
            f'<span>{h(title_display)}<span class="badge">{kp_count}考点</span></span>'
            f'<span class="arrow">▼</span></div>'
        )
        parts.append('<div class="ch-body">')

        for kp_num_item, kp_title, subs in kps:
            kp_title = kp_title.replace('**', '')
            # 解析考点序号
            num_match = re.match(r'考点(\d+)', kp_title)
            if num_match:
                display_num = f'考点{num_match.group(1)}'
                name = kp_title.split('：', 1)[-1] if '：' in kp_title else kp_title
            else:
                display_num = f'考点{kp_num_item}'
                name = kp_title if kp_title else '(未命名)'

            parts.append(
                f'<div class="kp">'
                f'<div class="kp-title" onclick="toggleKp(this)" role="button" tabindex="0" aria-expanded="false">'
                f'<span class="kp-num">{display_num}</span>'
                f'<span class="kp-name">{h(name)}</span>'
                f'<span class="kp-arrow">▼</span></div>'
                f'<div class="kp-detail">'
            )

            # 子主题
            for sub_title, content_lines in subs:
                if sub_title:
                    parts.append(f'<span class="kp-sub">{h(sub_title.replace("**", ""))}</span>')
                for line in content_lines:
                    if isinstance(line, tuple) and line[0] == '__table__':
                        parts.append(line[1])
                    elif isinstance(line, str):
                        line = line.strip()
                        if not line:
                            continue
                        # 检测公式（含 →, ∨, ∧, ¬ 等符号的行）
                        if any(sym in line for sym in ['→', '↔', '⇔', '∨', '∧', '¬', '∀', '∃']):
                            parts.append(f'<span class="formula">{format_content(line)}</span>')
                        else:
                            out = format_content(line)
                            if '<br>' in out:
                                parts.append(out)
                            else:
                                parts.append(f'{out}<br>')

            # 如果有 pending_text（属于考点但不属于任何子主题），放在子主题之后
            # 但我们已经把 pending_text 作为内容加入了

            parts.append('</div></div>')

        parts.append('</div></div>')
        ch_html_parts.append('\n'.join(parts))

    chapter_html = '\n\n'.join(ch_html_parts)

    # 组装完整 HTML
    html = CSS.replace('{ch_count}', str(ch_index)) + '\n' + chapter_html + '\n' + FOOTER
    return html


def main():
    parser = argparse.ArgumentParser(description='从 docx 生成思维导图 HTML')
    parser.add_argument('--docx', default=None, help='.docx 文件路径（默认自动查找）')
    parser.add_argument('-o', '--output', default=None, help='输出 HTML 路径（默认 index.html）')
    args = parser.parse_args()

    # 查找 docx
    docx_path = args.docx
    if docx_path is None:
        candidates = [f for f in os.listdir(PROJECT_ROOT) if f.endswith('.docx') and not f.startswith('~')]
        if not candidates:
            print('❌ 未找到 .docx 文件')
            sys.exit(1)
        docx_path = os.path.join(PROJECT_ROOT, candidates[0])
        print(f'📄 使用: {os.path.basename(docx_path)}')

    output = args.output or os.path.join(PROJECT_ROOT, 'index.html')

    print(f'📖 正在解析: {docx_path}')
    html = docx_to_html(docx_path)

    with open(output, 'w', encoding='utf-8') as f:
        f.write(html)

    file_size = os.path.getsize(output)
    line_count = html.count('\n')
    print(f'✅ 生成成功: {output}')
    print(f'   大小: {file_size/1024:.1f} KB, {line_count} 行')

    # 验证 HTML 结构
    div_open = html.count('<div')
    div_close = html.count('</div>')
    if div_open == div_close:
        print(f'   ✅ div 标签平衡: {div_open}')
    else:
        print(f'   ⚠️  div 标签不平衡: open={div_open} close={div_close}')


if __name__ == '__main__':
    main()
