"""분석 결과를 단일 HTML 리포트로 렌더링한다.

시트 간 관계도는 외부 라이브러리 없이 순수 Canvas + JavaScript 로 그려지므로
다운로드한 HTML 을 인터넷 없이 열어도 드래그·화살표가 그대로 동작한다.

노드 색은 무지개(장식)가 아니라 시트의 역할(입력단/가공/산출단/독립)을 의미로
구분한다 — 분석 도구다운 신뢰감·가독성을 위해서다.
"""

from __future__ import annotations

import html
import json
from datetime import datetime

from .analyzer import WorkbookAnalysis, SheetInfo


# 역할별 색 (Excel 그린톤 정렬, 절제된 3+1색). JS 의 ROLE_COLORS 와 동일하게 유지.
ROLE_COLORS = {
    "input":        {"bg": "#e6f4f1", "border": "#1a8f7a", "text": "#0e5d50", "label": "입력단"},
    "intermediate": {"bg": "#eef1f4", "border": "#7e8a99", "text": "#3f4a57", "label": "가공"},
    "output":       {"bg": "#e4f3e9", "border": "#1a6b3c", "text": "#14532b", "label": "산출단"},
    "isolated":     {"bg": "#f3f4f6", "border": "#b0b7c0", "text": "#5a6470", "label": "독립"},
}


def _esc(text) -> str:
    return html.escape(str(text), quote=True)


def _roles(a: WorkbookAnalysis) -> dict[str, str]:
    """시트명 -> 역할 문자열."""
    sources = {d.from_sheet for d in a.dependencies}     # 다른 시트가 참조함
    consumers = {d.to_sheet for d in a.dependencies}      # 다른 시트를 참조함
    roles = {}
    for s in a.sheets:
        is_src, is_con = s.name in sources, s.name in consumers
        if is_src and not is_con:
            roles[s.name] = "input"
        elif is_con and not is_src:
            roles[s.name] = "output"
        elif is_src and is_con:
            roles[s.name] = "intermediate"
        else:
            roles[s.name] = "isolated"
    return roles


def _badge(name: str, role: str) -> str:
    c = ROLE_COLORS[role]
    style = f"background:{c['bg']};border-color:{c['border']};color:{c['text']}"
    return f'<span class="sheet-badge" style="{style}">{_esc(name)}</span>'


def _stat_chips(a: WorkbookAnalysis) -> str:
    cyc = a.has_cycle()
    chips = [
        ("시트 수", str(len(a.sheets)), ""),
        ("총 수식 수", f"{a.total_formulas:,}", ""),
        ("시트 간 참조", str(len(a.dependencies)), ""),
        ("순환 참조", "있음" if cyc else "없음", "danger" if cyc else "ok"),
    ]
    out = []
    for label, value, cls in chips:
        vclass = f" {cls}" if cls else ""
        out.append(
            f'<div class="stat-chip"><div class="value{vclass}">{_esc(value)}</div>'
            f'<div class="label">{_esc(label)}</div></div>'
        )
    return "\n".join(out)


def _sheet_table_rows(a: WorkbookAnalysis, roles: dict[str, str]) -> str:
    rows = []
    for s in a.sheets:
        refs = "、".join(sorted(s.references)) if s.references else '<span class="muted">—</span>'
        state_tag = ""
        if s.state != "visible":
            state_tag = f' <span class="tag-hidden">{_esc(s.state)}</span>'
        rows.append(
            "<tr>"
            f"<td>{_badge(s.name, roles[s.name])}{state_tag}</td>"
            f"<td class='mono'>{_esc(s.dimensions)}</td>"
            f"<td class='num'>{s.n_rows} × {s.n_cols}</td>"
            f"<td class='num'>{s.formula_count:,}</td>"
            f"<td class='num'>{s.number_count:,}</td>"
            f"<td class='num'>{s.text_count:,}</td>"
            f"<td class='num'>{s.formula_ratio * 100:.1f}%</td>"
            f"<td>{refs}</td>"
            "</tr>"
        )
    return "\n".join(rows)


def _formula_panels(a: WorkbookAnalysis, roles: dict[str, str]) -> str:
    sheets_with_formulas = [s for s in a.sheets if s.formula_count > 0]
    if not sheets_with_formulas:
        return '<p class="muted">수식이 있는 시트가 없습니다.</p>'

    panels = []
    for s in sheets_with_formulas:
        body_rows = "\n".join(
            f"<tr><td class='cell-coord'>{_esc(coord)}</td>"
            f"<td class='formula-text'>{_esc(formula)}</td></tr>"
            for coord, formula in s.formulas
        )
        # 기본 접힘(collapsed): 헤더에 collapsed, 본문에 hidden
        panels.append(
            "<div class='formula-sheet'>"
            "<div class='formula-sheet-header collapsed' onclick='toggleCollapse(this)'>"
            f"<span class='fs-title'>{_badge(s.name, roles[s.name])}"
            f"<span class='badge'>수식 {s.formula_count:,}</span></span>"
            "<span class='toggle-icon'>▾</span>"
            "</div>"
            "<div class='formula-sheet-body hidden'>"
            "<div class='fml-scroll'>"
            "<table class='formula-table'>"
            "<thead><tr><th>셀</th><th>수식</th></tr></thead>"
            f"<tbody>{body_rows}</tbody>"
            "</table></div></div>"
            "</div>"
        )
    return "\n".join(panels)


def _legend() -> str:
    items = []
    for role in ("input", "intermediate", "output"):
        c = ROLE_COLORS[role]
        items.append(
            f'<span class="lg-item"><span class="lg-box" '
            f'style="background:{c["bg"]};border-color:{c["border"]}"></span>{c["label"]}</span>'
        )
    return "".join(items)


def _flow_caption(a: WorkbookAnalysis) -> str:
    ins, outs = a.input_sheets(), a.output_sheets()
    parts = []
    if ins:
        parts.append("입력단: <b>" + _esc("、".join(ins)) + "</b>")
    if outs:
        parts.append("산출단: <b>" + _esc("、".join(outs)) + "</b>")
    return " &nbsp;·&nbsp; ".join(parts)


def _warning_block(a: WorkbookAnalysis) -> str:
    items = list(a.warnings)
    cycle = a.find_cycle()
    if cycle:
        items.append("순환 참조가 있습니다: " + " → ".join(cycle))
    if a.external_refs:
        items.append(
            f"외부 통합문서를 참조하는 수식이 {a.external_refs}개 있습니다. "
            "이 관계는 다른 파일에 의존하므로 관계도에는 표시되지 않습니다."
        )
    if a.self_refs:
        items.append(
            f"같은 시트 안을 참조하는 수식이 {a.self_refs}개 있습니다(관계도에서는 생략)."
        )
    if not items:
        return ""
    lis = "\n".join(f"<li>{_esc(x)}</li>" for x in items)
    return f'<div class="alert alert-warn"><b>참고</b><ul>{lis}</ul></div>'


def _graph_data(a: WorkbookAnalysis, roles: dict[str, str]) -> tuple[str, str]:
    name_to_idx = {s.name: s.index for s in a.sheets}
    nodes = [{"id": s.index, "name": s.name, "role": roles[s.name]} for s in a.sheets]
    edges = []
    for d in a.dependencies:
        if d.from_sheet in name_to_idx and d.to_sheet in name_to_idx:
            edges.append({"from": name_to_idx[d.from_sheet], "to": name_to_idx[d.to_sheet]})
    return (
        json.dumps(nodes, ensure_ascii=False),
        json.dumps(edges, ensure_ascii=False),
    )


# ── 정적 자원 (CSS / JS) ─────────────────────────────────────────────────────

_STYLE = """
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: "Malgun Gothic","맑은 고딕",-apple-system,"Segoe UI",sans-serif;
         background:#eef1f4; color:#222b36; line-height:1.55; }
  .header { background:linear-gradient(135deg,#155d34,#2a8c54); color:#fff;
            padding:24px 32px; box-shadow:0 1px 0 rgba(0,0,0,.08); }
  .header h1 { font-size:1.4rem; font-weight:700; letter-spacing:-.01em; }
  .header .subtitle { font-size:.86rem; opacity:.9; margin-top:3px; }
  .container { max-width:1400px; margin:24px auto; padding:0 16px; }

  .stats-bar { display:flex; gap:14px; flex-wrap:wrap; margin-bottom:22px; }
  .stat-chip { background:#fff; border:1px solid #e3e8ee; border-radius:8px;
               padding:14px 24px; text-align:center; min-width:130px; }
  .stat-chip .value { font-size:1.55rem; font-weight:700; color:#155d34; font-variant-numeric:tabular-nums; }
  .stat-chip .value.danger { color:#b23b3b; }
  .stat-chip .label { font-size:.76rem; color:#6b7682; margin-top:3px; }

  .card { background:#fff; border:1px solid #e3e8ee; border-radius:8px;
          margin-bottom:20px; overflow:hidden; }
  .card-header { background:#f6f8fa; padding:13px 18px 13px 16px; border-bottom:1px solid #e9edf1;
                 border-left:3px solid #2a8c54; font-weight:700; font-size:.98rem; color:#2b3a48;
                 cursor:pointer; display:flex; align-items:center; gap:8px; user-select:none; }
  .card-header:hover { background:#eef2f5; }
  .card-header .toggle-icon { margin-left:auto; font-size:1.25rem; line-height:1; color:#6b7682; transition:transform .2s; }
  .card-header.collapsed .toggle-icon { transform:rotate(90deg); }
  .card-body { padding:18px; }
  .card-body.hidden { display:none; }

  table { border-collapse:collapse; width:100%; }
  .summary-table { font-size:.86rem; }
  .summary-table th { background:#155d34; color:#fff; padding:10px 12px; text-align:left;
                      font-weight:600; white-space:nowrap; }
  .summary-table td { padding:9px 12px; border-bottom:1px solid #eef1f4; vertical-align:middle; }
  .summary-table tr:last-child td { border-bottom:none; }
  .summary-table tr:hover td { background:#f5faf7; }
  .num { text-align:right; font-variant-numeric:tabular-nums; }
  .mono { font-family:"Consolas",monospace; font-size:.82rem; color:#48535f; }
  .muted { color:#9aa4af; }

  .sheet-badge { display:inline-block; padding:3px 11px; border-radius:5px;
                 border:1px solid; font-size:.82rem; font-weight:600; white-space:nowrap; }
  .tag-hidden { display:inline-block; background:#fbeaea; color:#b23b3b;
                font-size:.72rem; padding:1px 7px; border-radius:4px; }

  /* 그래프 */
  .graph-toolbar { display:flex; align-items:center; justify-content:space-between;
                   margin-bottom:10px; flex-wrap:wrap; gap:10px; }
  .graph-hint { font-size:.8rem; color:#6b7682; }
  .btn-reset { padding:6px 14px; border:1px solid #cdd5dd; background:#fff; color:#3f4a57;
               border-radius:6px; font-size:.82rem; font-weight:600; cursor:pointer; }
  .btn-reset:hover { background:#f3f6f8; border-color:#2a8c54; color:#155d34; }
  #graphCanvas { width:100%; display:block; border:1px solid #e9edf1; border-radius:6px;
                 background:#fbfcfd; cursor:default; }
  .legend { margin-top:10px; font-size:.8rem; color:#6b7682; display:flex; gap:16px; flex-wrap:wrap; align-items:center; }
  .lg-item { display:inline-flex; align-items:center; gap:6px; }
  .lg-box { width:14px; height:14px; border-radius:3px; border:1.5px solid; display:inline-block; }
  .flow-caption { font-size:.8rem; color:#55606b; margin-top:8px; }

  /* 시트별 수식 상세 */
  .formula-sheet { margin-bottom:8px; border:1px solid #e9edf1; border-radius:6px; overflow:hidden; }
  .formula-sheet-header { padding:9px 14px; cursor:pointer; background:#f6f8fa; display:flex;
                          align-items:center; justify-content:space-between; user-select:none; }
  .formula-sheet-header:hover { background:#eef2f5; }
  .formula-sheet-header .fs-title { display:inline-flex; align-items:center; gap:8px; }
  .formula-sheet-header .badge { background:#e3e8ee; color:#55606b; padding:2px 8px;
                                 border-radius:10px; font-size:.74rem; font-weight:600; }
  .formula-sheet-header .toggle-icon { font-size:1.15rem; line-height:1; color:#6b7682; transition:transform .2s; }
  .formula-sheet-header.collapsed .toggle-icon { transform:rotate(90deg); }
  .formula-sheet-body.hidden { display:none; }
  /* 펼쳤을 때 약 12줄까지만 보이고 그 이상은 스크롤 */
  .fml-scroll { max-height:340px; overflow:auto; }
  .formula-table { font-size:.82rem; }
  .formula-table th { background:#f6f8fa; padding:7px 12px; text-align:left; font-weight:600;
                      color:#55606b; border-bottom:1px solid #e9edf1; position:sticky; top:0; z-index:1; }
  .formula-table td { padding:5px 12px; border-bottom:1px solid #f1f4f7; vertical-align:top; }
  .cell-coord { font-family:"Consolas",monospace; font-weight:700; color:#155d34; white-space:nowrap; }
  .formula-text { font-family:"Consolas",monospace; color:#37414c; word-break:break-all; }

  .alert { padding:13px 16px; border-radius:6px; font-size:.88rem; }
  .alert-warn { background:#fdf6e3; border:1px solid #f0e0a8; color:#6b5618; }
  .alert ul { margin:8px 0 0 20px; }
"""

_SCRIPT = r"""
// ── 접철 토글 (큰 구획 + 시트별 패널 공용) ──────────────────────────────────
function toggleCollapse(headerEl) {
  const body = headerEl.nextElementSibling;
  const nowHidden = body.classList.toggle('hidden');
  headerEl.classList.toggle('collapsed');
  if (!nowHidden && body.querySelector && body.querySelector('#graphCanvas')) {
    setupCanvas(); render();
  }
}

// ── 그래프 데이터 ────────────────────────────────────────────────────────────
const NODES = __NODES__;
const EDGES = __EDGES__;

// 역할별 색 (report.py 의 ROLE_COLORS 와 동일)
const ROLE_COLORS = {
  input:        { bg:'#e6f4f1', border:'#1a8f7a', text:'#0e5d50' },
  intermediate: { bg:'#eef1f4', border:'#7e8a99', text:'#3f4a57' },
  output:       { bg:'#e4f3e9', border:'#1a6b3c', text:'#14532b' },
  isolated:     { bg:'#f3f4f6', border:'#b0b7c0', text:'#5a6470' }
};

const NODE_H = 34, MIN_SPACING = 64, PAD_X = 130, PAD_Y = 50, FONT = '600 12px "Malgun Gothic", sans-serif';
const canvas = document.getElementById('graphCanvas');
const ctx = canvas.getContext('2d');
let dpr = 1, pos = {}, sizes = {}, dragNode = null, dragOfsX = 0, dragOfsY = 0;

function computeSizes() {
  ctx.font = FONT;
  NODES.forEach(nd => {
    const w = Math.max(72, Math.ceil(ctx.measureText(nd.name).width) + 28);
    sizes[nd.id] = { w, h: NODE_H };
  });
}

function setupCanvas() {
  dpr = window.devicePixelRatio || 1;
  const cssW = canvas.parentElement.clientWidth;
  const cssH = calcRequiredHeight();
  canvas.style.height = cssH + 'px';
  canvas.width = cssW * dpr; canvas.height = cssH * dpr;
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  computeSizes();
}

function buildLayers() {
  const inDeg = {}, adj = {};
  NODES.forEach(nd => { inDeg[nd.id] = 0; adj[nd.id] = []; });
  EDGES.forEach(e => { inDeg[e.to] = (inDeg[e.to] || 0) + 1; });
  EDGES.forEach(e => { adj[e.from].push(e.to); });
  const layer = {};
  const roots = NODES.filter(nd => inDeg[nd.id] === 0).map(nd => nd.id);
  roots.forEach(id => layer[id] = 0);
  const queue = [...roots];
  while (queue.length) {
    const cur = queue.shift();
    (adj[cur] || []).forEach(nb => {
      inDeg[nb]--;
      if (inDeg[nb] === 0) { layer[nb] = (layer[cur] || 0) + 1; queue.push(nb); }
    });
  }
  NODES.forEach(nd => { if (layer[nd.id] === undefined) layer[nd.id] = 0; });
  const groups = {};
  NODES.forEach(nd => { const l = layer[nd.id]; (groups[l] = groups[l] || []).push(nd.id); });
  return { layer, groups };
}

function calcRequiredHeight() {
  if (!NODES.length) return 380;
  const { groups } = buildLayers();
  const maxCount = Math.max(...Object.values(groups).map(g => g.length));
  return Math.max(380, PAD_Y * 2 + maxCount * MIN_SPACING);
}

function computeInitialPos() {
  const cssW = canvas.getBoundingClientRect().width;
  const cssH = canvas.getBoundingClientRect().height;
  const { groups } = buildLayers();
  const layerNums = Object.keys(groups).map(Number);
  const maxLayer = Math.max(...layerNums);
  const usableW = cssW - PAD_X * 2;
  const newPos = {};
  layerNums.forEach(l => {
    const ids = groups[l];
    const x = PAD_X + (l / Math.max(maxLayer, 1)) * usableW;
    const groupTop = (cssH - ids.length * MIN_SPACING) / 2;
    ids.forEach((id, i) => { newPos[id] = { x, y: groupTop + (i + 0.5) * MIN_SPACING }; });
  });
  return newPos;
}

function resetLayout() { pos = computeInitialPos(); render(); }

// 박스(center, 반폭 hw, 반높이 hh) 경계에서 방향 (ux,uy) 로 나가는 지점
function boxEdge(cx, cy, hw, hh, ux, uy) {
  const tx = ux !== 0 ? hw / Math.abs(ux) : Infinity;
  const ty = uy !== 0 ? hh / Math.abs(uy) : Infinity;
  const t = Math.min(tx, ty);
  return { x: cx + ux * t, y: cy + uy * t };
}

function roundRect(x, y, w, h, r) {
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.arcTo(x + w, y, x + w, y + h, r);
  ctx.arcTo(x + w, y + h, x, y + h, r);
  ctx.arcTo(x, y + h, x, y, r);
  ctx.arcTo(x, y, x + w, y, r);
  ctx.closePath();
}

function drawArrow(sx, sy, ex, ey, hi) {
  const dx = ex - sx, dy = ey - sy, len = Math.sqrt(dx*dx + dy*dy);
  if (len < 1) return;
  const theta = Math.atan2(dy, dx), color = hi ? '#c0392b' : '#9aa6b2';
  const headLen = 9, ang = Math.PI / 7;
  ctx.strokeStyle = color; ctx.lineWidth = hi ? 2.2 : 1.5;
  ctx.beginPath(); ctx.moveTo(sx, sy); ctx.lineTo(ex, ey); ctx.stroke();
  ctx.fillStyle = color;
  ctx.beginPath(); ctx.moveTo(ex, ey);
  ctx.lineTo(ex - headLen*Math.cos(theta-ang), ey - headLen*Math.sin(theta-ang));
  ctx.lineTo(ex - headLen*Math.cos(theta+ang), ey - headLen*Math.sin(theta+ang));
  ctx.closePath(); ctx.fill();
}

function render() {
  const W = canvas.getBoundingClientRect().width, H = canvas.getBoundingClientRect().height;
  ctx.clearRect(0, 0, W, H);
  if (!NODES.length) return;

  // 엣지
  EDGES.forEach(e => {
    const A = pos[e.from], B = pos[e.to], sa = sizes[e.from], sb = sizes[e.to];
    if (!A || !B || !sa || !sb) return;
    const dx = B.x - A.x, dy = B.y - A.y, len = Math.sqrt(dx*dx + dy*dy) || 1;
    const ux = dx/len, uy = dy/len;
    const s = boxEdge(A.x, A.y, sa.w/2, sa.h/2, ux, uy);
    const en = boxEdge(B.x, B.y, sb.w/2, sb.h/2, -ux, -uy);
    const hi = dragNode !== null && (e.from === dragNode || e.to === dragNode);
    drawArrow(s.x, s.y, en.x, en.y, hi);
  });

  // 노드 (사각형 박스)
  ctx.font = FONT; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
  NODES.forEach(nd => {
    const p = pos[nd.id], sz = sizes[nd.id]; if (!p || !sz) return;
    const c = ROLE_COLORS[nd.role] || ROLE_COLORS.isolated;
    const isDragging = nd.id === dragNode;
    const x = p.x - sz.w/2, y = p.y - sz.h/2;
    if (isDragging) { ctx.shadowColor = 'rgba(0,0,0,0.18)'; ctx.shadowBlur = 10; ctx.shadowOffsetY = 2; }
    roundRect(x, y, sz.w, sz.h, 6);
    ctx.fillStyle = c.bg; ctx.fill();
    ctx.shadowColor = 'transparent'; ctx.shadowBlur = 0; ctx.shadowOffsetY = 0;
    ctx.strokeStyle = isDragging ? '#2b3a48' : c.border; ctx.lineWidth = isDragging ? 2 : 1.4; ctx.stroke();
    ctx.fillStyle = c.text; ctx.fillText(nd.name, p.x, p.y + 0.5);
  });
}

function clientToCanvas(e) {
  const rect = canvas.getBoundingClientRect();
  const src = e.touches ? e.touches[0] : e;
  return { x: src.clientX - rect.left, y: src.clientY - rect.top };
}
function hitTest(mx, my) {
  for (let i = NODES.length - 1; i >= 0; i--) {
    const nd = NODES[i], p = pos[nd.id], sz = sizes[nd.id];
    if (!p || !sz) continue;
    if (Math.abs(mx - p.x) <= sz.w/2 && Math.abs(my - p.y) <= sz.h/2) return nd.id;
  }
  return null;
}
canvas.addEventListener('mousedown', e => {
  const {x, y} = clientToCanvas(e), hit = hitTest(x, y);
  if (hit !== null) { dragNode = hit; dragOfsX = x - pos[hit].x; dragOfsY = y - pos[hit].y;
    canvas.style.cursor = 'grabbing'; e.preventDefault(); }
});
canvas.addEventListener('mousemove', e => {
  const {x, y} = clientToCanvas(e);
  if (dragNode !== null) { pos[dragNode] = { x: x - dragOfsX, y: y - dragOfsY }; render(); }
  else { canvas.style.cursor = hitTest(x, y) !== null ? 'grab' : 'default'; }
});
canvas.addEventListener('mouseup', () => { dragNode = null; canvas.style.cursor = 'default'; });
canvas.addEventListener('mouseleave', () => { dragNode = null; canvas.style.cursor = 'default'; });
canvas.addEventListener('touchstart', e => {
  const {x, y} = clientToCanvas(e), hit = hitTest(x, y);
  if (hit !== null) { dragNode = hit; dragOfsX = x - pos[hit].x; dragOfsY = y - pos[hit].y; e.preventDefault(); }
}, { passive: false });
canvas.addEventListener('touchmove', e => {
  if (dragNode !== null) { const {x, y} = clientToCanvas(e);
    pos[dragNode] = { x: x - dragOfsX, y: y - dragOfsY }; render(); e.preventDefault(); }
}, { passive: false });
canvas.addEventListener('touchend', () => { dragNode = null; });

function initGraph() { setupCanvas(); pos = computeInitialPos(); render(); }
window.addEventListener('resize', () => { setupCanvas(); render(); });
window.addEventListener('load', initGraph);
"""


def render_html(a: WorkbookAnalysis) -> str:
    roles = _roles(a)
    nodes_json, edges_json = _graph_data(a, roles)
    script = _SCRIPT.replace("__NODES__", nodes_json).replace("__EDGES__", edges_json)
    generated = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Excel 구조 분석 · {_esc(a.file_name)}</title>
<style>{_STYLE}</style>
</head>
<body>
<div class="header">
  <h1>Excel 구조 분석 리포트</h1>
  <div class="subtitle">{_esc(a.file_name)} &nbsp;|&nbsp; 분석일: {generated}</div>
</div>

<div class="container">

  <div class="stats-bar">
{_stat_chips(a)}
  </div>

  <div class="card">
    <div class="card-header" onclick="toggleCollapse(this)">
      시트 간 관계도
      <span class="toggle-icon">▾</span>
    </div>
    <div class="card-body">
      <div class="graph-toolbar">
        <span class="graph-hint">노드(시트)를 드래그해 배치를 바꿀 수 있습니다. 화살표는 데이터 흐름(참조 대상 → 참조하는 시트) 방향입니다.</span>
        <button class="btn-reset" onclick="resetLayout()">배치 초기화</button>
      </div>
      <canvas id="graphCanvas"></canvas>
      <div class="legend">{_legend()}</div>
      <div class="flow-caption">{_flow_caption(a)}</div>
    </div>
  </div>

  <div class="card">
    <div class="card-header" onclick="toggleCollapse(this)">
      시트 목록 및 사용 범위
      <span class="toggle-icon">▾</span>
    </div>
    <div class="card-body" style="padding:0">
      <table class="summary-table">
        <thead><tr>
          <th>시트명</th><th>사용 범위</th><th>크기</th><th>수식 수</th>
          <th>숫자 셀</th><th>텍스트 셀</th><th>수식 비율</th><th>참조하는 시트</th>
        </tr></thead>
        <tbody>
{_sheet_table_rows(a, roles)}
        </tbody>
      </table>
    </div>
  </div>

  <div class="card">
    <div class="card-header" onclick="toggleCollapse(this)">
      시트별 수식 상세
      <span class="toggle-icon">▾</span>
    </div>
    <div class="card-body">
{_formula_panels(a, roles)}
    </div>
  </div>

  {_warning_block(a)}

</div>
<script>{script}</script>
</body>
</html>
"""


def write_report(analysis: WorkbookAnalysis, output_path: str) -> None:
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(render_html(analysis))
