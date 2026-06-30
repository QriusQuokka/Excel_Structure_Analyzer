"""분석 결과를 단일 HTML 리포트로 렌더링한다.

시트 간 관계도와 셀 단위 추적기는 모두 외부 라이브러리 없이 순수 Canvas/SVG +
JavaScript 로 그려지므로, 다운로드한 HTML 을 인터넷 없이 열어도 드래그·화살표·
추적이 그대로 동작한다(프라이버시: 데이터가 기기 밖으로 나가지 않는다).

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
            f"<tr class='fml-row' data-sheet=\"{_esc(s.name)}\" data-coord=\"{_esc(coord)}\" "
            f"title='클릭하면 이 셀을 셀 단위 추적에서 펼칩니다'>"
            f"<td class='cell-coord'>{_esc(coord)}</td>"
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
            f"같은 시트 안을 참조하는 수식이 {a.self_refs}개 있습니다"
            "(시트 관계도에서는 생략, 셀 추적에는 반영)."
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


def _cells_json(a: WorkbookAnalysis) -> str:
    """셀 단위 그래프를 추적기 JS 가 쓰는 형태의 JSON 으로."""
    out = {}
    for key, n in a.cells.items():
        precs = []
        for p in n.precedents:
            if p.is_external:
                precs.append({"kind": "external", "raw": p.raw})
            else:
                precs.append({"kind": "range" if p.is_range else "cell",
                              "sheet": p.sheet, "ref": p.ref})
        out[key] = {"formula": n.formula, "value": n.value,
                    "precedents": precs, "dependents": n.dependents}
    return json.dumps(out, ensure_ascii=False)


def _sheets_json(a: WorkbookAnalysis) -> str:
    return json.dumps([s.name for s in a.sheets], ensure_ascii=False)


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

  /* 시트쌍 드릴다운 */
  #edgeDetail { margin-top:14px; font-size:.85rem; }
  #edgeDetail.empty { color:#9aa4af; }
  .pair { padding:6px 12px; border:1px solid #eef1f4; border-bottom:none; display:flex; gap:10px; align-items:center; }
  .pair:last-child { border-bottom:1px solid #eef1f4; }
  .pair .dst { font-family:Consolas,monospace; font-weight:700; color:#155d34; cursor:pointer; }
  .pair .dst:hover { text-decoration:underline; }
  .pair .arr { color:#9aa6b2; }
  .pair .src { font-family:Consolas,monospace; color:#37414c; }
  .pair .fml { margin-left:auto; font-family:Consolas,monospace; font-size:.76rem; color:#8a95a1;
               max-width:46%; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }

  /* 셀 추적기 */
  .tracer-hint { font-size:.8rem; color:#6b7682; margin-bottom:10px; }
  .tracer-controls { display:flex; gap:8px; align-items:center; flex-wrap:wrap; margin-bottom:6px; }
  .tracer-controls label { font-size:.85rem; color:#3f4a57; }
  .tracer-controls select, .tracer-controls input { font-size:.86rem; padding:6px 10px;
               border:1px solid #cdd5dd; border-radius:6px; font-family:inherit; }
  .tracer-controls button { font-size:.86rem; padding:6px 14px; border:1px solid #2a8c54;
               background:#2a8c54; color:#fff; border-radius:6px; font-weight:600; cursor:pointer; }
  .tracer-controls button:hover { background:#247a49; }
  .crumb { font-size:.78rem; color:#6b7682; margin:8px 0; }
  .crumb a { color:#1a8f7a; cursor:pointer; text-decoration:underline; }
  .tracer { position:relative; margin-top:10px; }
  .cols { position:relative; display:flex; gap:48px; align-items:flex-start; overflow-x:auto; padding:6px 4px 12px; }
  .flowcol { display:flex; flex-direction:column; gap:16px; min-width:118px; z-index:1; }
  .flowcol h4 { font-size:.72rem; color:#6b7682; font-weight:700; text-align:center; white-space:nowrap; margin-bottom:2px; }
  .group { border:1px solid #d4dbe2; border-radius:8px; background:#fff; padding:6px; box-shadow:0 1px 2px rgba(0,0,0,.05); }
  .group.hassel { border-color:#1a6b3c; box-shadow:0 2px 8px rgba(26,107,60,.16); }
  .gh { font-size:.74rem; font-weight:700; color:#3f4a57; text-align:center; padding:2px 4px 5px;
        border-bottom:1px solid #eef1f4; margin-bottom:5px; white-space:nowrap; }
  .group.hassel .gh { color:#14532b; }
  .gcells { display:flex; flex-direction:column; gap:5px; }
  .gcell { position:relative; z-index:1; font-family:Consolas,monospace; font-size:.8rem; color:#2b3a48;
           background:#f7f9fb; border:1px solid #e3e8ee; border-radius:5px; padding:5px 9px; cursor:pointer;
           text-align:center; white-space:nowrap; }
  .gcell:hover { border-color:#2a8c54; background:#eef7f1; }
  .gcell.sel { background:#e4f3e9; border-color:#1a6b3c; color:#14532b; font-weight:700; }
  .gcell.leaf { cursor:default; color:#6b7682; border-style:dashed; }
  .gcell.leaf:hover { background:#f7f9fb; border-color:#e3e8ee; }
  .gcell small { display:block; font-size:.62rem; color:#9aa4af; font-family:inherit; font-weight:400; }
  .gcell.sel small { color:#3f7d56; }
  .clipnote { font-size:.7rem; color:#b0763a; text-align:center; margin-top:6px; white-space:nowrap; }
  #flowSvg { position:absolute; top:0; left:0; pointer-events:none; overflow:visible; z-index:0; }
  .tracer-none { color:#aab2bc; font-size:.85rem; text-align:center; padding:18px 0; }

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
  .formula-table tbody tr.fml-row { cursor:pointer; }
  .formula-table tbody tr.fml-row:hover td { background:#f3faf6; }
  .formula-table tbody tr.fml-row:hover .cell-coord { text-decoration:underline; }
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

// ── 데이터 (서버에서 주입) ───────────────────────────────────────────────────
const NODES = __NODES__;
const EDGES = __EDGES__;
const EDGECELLS = __EDGECELLS__;
const CELLS = __CELLS__;
const SHEETS = __SHEETS__;

// 역할별 색 (report.py 의 ROLE_COLORS 와 동일)
const ROLE_COLORS = {
  input:        { bg:'#e6f4f1', border:'#1a8f7a', text:'#0e5d50' },
  intermediate: { bg:'#eef1f4', border:'#7e8a99', text:'#3f4a57' },
  output:       { bg:'#e4f3e9', border:'#1a6b3c', text:'#14532b' },
  isolated:     { bg:'#f3f4f6', border:'#b0b7c0', text:'#5a6470' }
};

// ── 시트 간 관계도 (Canvas) ─────────────────────────────────────────────────
const NODE_H = 34, MIN_SPACING = 64, PAD_X = 130, PAD_Y = 50, FONT = '600 12px "Malgun Gothic", sans-serif';
const canvas = document.getElementById('graphCanvas');
const ctx = canvas.getContext('2d');
let dpr = 1, pos = {}, sizes = {}, dragNode = null, dragOfsX = 0, dragOfsY = 0, seg = [], hoverSeg = -1;

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
  seg = [];
  if (!NODES.length) return;
  EDGES.forEach((e, i) => {
    const A = pos[e.from], B = pos[e.to], sa = sizes[e.from], sb = sizes[e.to];
    if (!A || !B || !sa || !sb) return;
    const dx = B.x - A.x, dy = B.y - A.y, len = Math.sqrt(dx*dx + dy*dy) || 1;
    const ux = dx/len, uy = dy/len;
    const s = boxEdge(A.x, A.y, sa.w/2, sa.h/2, ux, uy);
    const en = boxEdge(B.x, B.y, sb.w/2, sb.h/2, -ux, -uy);
    const hi = i === hoverSeg || (dragNode !== null && (e.from === dragNode || e.to === dragNode));
    drawArrow(s.x, s.y, en.x, en.y, hi);
    seg.push({ i, x1: s.x, y1: s.y, x2: en.x, y2: en.y, from: e.from, to: e.to });
  });
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
function hitNode(mx, my) {
  for (let i = NODES.length - 1; i >= 0; i--) {
    const nd = NODES[i], p = pos[nd.id], sz = sizes[nd.id];
    if (!p || !sz) continue;
    if (Math.abs(mx - p.x) <= sz.w/2 && Math.abs(my - p.y) <= sz.h/2) return nd.id;
  }
  return null;
}
function distSeg(px, py, s) {
  const dx = s.x2 - s.x1, dy = s.y2 - s.y1, l2 = dx*dx + dy*dy;
  let t = l2 ? ((px - s.x1)*dx + (py - s.y1)*dy) / l2 : 0;
  t = Math.max(0, Math.min(1, t));
  return Math.hypot(px - (s.x1 + t*dx), py - (s.y1 + t*dy));
}
function hitEdge(mx, my) { let best = -1, bd = 8; seg.forEach(s => { const d = distSeg(mx, my, s); if (d < bd) { bd = d; best = s.i; } }); return best; }
canvas.addEventListener('mousedown', e => {
  const {x, y} = clientToCanvas(e), hit = hitNode(x, y);
  if (hit !== null) { dragNode = hit; dragOfsX = x - pos[hit].x; dragOfsY = y - pos[hit].y;
    canvas.style.cursor = 'grabbing'; e.preventDefault(); }
});
canvas.addEventListener('mousemove', e => {
  const {x, y} = clientToCanvas(e);
  if (dragNode !== null) { pos[dragNode] = { x: x - dragOfsX, y: y - dragOfsY }; render(); }
  else { const en = hitEdge(x, y); if (en !== hoverSeg) { hoverSeg = en; render(); }
    canvas.style.cursor = hitNode(x, y) !== null ? 'grab' : (en >= 0 ? 'pointer' : 'default'); }
});
canvas.addEventListener('mouseup', () => { dragNode = null; canvas.style.cursor = 'default'; });
canvas.addEventListener('mouseleave', () => { dragNode = null; if (hoverSeg !== -1) { hoverSeg = -1; render(); } });
canvas.addEventListener('click', e => {
  const {x, y} = clientToCanvas(e); if (hitNode(x, y) !== null) return;
  const en = hitEdge(x, y); if (en >= 0) { const s = EDGES[en]; showEdge(s.from, s.to); }
});
canvas.addEventListener('touchstart', e => {
  const {x, y} = clientToCanvas(e), hit = hitNode(x, y);
  if (hit !== null) { dragNode = hit; dragOfsX = x - pos[hit].x; dragOfsY = y - pos[hit].y; e.preventDefault(); }
}, { passive: false });
canvas.addEventListener('touchmove', e => {
  if (dragNode !== null) { const {x, y} = clientToCanvas(e);
    pos[dragNode] = { x: x - dragOfsX, y: y - dragOfsY }; render(); e.preventDefault(); }
}, { passive: false });
canvas.addEventListener('touchend', () => { dragNode = null; });

function nameOf(id) { const n = NODES.find(n => n.id === id); return n ? n.name : ''; }
function escHtml(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function jsAttr(s) { return escHtml(String(s).replace(/\\/g,'\\\\').replace(/'/g,"\\'")); }
function showEdge(from, to) {
  const fn = nameOf(from), tn = nameOf(to), list = EDGECELLS[fn + '->' + tn] || [];
  const el = document.getElementById('edgeDetail');
  el.classList.remove('empty');
  let h = '<div style="margin-bottom:8px"><b>' + escHtml(fn) + '</b> → <b>' + escHtml(tn) +
          '</b> 사이 셀 연결 ' + list.length + '건 <span style="color:#8a95a1;font-size:.8rem">(소비 셀 ← 출처 셀)</span></div>';
  list.forEach(p => {
    h += '<div class="pair"><span class="dst" onclick="quickKey(\'' + jsAttr(p.dst) + '\')">' + escHtml(p.dst) + '</span>'
       + '<span class="arr">←</span><span class="src">' + escHtml(p.src) + '</span>'
       + '<span class="fml" title="' + escHtml(p.formula) + '">' + escHtml(p.formula) + '</span></div>';
  });
  el.innerHTML = h;
}

// ── 셀 단위 추적기 ───────────────────────────────────────────────────────────
const LIMIT = 12, WINDOW = 5;
let path = [];
function populateTracer() {
  const sel = document.getElementById('selSheet');
  SHEETS.forEach(s => { const o = document.createElement('option'); o.value = s; o.textContent = s; sel.appendChild(o); });
}
function quick(sheet, cell) { document.getElementById('selSheet').value = sheet; document.getElementById('inpCell').value = cell; trace(); }
function quickKey(key) { const i = key.lastIndexOf('!'); center(key.slice(0, i), key.slice(i + 1), true); }
function trace() { center(document.getElementById('selSheet').value,
    document.getElementById('inpCell').value.trim().toUpperCase().replace(/\$/g, ''), true); }
function center(sheet, coord, reset) {
  if (!coord) return;
  const key = sheet + '!' + coord;
  if (reset) path = [key];
  else { const ix = path.indexOf(key); path = ix >= 0 ? path.slice(0, ix + 1) : [...path, key]; }
  document.getElementById('selSheet').value = sheet;
  document.getElementById('inpCell').value = coord;
  document.getElementById('tracer').style.display = 'block';
  renderCrumb(); renderFlow(key);
}
function renderCrumb() {
  document.getElementById('crumb').innerHTML = '경로: ' + path.map((k, i) =>
    i === path.length - 1 ? ('<b>' + escHtml(k) + '</b>') : ('<a onclick="jump(' + i + ')">' + escHtml(k) + '</a>')).join(' → ');
}
function jump(i) { const k = path[i], j = k.lastIndexOf('!'); path = path.slice(0, i + 1); center(k.slice(0, j), k.slice(j + 1), false); }

function bfsUp(start) {
  const seen = {}, edges = [], q = [[start, 0]]; seen[start] = 1;
  while (q.length) { const it = q.shift(), k = it[0], d = it[1]; if (d <= -LIMIT) continue;
    const c = CELLS[k]; if (!c) continue;
    c.precedents.forEach(p => { const sk = p.kind === 'external' ? ('ext:' + p.raw) : (p.sheet + '!' + p.ref);
      edges.push([sk, k]); if (!seen[sk]) { seen[sk] = 1; q.push([sk, d - 1]); } }); }
  return { seen, edges };
}
function bfsDown(start) {
  const seen = {}, edges = [], q = [[start, 0]]; seen[start] = 1;
  while (q.length) { const it = q.shift(), k = it[0], d = it[1]; if (d >= LIMIT) continue;
    const c = CELLS[k]; if (!c) continue;
    (c.dependents || []).forEach(dep => { edges.push([k, dep]); if (!seen[dep]) { seen[dep] = 1; q.push([dep, d + 1]); } }); }
  return { seen, edges };
}
function subgraph(sel0) {
  const up = bfsUp(sel0), dn = bfsDown(sel0), nodes = {};
  Object.keys(up.seen).forEach(k => nodes[k] = 1); Object.keys(dn.seen).forEach(k => nodes[k] = 1); nodes[sel0] = 1;
  const eset = {}, edges = [];
  up.edges.concat(dn.edges).forEach(e => { if (!nodes[e[0]] || !nodes[e[1]]) return;
    const id = e[0] + '' + e[1]; if (!eset[id]) { eset[id] = 1; edges.push(e); } });
  return { nodes: Object.keys(nodes), edges };
}
function longestLevels(nodes, edges) {     // 부분그래프 내 최장 입력경로 길이로 레벨 결정
  const preds = {}; nodes.forEach(n => preds[n] = []); edges.forEach(e => { if (preds[e[1]]) preds[e[1]].push(e[0]); });
  const memo = {}, stk = {};
  function lv(n) { if (memo[n] !== undefined) return memo[n]; stk[n] = 1; let m = 0;
    (preds[n] || []).forEach(p => { if (!stk[p]) m = Math.max(m, lv(p) + 1); }); stk[n] = 0; return memo[n] = m; }
  nodes.forEach(lv); return memo;
}
function groupKey(n) { return n.indexOf('ext:') === 0 ? '외부 통합문서' : n.slice(0, n.lastIndexOf('!')); }
function nodeMeta(n) {
  if (n.indexOf('ext:') === 0) return { clickable: false, label: n.slice(4), sub: '외부파일' };
  const cell = n.slice(n.lastIndexOf('!') + 1);
  if (cell.indexOf(':') >= 0) return { clickable: false, label: cell, sub: '범위' };
  const c = CELLS[n];
  if (!c) return { clickable: false, label: cell, sub: '없음' };
  return { clickable: true, label: cell, sub: c.formula ? '수식' : '입력값' };
}
function cssesc(s) { return s.replace(/["\\]/g, '\\$&'); }
function renderFlow(sel0) {
  const g = subgraph(sel0), lv = longestLevels(g.nodes, g.edges), s0 = lv[sel0] || 0;
  let mn = s0, mx = s0; g.nodes.forEach(n => { if (lv[n] < mn) mn = lv[n]; if (lv[n] > mx) mx = lv[n]; });
  const upAll = s0 - mn, downAll = mx - s0; let L = upAll, R = downAll;
  while (L + R + 1 > WINDOW) { if (L >= R && L > 0) L--; else if (R > 0) R--; else break; }
  const lo = s0 - L, hi = s0 + R, show = n => lv[n] >= lo && lv[n] <= hi;
  const cols = document.getElementById('cols'); cols.innerHTML = '<svg id="flowSvg"></svg>';
  for (let l = lo; l <= hi; l++) {
    const col = document.createElement('div'); col.className = 'flowcol';
    const rel = l - s0, head = document.createElement('h4');
    head.textContent = rel < 0 ? ('상류 ' + (-rel) + '단') : (rel > 0 ? ('하류 ' + rel + '단') : '선택 셀');
    col.appendChild(head);
    const groups = {}; g.nodes.forEach(n => { if (lv[n] !== l) return; const gk = groupKey(n); (groups[gk] = groups[gk] || []).push(n); });
    const gks = Object.keys(groups);
    if (!gks.length) { const e = document.createElement('div'); e.className = 'tracer-none'; e.textContent = '—'; col.appendChild(e); }
    gks.forEach(gk => {
      const arr = groups[gk], hasSel = arr.indexOf(sel0) >= 0;
      const box = document.createElement('div'); box.className = 'group' + (hasSel ? ' hassel' : '');
      const gh = document.createElement('div'); gh.className = 'gh'; gh.textContent = gk; box.appendChild(gh);
      const gc = document.createElement('div'); gc.className = 'gcells';
      arr.forEach(n => {
        const m = nodeMeta(n), chip = document.createElement('div');
        chip.className = 'gcell' + (n === sel0 ? ' sel' : '') + (m.clickable ? '' : ' leaf');
        chip.setAttribute('data-key', n);
        chip.innerHTML = escHtml(m.label) + '<small>' + m.sub + '</small>';
        if (m.clickable) chip.onclick = () => { const j = n.lastIndexOf('!'); center(n.slice(0, j), n.slice(j + 1), false); };
        gc.appendChild(chip);
      });
      box.appendChild(gc); col.appendChild(box);
    });
    if (l === lo && upAll > L) { const c = document.createElement('div'); c.className = 'clipnote'; c.textContent = '◂ 상류 ' + (upAll - L) + '단 더 있음'; col.appendChild(c); }
    if (l === hi && downAll > R) { const c = document.createElement('div'); c.className = 'clipnote'; c.textContent = '하류 ' + (downAll - R) + '단 더 있음 ▸'; col.appendChild(c); }
    cols.appendChild(col);
  }
  requestAnimationFrame(() => drawFlowEdges(g.edges.filter(e => show(e[0]) && show(e[1]))));
}
function drawFlowEdges(edges) {
  const cols = document.getElementById('cols'), svg = document.getElementById('flowSvg');
  if (!svg) return;
  const base = cols.getBoundingClientRect();
  svg.setAttribute('width', cols.scrollWidth); svg.setAttribute('height', cols.scrollHeight);
  let body = '<defs><marker id="ah" markerWidth="9" markerHeight="9" refX="7" refY="3" orient="auto">'
           + '<path d="M0,0 L7,3 L0,6 Z" fill="#9aa6b2"/></marker></defs>';
  edges.forEach(e => {
    const A = cols.querySelector('[data-key="' + cssesc(e[0]) + '"]'), B = cols.querySelector('[data-key="' + cssesc(e[1]) + '"]');
    if (!A || !B) return;
    const ra = A.getBoundingClientRect(), rb = B.getBoundingClientRect();
    const x1 = ra.right - base.left, y1 = ra.top - base.top + ra.height / 2;
    const x2 = rb.left - base.left, y2 = rb.top - base.top + rb.height / 2, mx = (x1 + x2) / 2;
    body += '<path d="M' + x1 + ',' + y1 + ' C' + mx + ',' + y1 + ' ' + mx + ',' + y2 + ' ' + (x2 - 3) + ',' + y2 + '" fill="none" stroke="#9aa6b2" stroke-width="1.6" marker-end="url(#ah)"/>';
  });
  svg.innerHTML = body;
}

// 수식 상세에서 행 클릭 → 셀 단위 추적으로 연결
function traceCell(sheet, coord) {
  const body = document.getElementById('tracer').closest('.card-body');
  if (body && body.classList.contains('hidden')) {
    body.classList.remove('hidden');
    if (body.previousElementSibling) body.previousElementSibling.classList.remove('collapsed');
  }
  center(sheet, coord, true);
  document.getElementById('tracer').scrollIntoView({ behavior: 'smooth', block: 'center' });
}
document.addEventListener('click', e => {
  const tr = e.target.closest && e.target.closest('.fml-row');
  if (!tr) return;
  if (window.getSelection && String(window.getSelection())) return;  // 드래그 선택 중이면 무시
  traceCell(tr.dataset.sheet, tr.dataset.coord);
});

function initGraph() { setupCanvas(); pos = computeInitialPos(); render(); populateTracer(); }
window.addEventListener('resize', () => { setupCanvas(); render(); });
window.addEventListener('load', initGraph);
"""


def render_html(a: WorkbookAnalysis) -> str:
    roles = _roles(a)
    nodes_json, edges_json = _graph_data(a, roles)
    script = (
        _SCRIPT
        .replace("__NODES__", nodes_json)
        .replace("__EDGES__", edges_json)
        .replace("__EDGECELLS__", json.dumps(a.edge_cells, ensure_ascii=False))
        .replace("__CELLS__", _cells_json(a))
        .replace("__SHEETS__", _sheets_json(a))
    )
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
        <span class="graph-hint">노드(시트)를 드래그해 배치를 바꿀 수 있습니다. 화살표는 데이터 흐름(참조 대상 → 참조하는 시트) 방향이며, <b>화살표를 클릭</b>하면 아래에 두 시트 사이의 셀 연결이 나옵니다.</span>
        <button class="btn-reset" onclick="resetLayout()">배치 초기화</button>
      </div>
      <canvas id="graphCanvas"></canvas>
      <div class="legend">{_legend()}</div>
      <div class="flow-caption">{_flow_caption(a)}</div>
      <div id="edgeDetail" class="empty">시트를 잇는 화살표를 클릭하면 여기에 셀 연결(소비 셀 ← 출처 셀)이 표시됩니다.</div>
    </div>
  </div>

  <div class="card">
    <div class="card-header" onclick="toggleCollapse(this)">
      셀 단위 추적
      <span class="toggle-icon">▾</span>
    </div>
    <div class="card-body">
      <div class="tracer-hint">데이터는 <b>왼쪽(출처) → 오른쪽(소비)</b>으로 흐릅니다. 시트·셀을 고르고 [추적]하면 선택 셀 기준 전후로 최대 5단계까지 펼쳐집니다 — 최초 입력값이면 맨 왼쪽, 모두 받기만 하는 최종 셀이면 맨 오른쪽에 놓입니다. 박스를 클릭하면 그 셀로 이동해 계속 따라갈 수 있습니다.</div>
      <div class="tracer-controls">
        <label>시트 <select id="selSheet"></select></label>
        <label>셀 <input id="inpCell" size="7" placeholder="예: A1"></label>
        <button onclick="trace()">추적</button>
      </div>
      <div class="crumb" id="crumb"></div>
      <div class="tracer" id="tracer" style="display:none">
        <div class="cols" id="cols"><svg id="flowSvg"></svg></div>
      </div>
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
