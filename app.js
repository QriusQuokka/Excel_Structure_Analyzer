/* Excel 구조 분석기 — 브라우저(Pyodide) 구동 스크립트.
 *
 * 동작 원리:
 *   1) Pyodide(Python→WASM)를 CDN에서 로드한다(코드만 받음, 데이터 아님).
 *   2) micropip 으로 openpyxl 을 설치한다.
 *   3) CLI 와 똑같은 분석 엔진(excel_analyzer/*.py)을 fetch 해 가상 파일시스템에 써넣는다.
 *   4) 사용자가 고른 파일(여러 개 가능)은 "업로드 없이" 브라우저가 바이트로 읽어 가상 FS 에 기록한다.
 *   5) 파일마다 analyze_workbook() + render_html() 을 실행해 HTML 문자열을 받는다.
 *   6) 파일별 리포트를 탭으로 전환 미리보기 + 개별/전체(ZIP) 다운로드를 제공한다.
 *
 * 사용자 파일 바이트는 이 탭의 메모리(가상 FS)에만 존재하며 외부로 전송되지 않는다.
 * 흐름: 파일 선택 → "분석 대기" 목록 → [분석 시작] 클릭 → 결과.
 */

// ── 설정 ────────────────────────────────────────────────────────────────────
const PYODIDE_VERSION = "0.26.4";
const PYODIDE_INDEX_URL = `https://cdn.jsdelivr.net/pyodide/v${PYODIDE_VERSION}/full/`;

// 가상 FS 에 써넣을 엔진 파일들(사이트 루트 기준 상대경로). CLI 와 동일한 단일 소스.
const ENGINE_FILES = [
  "excel_analyzer/__init__.py",
  "excel_analyzer/formula_parser.py",
  "excel_analyzer/analyzer.py",
  "excel_analyzer/report.py",
];

const UPLOAD_DIR = "/home/pyodide/_uploads";

// ── DOM ──────────────────────────────────────────────────────────────────────
const $ = (id) => document.getElementById(id);
const statusEl = $("status"), statusText = $("statusText");
const dropzone = $("dropzone"), fileInput = $("fileInput");
const staging = $("staging"), stageList = $("stageList"), stageCount = $("stageCount");
const btnAnalyze = $("btnAnalyze"), btnAddMore = $("btnAddMore"), btnClearAll = $("btnClearAll");
const working = $("working"), workingText = $("workingText");
const errbox = $("errbox");
const result = $("result"), resultTitle = $("resultTitle");
const fileTabs = $("fileTabs"), previewFrame = $("previewFrame");
const btnDownload = $("btnDownload"), btnDownloadZip = $("btnDownloadZip"), btnReset = $("btnReset");

let pyodide = null;
let engineReady = false;

let pendingFiles = [];   // 분석 대기 파일 목록(File[])
let results = [];        // 분석 결과 [{name, baseName, ok, html, error, sheets, deps}]
let activeIdx = 0;       // 현재 보고 있는 결과 인덱스

// ── 상태 표시 헬퍼 ────────────────────────────────────────────────────────────
function setStatus(kind, text, spinner = false) {
  statusEl.className = `status ${kind}`;
  statusText.textContent = text;
  statusEl.firstElementChild.style.display = spinner ? "" : "none";
}
function showError(msg) { errbox.style.display = "block"; errbox.textContent = msg; }
function clearError() { errbox.style.display = "none"; errbox.textContent = ""; }
function fmtSize(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / 1024 / 1024).toFixed(1) + " MB";
}

// ── 1) Pyodide 스크립트 동적 로드 ────────────────────────────────────────────
function loadPyodideScript() {
  return new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = PYODIDE_INDEX_URL + "pyodide.js";
    s.onload = resolve;
    s.onerror = () => reject(new Error("Pyodide 스크립트를 불러오지 못했습니다(인터넷 연결 확인)."));
    document.head.appendChild(s);
  });
}

// ── 2) 엔진 부팅: Pyodide + openpyxl + 분석 엔진 ─────────────────────────────
async function bootEngine() {
  try {
    await loadPyodideScript();
    pyodide = await loadPyodide({ indexURL: PYODIDE_INDEX_URL });

    setStatus("loading", "openpyxl 라이브러리를 설치하는 중…", true);
    await pyodide.loadPackage("micropip");
    const micropip = pyodide.pyimport("micropip");
    await micropip.install("openpyxl");

    setStatus("loading", "분석 엔진을 불러오는 중…", true);
    await writeEngineFiles();

    // 엔진 함수 정의(한 번만). 경로를 받아 결과를 JSON 문자열로 돌려준다.
    pyodide.runPython(`
import json
from excel_analyzer import analyze_workbook
from excel_analyzer.report import render_html

def _run_analysis(path):
    analysis = analyze_workbook(path)
    return json.dumps({
        "html": render_html(analysis),
        "name": analysis.file_name,
        "sheets": len(analysis.sheets),
        "deps": len(analysis.dependencies),
    })
`);

    try { pyodide.FS.mkdir(UPLOAD_DIR); } catch (_) { /* 이미 있으면 무시 */ }

    engineReady = true;
    setStatus("ready", "준비 완료 — 엑셀 파일을 끌어다 놓거나 선택하세요(여러 개 가능).", false);
    dropzone.classList.remove("disabled");
  } catch (e) {
    console.error(e);
    setStatus("error", "엔진 준비에 실패했습니다: " + (e.message || e), false);
  }
}

// 엔진 .py 파일들을 fetch 해 가상 FS 의 excel_analyzer 패키지로 기록.
async function writeEngineFiles() {
  try { pyodide.FS.mkdir("/home/pyodide/excel_analyzer"); } catch (_) {}
  for (const rel of ENGINE_FILES) {
    const res = await fetch(rel, { cache: "no-cache" });
    if (!res.ok) throw new Error(`엔진 파일을 불러오지 못했습니다: ${rel} (${res.status})`);
    const text = await res.text();
    pyodide.FS.writeFile("/home/pyodide/" + rel, text);
  }
}

// ── 3) 파일 선택 → 분석 대기 목록 ────────────────────────────────────────────
function addFiles(fileLike) {
  const incoming = Array.from(fileLike);
  for (const f of incoming) {
    // 이름+크기로 중복 방지
    if (!pendingFiles.some((p) => p.name === f.name && p.size === f.size)) {
      pendingFiles.push(f);
    }
  }
  renderStaging();
}

function renderStaging() {
  stageCount.textContent = pendingFiles.length;
  staging.style.display = pendingFiles.length ? "block" : "none";
  btnAnalyze.disabled = pendingFiles.length === 0;

  stageList.innerHTML = "";
  pendingFiles.forEach((f, i) => {
    const li = document.createElement("li");
    li.className = "stage-item";
    const ok = /\.(xlsx|xlsm)$/i.test(f.name);
    li.innerHTML =
      `<span class="nm">${escapeHtml(f.name)}</span>` +
      `<span class="sz">${fmtSize(f.size)}</span>` +
      (ok ? "" : `<span class="bad">⚠ .xlsx/.xlsm 아님</span>`) +
      `<button class="rm" title="제거" data-i="${i}">✕</button>`;
    stageList.appendChild(li);
  });
}

stageList.addEventListener("click", (e) => {
  const btn = e.target.closest(".rm");
  if (!btn) return;
  pendingFiles.splice(Number(btn.dataset.i), 1);
  renderStaging();
});

btnAddMore.addEventListener("click", () => { if (engineReady) fileInput.click(); });
btnClearAll.addEventListener("click", () => { pendingFiles = []; renderStaging(); });

// ── 4) 분석 실행(여러 파일 순차) ─────────────────────────────────────────────
btnAnalyze.addEventListener("click", analyzePending);

async function analyzePending() {
  if (!engineReady || pendingFiles.length === 0) return;
  clearError();
  staging.style.display = "none";
  dropzone.style.display = "none";
  result.style.display = "none";
  working.style.display = "flex";

  const files = pendingFiles.slice();
  results = [];

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    workingText.textContent = `분석 중… (${i + 1}/${files.length}) ${file.name}`;
    // UI 갱신 틈을 준다(무거운 동기 호출 전에 화면이 멈춘 듯 보이지 않도록).
    await new Promise((r) => setTimeout(r, 0));

    try {
      const buf = new Uint8Array(await file.arrayBuffer());
      const safeName = file.name.replace(/[\/\\]/g, "_");
      const fsPath = `${UPLOAD_DIR}/${safeName}`;
      pyodide.FS.writeFile(fsPath, buf);

      pyodide.globals.set("_upload_path", fsPath);
      const out = await pyodide.runPythonAsync("_run_analysis(_upload_path)");
      try { pyodide.FS.unlink(fsPath); } catch (_) {}

      const data = JSON.parse(out);
      results.push({
        name: file.name,
        baseName: safeName.replace(/\.(xlsx|xlsm)$/i, ""),
        ok: true, html: data.html, error: "",
        sheets: data.sheets, deps: data.deps,
      });
    } catch (e) {
      console.error(e);
      results.push({
        name: file.name,
        baseName: file.name.replace(/\.(xlsx|xlsm)$/i, ""),
        ok: false, html: "", error: cleanPyError(e), sheets: 0, deps: 0,
      });
    }
  }

  working.style.display = "none";
  pendingFiles = [];
  renderStaging();
  showResults();
}

// Python 예외 메시지에서 사용자에게 보여줄 핵심만 추린다.
function cleanPyError(e) {
  const msg = String(e.message || e);
  const m = msg.match(/(?:ValueError|Exception|KeyError|TypeError):\s*([\s\S]+?)\s*$/);
  if (m) return m[1].trim();
  return "분석 중 문제가 발생했습니다.\n" + msg;
}

// ── 5) 결과 표시 ─────────────────────────────────────────────────────────────
function showResults() {
  const okCount = results.filter((r) => r.ok).length;
  const failCount = results.length - okCount;
  resultTitle.textContent =
    `분석 결과 — 성공 ${okCount}개` + (failCount ? ` · 실패 ${failCount}개` : "");

  // 파일 탭
  fileTabs.innerHTML = "";
  results.forEach((r, i) => {
    const tab = document.createElement("button");
    tab.className = "file-tab" + (r.ok ? "" : " failed");
    tab.dataset.i = i;
    const badge = r.ok ? `시트 ${r.sheets}` : "실패";
    tab.innerHTML = `<span>${escapeHtml(r.name)}</span><span class="tab-badge">${badge}</span>`;
    fileTabs.appendChild(tab);
  });
  // 탭이 1개뿐이면 굳이 보여주지 않는다.
  fileTabs.style.display = results.length > 1 ? "flex" : "none";

  // ZIP 버튼은 성공 결과가 2개 이상일 때만 의미가 있다.
  btnDownloadZip.style.display = okCount >= 2 ? "" : "none";

  result.style.display = "block";
  // 첫 성공 파일을 기본 선택(없으면 0번).
  const first = results.findIndex((r) => r.ok);
  selectResult(first >= 0 ? first : 0);
  result.scrollIntoView({ behavior: "smooth", block: "start" });
}

function selectResult(idx) {
  activeIdx = idx;
  const r = results[idx];
  [...fileTabs.children].forEach((t, i) => t.classList.toggle("active", i === idx));

  if (r.ok) {
    previewFrame.srcdoc = r.html;
    btnDownload.disabled = false;
  } else {
    previewFrame.srcdoc = errorPage(r.name, r.error);
    btnDownload.disabled = true;
  }
}

fileTabs.addEventListener("click", (e) => {
  const tab = e.target.closest(".file-tab");
  if (tab) selectResult(Number(tab.dataset.i));
});

// 실패 파일 미리보기에 보여줄 간단한 안내 페이지.
function errorPage(name, msg) {
  return `<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">
<style>body{font-family:"Segoe UI","Malgun Gothic",sans-serif;padding:40px;color:#8a2e2e;}
.box{background:#fbeaea;border:1px solid #e6b8b8;border-radius:10px;padding:24px;max-width:640px;}
h2{margin:0 0 10px;font-size:1.1rem;} pre{white-space:pre-wrap;font:inherit;margin:0;color:#7a2a2a;}</style>
</head><body><div class="box"><h2>⚠ '${escapeHtml(name)}' 분석 실패</h2>
<pre>${escapeHtml(msg)}</pre></div></body></html>`;
}

// ── 6) 다운로드(개별 / ZIP) ──────────────────────────────────────────────────
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

btnDownload.addEventListener("click", () => {
  const r = results[activeIdx];
  if (!r || !r.ok) return;
  downloadBlob(new Blob([r.html], { type: "text/html;charset=utf-8" }),
    `${r.baseName || "분석결과"}_analysis.html`);
});

btnDownloadZip.addEventListener("click", async () => {
  const ok = results.filter((r) => r.ok);
  if (ok.length === 0) return;
  if (typeof JSZip === "undefined") {
    showError("ZIP 라이브러리를 불러오지 못했습니다(인터넷 연결 확인). 파일별 다운로드는 가능합니다.");
    return;
  }
  btnDownloadZip.disabled = true;
  const prev = btnDownloadZip.textContent;
  btnDownloadZip.textContent = "ZIP 생성 중…";
  try {
    const zip = new JSZip();
    const used = new Set();
    for (const r of ok) {
      let fname = `${r.baseName || "분석결과"}_analysis.html`;
      // 같은 이름 충돌 시 번호 부여
      let n = 1, base = fname;
      while (used.has(fname)) fname = base.replace(/\.html$/, `(${n++}).html`);
      used.add(fname);
      zip.file(fname, r.html);
    }
    const blob = await zip.generateAsync({ type: "blob" });
    downloadBlob(blob, "excel_분석결과.zip");
  } catch (e) {
    console.error(e);
    showError("ZIP 생성 중 문제가 발생했습니다: " + (e.message || e));
  } finally {
    btnDownloadZip.disabled = false;
    btnDownloadZip.textContent = prev;
  }
});

// ── 7) 초기화 ────────────────────────────────────────────────────────────────
btnReset.addEventListener("click", () => {
  results = []; pendingFiles = []; activeIdx = 0;
  clearError();
  result.style.display = "none";
  renderStaging();
  dropzone.style.display = "";
  fileInput.value = "";
  dropzone.scrollIntoView({ behavior: "smooth", block: "center" });
});

// ── 8) 드래그&드롭 / 파일 선택 배선 ──────────────────────────────────────────
dropzone.addEventListener("click", () => { if (engineReady) fileInput.click(); });
fileInput.addEventListener("change", () => {
  if (fileInput.files.length) addFiles(fileInput.files);
  fileInput.value = "";  // 같은 파일 다시 선택 가능하도록 초기화
});

["dragenter", "dragover"].forEach((ev) =>
  dropzone.addEventListener(ev, (e) => {
    e.preventDefault();
    if (engineReady) dropzone.classList.add("dragover");
  })
);
["dragleave", "drop"].forEach((ev) =>
  dropzone.addEventListener(ev, (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
  })
);
dropzone.addEventListener("drop", (e) => {
  if (!engineReady) return;
  if (e.dataTransfer.files.length) addFiles(e.dataTransfer.files);
});

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

// ── 시작 ─────────────────────────────────────────────────────────────────────
bootEngine();
