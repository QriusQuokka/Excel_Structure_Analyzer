/* Excel 구조 분석기 — 브라우저(Pyodide) 구동 스크립트.
 *
 * 동작 원리:
 *   1) Pyodide(Python→WASM)를 CDN에서 로드한다(코드만 받음, 데이터 아님).
 *   2) micropip 으로 openpyxl 을 설치한다.
 *   3) CLI 와 똑같은 분석 엔진(excel_analyzer/*.py)을 fetch 해 가상 파일시스템에 써넣는다.
 *   4) 사용자가 고른 파일은 "업로드 없이" 브라우저가 바이트로 읽어 가상 FS 에 기록한다.
 *   5) analyze_workbook() + render_html() 을 실행해 HTML 문자열을 받는다.
 *   6) 같은 페이지 iframe 으로 미리보기 + 다운로드 버튼을 제공한다.
 *
 * 사용자 파일 바이트는 이 탭의 메모리(가상 FS)에만 존재하며 외부로 전송되지 않는다.
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
const working = $("working"), workingText = $("workingText");
const errbox = $("errbox");
const result = $("result"), resultFname = $("resultFname");
const previewFrame = $("previewFrame");
const btnDownload = $("btnDownload"), btnReset = $("btnReset");

let pyodide = null;
let engineReady = false;
let lastHtml = "";       // 마지막 분석 결과 HTML
let lastBaseName = "";   // 다운로드 파일명용(확장자 제외 원본 이름)

// ── 상태 표시 헬퍼 ────────────────────────────────────────────────────────────
function setStatus(kind, text, spinner = false) {
  statusEl.className = `status ${kind}`;
  statusText.textContent = text;
  statusEl.firstElementChild.style.display = spinner ? "" : "none";
}
function showError(msg) {
  errbox.style.display = "block";
  errbox.textContent = msg;
}
function clearError() { errbox.style.display = "none"; errbox.textContent = ""; }

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

    // 엔진 함수 정의(한 번만). 경로를 받아 HTML 문자열을 돌려준다.
    pyodide.runPython(`
from excel_analyzer import analyze_workbook
from excel_analyzer.report import render_html

def _run_analysis(path):
    analysis = analyze_workbook(path)
    return render_html(analysis)
`);

    // 업로드 디렉터리 준비
    try { pyodide.FS.mkdir(UPLOAD_DIR); } catch (_) { /* 이미 있으면 무시 */ }

    engineReady = true;
    setStatus("ready", "준비 완료 — 엑셀 파일을 끌어다 놓거나 선택하세요.", false);
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

// ── 3) 파일 분석 ─────────────────────────────────────────────────────────────
async function analyzeFile(file) {
  if (!engineReady) return;
  clearError();
  result.style.display = "none";
  working.style.display = "flex";
  workingText.textContent = `'${file.name}' 분석 중…`;

  try {
    // 파일을 업로드 없이 브라우저에서 바이트로 읽어 가상 FS 에 기록.
    const buf = new Uint8Array(await file.arrayBuffer());
    const safeName = file.name.replace(/[\/\\]/g, "_");
    const fsPath = `${UPLOAD_DIR}/${safeName}`;
    pyodide.FS.writeFile(fsPath, buf);

    // 엔진 실행(파일 경로 → HTML 문자열).
    pyodide.globals.set("_upload_path", fsPath);
    const html = await pyodide.runPythonAsync("_run_analysis(_upload_path)");

    // 사용한 파일은 FS 에서 즉시 제거(메모리 정리).
    try { pyodide.FS.unlink(fsPath); } catch (_) {}

    lastHtml = html;
    lastBaseName = safeName.replace(/\.(xlsx|xlsm)$/i, "");
    showResult(file.name, html);
  } catch (e) {
    console.error(e);
    showError(cleanPyError(e));
  } finally {
    working.style.display = "none";
  }
}

// Python 예외 메시지에서 사용자에게 보여줄 핵심만 추린다.
function cleanPyError(e) {
  const msg = String(e.message || e);
  // Pyodide PythonError 는 마지막 줄에 "ValueError: ..." 형태로 실제 메시지가 온다.
  const m = msg.match(/(?:ValueError|Exception|KeyError|TypeError):\s*([\s\S]+?)\s*$/);
  if (m) return m[1].trim();
  return "분석 중 문제가 발생했습니다.\n" + msg;
}

function showResult(displayName, html) {
  resultFname.textContent = displayName;
  previewFrame.srcdoc = html;
  result.style.display = "block";
  result.scrollIntoView({ behavior: "smooth", block: "start" });
}

// ── 4) 다운로드 / 초기화 ─────────────────────────────────────────────────────
btnDownload.addEventListener("click", () => {
  if (!lastHtml) return;
  const blob = new Blob([lastHtml], { type: "text/html;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${lastBaseName || "분석결과"}_analysis.html`;
  a.click();
  URL.revokeObjectURL(url);
});

btnReset.addEventListener("click", () => {
  result.style.display = "none";
  clearError();
  lastHtml = ""; lastBaseName = "";
  fileInput.value = "";
  dropzone.scrollIntoView({ behavior: "smooth", block: "center" });
});

// ── 5) 드래그&드롭 / 파일 선택 배선 ──────────────────────────────────────────
dropzone.addEventListener("click", () => { if (engineReady) fileInput.click(); });
fileInput.addEventListener("change", () => {
  if (fileInput.files[0]) analyzeFile(fileInput.files[0]);
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
  const file = e.dataTransfer.files[0];
  if (file) analyzeFile(file);
});

// ── 시작 ─────────────────────────────────────────────────────────────────────
bootEngine();
