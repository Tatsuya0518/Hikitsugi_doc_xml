/**
 * 引継ぎドキュメントシステム – フロントエンドロジック
 *
 * 機能:
 *  - 左サイドバーにチーム→作業コード→ドキュメントのツリーナビゲーション
 *  - ExcelをXMLに変換して表示（書式・画像・結合セル保持）
 *  - XMLソースエディタ（CodeMirror）
 *  - セル直接編集（編集モード）
 *  - ドキュメント横断全文検索
 *  - Excelエクスポート
 */
"use strict";

// ── State ─────────────────────────────────────────────────────────────────
const state = {
  currentDoc: null,   // { team, code, filename }
  xmlText: "",        // Raw XML string for current document
  editMode: false,
  cmEditor: null,     // CodeMirror instance
  currentTab: "view", // "view" | "xml"
  modifiedCells: {},  // { "row,col": newValue }
};

// ── DOM helpers ───────────────────────────────────────────────────────────
const $ = (id) => document.getElementById(id);

// ── Escape HTML ───────────────────────────────────────────────────────────
function esc(str) {
  if (str === null || str === undefined) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// ── Column number → letter(s) ─────────────────────────────────────────────
function colLetter(n) {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// ── Navigation tree ───────────────────────────────────────────────────────
async function loadStructure() {
  const tree = $("nav-tree");
  tree.innerHTML = '<p class="nav-empty">読み込み中…</p>';
  try {
    const resp = await fetch("/api/structure");
    const data = await resp.json();
    if (!data.teams || data.teams.length === 0) {
      tree.innerHTML =
        '<p class="nav-empty">data/watch/ にドキュメントがありません。</p>';
      return;
    }
    tree.innerHTML = "";
    data.teams.forEach((team) => renderTeamNode(team, tree));
  } catch (e) {
    tree.innerHTML = `<p class="nav-empty" style="color:red">読み込みエラー: ${esc(e.message)}</p>`;
  }
}

function renderTeamNode(team, container) {
  const wrapper = document.createElement("div");

  const teamEl = document.createElement("div");
  teamEl.className = "tree-team";
  teamEl.innerHTML = `<i class="tree-arrow">▶</i> 🏢 ${esc(team.name)}`;
  wrapper.appendChild(teamEl);

  const children = document.createElement("div");
  children.className = "tree-children";
  team.codes.forEach((code) => renderCodeNode(team.name, code, children));
  wrapper.appendChild(children);

  teamEl.addEventListener("click", () => {
    children.classList.toggle("open");
    teamEl.classList.toggle("open");
  });

  container.appendChild(wrapper);
}

function renderCodeNode(teamName, code, container) {
  const wrapper = document.createElement("div");

  const codeEl = document.createElement("div");
  codeEl.className = "tree-code";
  codeEl.innerHTML = `<i class="tree-arrow">▶</i> 📁 ${esc(code.name)}`;
  wrapper.appendChild(codeEl);

  const children = document.createElement("div");
  children.className = "tree-children";
  code.documents.forEach((filename) => {
    const docEl = document.createElement("div");
    docEl.className = "tree-doc";
    docEl.dataset.team = teamName;
    docEl.dataset.code = code.name;
    docEl.dataset.filename = filename;
    docEl.innerHTML = `📄 ${esc(filename)}`;
    docEl.addEventListener("click", () => openDocument(teamName, code.name, filename, docEl));
    children.appendChild(docEl);
  });
  wrapper.appendChild(children);

  codeEl.addEventListener("click", () => {
    children.classList.toggle("open");
    codeEl.classList.toggle("open");
  });

  container.appendChild(wrapper);
}

// ── Document loading ──────────────────────────────────────────────────────
async function openDocument(team, code, filename, navEl) {
  // Mark active in sidebar
  document.querySelectorAll(".tree-doc.active").forEach((el) =>
    el.classList.remove("active")
  );
  if (navEl) navEl.classList.add("active");

  // Expand parent nodes
  if (navEl) {
    let parent = navEl.parentElement;
    while (parent) {
      if (parent.classList.contains("tree-children")) {
        parent.classList.add("open");
        const prev = parent.previousElementSibling;
        if (prev) prev.classList.add("open");
      }
      parent = parent.parentElement;
    }
  }

  state.currentDoc = { team, code, filename };
  state.modifiedCells = {};
  leaveEditMode();

  $("doc-toolbar").style.display = "flex";
  $("doc-title").textContent = `${team} / ${code} / ${filename}`;
  $("welcome").style.display = "none";
  $("search-results").style.display = "none";

  showPane("view");
  $("view-pane").style.display = "flex";
  $("spreadsheet-container").innerHTML = loadingHTML();
  $("sheet-tabs").innerHTML = "";

  try {
    const resp = await fetch(`/api/document/${enc(team)}/${enc(code)}/${enc(filename)}`);
    if (!resp.ok) {
      const err = await resp.json();
      $("spreadsheet-container").innerHTML = errorHTML(err.error || "エラー");
      return;
    }
    state.xmlText = await resp.text();
    renderXmlToSpreadsheet(state.xmlText);
    if (state.cmEditor) state.cmEditor.setValue(state.xmlText);
  } catch (e) {
    $("spreadsheet-container").innerHTML = errorHTML(e.message);
  }
}

function enc(s) {
  return encodeURIComponent(s);
}

// ── Spreadsheet renderer ──────────────────────────────────────────────────
function renderXmlToSpreadsheet(xmlText) {
  let xmlDoc;
  try {
    xmlDoc = new DOMParser().parseFromString(xmlText, "application/xml");
    const parseErr = xmlDoc.querySelector("parsererror");
    if (parseErr) throw new Error(parseErr.textContent);
  } catch (e) {
    $("spreadsheet-container").innerHTML = errorHTML("XMLの解析に失敗しました: " + e.message);
    return;
  }

  const sheets = Array.from(xmlDoc.querySelectorAll("workbook > sheet"));
  if (sheets.length === 0) {
    $("spreadsheet-container").innerHTML = '<p style="padding:16px;color:#666">シートが見つかりません</p>';
    return;
  }

  // Build sheet tabs
  const tabsEl = $("sheet-tabs");
  tabsEl.innerHTML = "";
  const container = $("spreadsheet-container");
  container.innerHTML = "";

  sheets.forEach((sheet, idx) => {
    const name = sheet.getAttribute("name") || `Sheet${idx + 1}`;

    const tab = document.createElement("div");
    tab.className = "sheet-tab" + (idx === 0 ? " active" : "");
    tab.textContent = name;
    tab.dataset.idx = idx;
    tab.addEventListener("click", () => {
      document.querySelectorAll(".sheet-tab").forEach((t) => t.classList.remove("active"));
      document.querySelectorAll(".sheet-content").forEach((c) => c.classList.remove("active"));
      tab.classList.add("active");
      document.querySelector(`.sheet-content[data-idx="${idx}"]`).classList.add("active");
    });
    tabsEl.appendChild(tab);

    const contentDiv = document.createElement("div");
    contentDiv.className = "sheet-content" + (idx === 0 ? " active" : "");
    contentDiv.dataset.idx = idx;
    contentDiv.appendChild(buildSheetTable(sheet));
    container.appendChild(contentDiv);
  });
}

function buildSheetTable(sheetElem) {
  const maxRow = parseInt(sheetElem.querySelector("dimensions")?.getAttribute("maxRow") || 0);
  const maxCol = parseInt(sheetElem.querySelector("dimensions")?.getAttribute("maxCol") || 0);

  if (maxRow === 0 || maxCol === 0) {
    const p = document.createElement("p");
    p.style.cssText = "padding:16px;color:#666";
    p.textContent = "（空のシート）";
    return p;
  }

  // Build lookup: "row,col" → <cell> element
  const cellMap = {};
  sheetElem.querySelectorAll("rows > row > cell").forEach((c) => {
    cellMap[`${c.getAttribute("row")},${c.getAttribute("col")}`] = c;
  });

  // Build merge spanMap: "startRow,startCol" → { rowspan, colspan }
  const spanMap = {};
  const skipSet = new Set();
  sheetElem.querySelectorAll("merges > merge").forEach((m) => {
    const sr = parseInt(m.getAttribute("startRow"));
    const sc = parseInt(m.getAttribute("startCol"));
    const er = parseInt(m.getAttribute("endRow"));
    const ec = parseInt(m.getAttribute("endCol"));
    const rs = er - sr + 1;
    const cs = ec - sc + 1;
    if (rs > 1 || cs > 1) spanMap[`${sr},${sc}`] = { rowspan: rs, colspan: cs };
    for (let r = sr; r <= er; r++) {
      for (let c = sc; c <= ec; c++) {
        if (r !== sr || c !== sc) skipSet.add(`${r},${c}`);
      }
    }
  });

  // Column widths (px ≈ width * 7 for default font)
  const colWidths = {};
  sheetElem.querySelectorAll("colWidths > col").forEach((c) => {
    colWidths[parseInt(c.getAttribute("index"))] =
      Math.max(40, parseFloat(c.getAttribute("width") || 8.43) * 7);
  });

  // Row heights (pt → px ≈ * 1.33)
  const rowHeights = {};
  sheetElem.querySelectorAll("rowHeights > row").forEach((r) => {
    rowHeights[parseInt(r.getAttribute("index"))] =
      Math.max(16, parseFloat(r.getAttribute("height") || 15) * 1.33);
  });

  // Build table DOM
  const table = document.createElement("table");
  table.className = "spreadsheet";

  // colgroup
  const cg = document.createElement("colgroup");
  const rowNumCol = document.createElement("col");
  rowNumCol.style.width = "36px";
  cg.appendChild(rowNumCol);
  for (let c = 1; c <= maxCol; c++) {
    const col = document.createElement("col");
    col.style.width = `${colWidths[c] || 56}px`;
    cg.appendChild(col);
  }
  table.appendChild(cg);

  // thead (column letters)
  const thead = document.createElement("thead");
  const htr = document.createElement("tr");
  htr.appendChild(document.createElement("th")); // corner
  for (let c = 1; c <= maxCol; c++) {
    const th = document.createElement("th");
    th.textContent = colLetter(c);
    htr.appendChild(th);
  }
  thead.appendChild(htr);
  table.appendChild(thead);

  // tbody
  const tbody = document.createElement("tbody");
  for (let r = 1; r <= maxRow; r++) {
    const tr = document.createElement("tr");
    const rh = rowHeights[r];
    if (rh) tr.style.height = `${rh}px`;

    // Row number cell
    const rnTh = document.createElement("th");
    rnTh.className = "row-num";
    rnTh.textContent = r;
    tr.appendChild(rnTh);

    for (let c = 1; c <= maxCol; c++) {
      const key = `${r},${c}`;
      if (skipSet.has(key)) continue;

      const td = document.createElement("td");
      td.dataset.row = r;
      td.dataset.col = c;

      const span = spanMap[key];
      if (span) {
        if (span.rowspan > 1) td.rowSpan = span.rowspan;
        if (span.colspan > 1) td.colSpan = span.colspan;
      }

      const cellXml = cellMap[key];
      if (cellXml) {
        const rawValue = cellXml.getAttribute("value") || "";
        td.textContent = rawValue;
        applyStyleToTd(td, cellXml.querySelector("format"));
      }

      // Editable on click in edit mode
      td.addEventListener("click", onCellClick);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);

  // Wrap + images
  const wrapper = document.createElement("div");
  wrapper.style.cssText = "display:inline-block;min-width:100%;";
  wrapper.appendChild(table);

  // Append embedded images after the table (simplified positioning)
  const imgList = sheetElem.querySelectorAll("images > image");
  if (imgList.length > 0) {
    imgList.forEach((imgXml) => {
      const imgData = imgXml.getAttribute("data");
      if (!imgData) return;
      const fmt = imgXml.getAttribute("format") || "png";
      const imgEl = document.createElement("img");
      imgEl.src = `data:image/${fmt};base64,${imgData}`;
      imgEl.alt = "Embedded image";
      const w = imgXml.getAttribute("width");
      const h = imgXml.getAttribute("height");
      if (w) imgEl.style.maxWidth = `${w}px`;
      if (h) imgEl.style.maxHeight = `${h}px`;
      imgEl.style.cssText += "display:block;margin:8px 0;";
      wrapper.appendChild(imgEl);
    });
  }

  return wrapper;
}

// ── Apply cell style ──────────────────────────────────────────────────────
const BORDER_CSS = {
  thin: "1px solid #999",
  medium: "2px solid #777",
  thick: "3px solid #555",
  dashed: "1px dashed #999",
  dotted: "1px dotted #999",
  double: "3px double #999",
  hair: "1px solid #ccc",
  mediumDashed: "2px dashed #777",
  dashDot: "1px dashed #999",
  mediumDashDot: "2px dashed #777",
  dashDotDot: "1px dotted #999",
  mediumDashDotDot: "2px dotted #777",
  slantDashDot: "1px dashed #999",
};

function applyStyleToTd(td, fmtEl) {
  if (!fmtEl) return;
  const s = td.style;

  if (fmtEl.getAttribute("bold") === "true")   s.fontWeight = "bold";
  if (fmtEl.getAttribute("italic") === "true") s.fontStyle  = "italic";
  if (fmtEl.getAttribute("underline") === "true") s.textDecoration = "underline";

  const fontSize = fmtEl.getAttribute("fontSize");
  if (fontSize) s.fontSize = `${fontSize}pt`;

  const fontName = fmtEl.getAttribute("fontName");
  if (fontName) s.fontFamily = `"${fontName}", sans-serif`;

  const fontColor = fmtEl.getAttribute("fontColor");
  if (fontColor) s.color = argbToHex(fontColor);

  const bgColor = fmtEl.getAttribute("bgColor");
  if (bgColor) s.backgroundColor = argbToHex(bgColor);

  const halign = fmtEl.getAttribute("halign");
  if (halign) s.textAlign = halign;

  const valign = fmtEl.getAttribute("valign");
  if (valign === "center") s.verticalAlign = "middle";
  else if (valign === "top") s.verticalAlign = "top";
  else if (valign === "bottom") s.verticalAlign = "bottom";

  if (fmtEl.getAttribute("wrapText") === "true") s.whiteSpace = "pre-wrap";

  ["Top", "Bottom", "Left", "Right"].forEach((side) => {
    const bs = fmtEl.getAttribute(`border${side}`);
    if (bs && bs !== "none") {
      const css = BORDER_CSS[bs] || "1px solid #999";
      s[`border${side}`] = css;
    }
  });
}

function argbToHex(argb) {
  if (!argb) return "";
  const s = argb.replace("#", "");
  if (s.length === 8) return `#${s.slice(2)}`; // strip alpha
  if (s.length === 6) return `#${s}`;
  return "";
}

// ── Edit mode ─────────────────────────────────────────────────────────────
function onCellClick(e) {
  if (!state.editMode) return;
  const td = e.currentTarget;
  if (td.classList.contains("editing")) return;

  td.classList.add("editing");
  const original = td.textContent;
  const input = document.createElement("input");
  input.type = "text";
  input.value = original;
  td.textContent = "";
  td.appendChild(input);
  input.focus();

  function commit() {
    const newVal = input.value;
    td.classList.remove("editing");
    td.textContent = newVal;
    if (newVal !== original) {
      const key = `${td.dataset.row},${td.dataset.col}`;
      state.modifiedCells[key] = newVal;
      td.dataset.modified = "true";
    }
    applyModifiedCellsToXml();
  }

  input.addEventListener("blur", commit);
  input.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter") { ev.preventDefault(); commit(); }
    if (ev.key === "Escape") { input.value = original; commit(); }
  });
}

function applyModifiedCellsToXml() {
  if (Object.keys(state.modifiedCells).length === 0) return;
  let xml = state.xmlText;
  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xml, "application/xml");
    for (const [key, val] of Object.entries(state.modifiedCells)) {
      const [row, col] = key.split(",");
      const cells = doc.querySelectorAll(
        `cell[row="${row}"][col="${col}"]`
      );
      cells.forEach((c) => c.setAttribute("value", val));
    }
    const serializer = new XMLSerializer();
    state.xmlText = serializer.serializeToString(doc);
    if (state.cmEditor) state.cmEditor.setValue(state.xmlText);
  } catch (e) {
    console.error("XML更新エラー:", e);
  }
}

function enterEditMode() {
  state.editMode = true;
  $("edit-toggle-btn").textContent = "👁️ 表示モード";
  $("edit-toggle-btn").classList.add("btn-primary");
  $("edit-toggle-btn").classList.remove("btn-secondary");
  $("save-btn").style.display = "inline-block";
  // Make cells look clickable
  document.querySelectorAll(".spreadsheet td").forEach((td) => {
    td.style.cursor = "cell";
    td.title = "クリックして編集";
  });
}

function leaveEditMode() {
  state.editMode = false;
  // Commit any open edit
  document.querySelectorAll(".spreadsheet td.editing").forEach((td) => {
    const input = td.querySelector("input");
    if (input) {
      td.textContent = input.value;
      td.classList.remove("editing");
    }
  });
  $("edit-toggle-btn").textContent = "✏️ 編集モード";
  $("edit-toggle-btn").classList.remove("btn-primary");
  $("edit-toggle-btn").classList.add("btn-secondary");
  $("save-btn").style.display = "none";
  document.querySelectorAll(".spreadsheet td").forEach((td) => {
    td.style.cursor = "";
    td.title = "";
  });
}

// ── Save ──────────────────────────────────────────────────────────────────
async function saveDocument() {
  const d = state.currentDoc;
  if (!d) return;

  // If XML pane is active, use CodeMirror content
  if (state.currentTab === "xml" && state.cmEditor) {
    state.xmlText = state.cmEditor.getValue();
  } else {
    applyModifiedCellsToXml();
  }

  const btn = $("save-btn");
  btn.textContent = "保存中…";
  btn.disabled = true;

  try {
    const resp = await fetch(
      `/api/document/${enc(d.team)}/${enc(d.code)}/${enc(d.filename)}/save`,
      { method: "POST", body: state.xmlText, headers: { "Content-Type": "application/xml" } }
    );
    const result = await resp.json();
    if (result.error) throw new Error(result.error);
    btn.textContent = "✅ 保存済み";
    state.modifiedCells = {};
    document.querySelectorAll(".spreadsheet td[data-modified]").forEach((td) =>
      td.removeAttribute("data-modified")
    );
    setTimeout(() => {
      btn.textContent = "💾 保存";
      btn.disabled = false;
    }, 2000);
  } catch (e) {
    alert("保存エラー: " + e.message);
    btn.textContent = "💾 保存";
    btn.disabled = false;
  }
}

// ── Export ────────────────────────────────────────────────────────────────
function exportDocument() {
  const d = state.currentDoc;
  if (!d) return;
  window.location.href = `/api/document/${enc(d.team)}/${enc(d.code)}/${enc(d.filename)}/export`;
}

// ── Tab switching ─────────────────────────────────────────────────────────
function showPane(tab) {
  state.currentTab = tab;
  const viewPane = $("view-pane");
  const xmlPane  = $("xml-pane");
  $("view-tab-btn").classList.toggle("active", tab === "view");
  $("xml-tab-btn").classList.toggle("active",  tab === "xml");

  if (tab === "view") {
    viewPane.style.display = "flex";
    xmlPane.style.display  = "none";
  } else {
    viewPane.style.display = "none";
    xmlPane.style.display  = "flex";
    // Initialise CodeMirror lazily
    if (!state.cmEditor) {
      state.cmEditor = CodeMirror.fromTextArea($("xml-editor"), {
        mode: "xml",
        theme: "material",
        lineNumbers: true,
        lineWrapping: false,
        matchBrackets: true,
        readOnly: false,
      });
      // Fill the editor container to available height
      state.cmEditor.getWrapperElement().style.height = "100%";
      state.cmEditor.refresh();
    }
    state.cmEditor.setValue(state.xmlText);
    setTimeout(() => state.cmEditor.refresh(), 50);
  }
}

// ── Search ────────────────────────────────────────────────────────────────
async function doSearch() {
  const q = $("search-input").value.trim();
  if (!q) return;

  $("search-results").style.display = "flex";
  $("welcome").style.display = "none";
  $("doc-toolbar").style.display = "none";
  $("view-pane").style.display  = "none";
  $("xml-pane").style.display   = "none";
  $("search-results-title").textContent = `「${q}」を検索中…`;
  $("search-results-body").innerHTML = loadingHTML();

  try {
    const resp = await fetch(`/api/search?q=${encodeURIComponent(q)}`);
    const data = await resp.json();
    renderSearchResults(data.results, q);
  } catch (e) {
    $("search-results-body").innerHTML = errorHTML(e.message);
  }
}

function renderSearchResults(results, q) {
  $("search-results-title").textContent =
    `「${esc(q)}」の検索結果: ${results.length} 件`;

  if (results.length === 0) {
    $("search-results-body").innerHTML =
      '<p class="search-no-results">該当するドキュメントが見つかりませんでした。</p>';
    return;
  }

  const html = results.map((r) => {
    const highlighted = esc(r.value).replace(
      new RegExp(esc(q), "gi"),
      (m) => `<mark style="background:#fff176">${m}</mark>`
    );
    return `
      <div class="search-result-item"
           data-team="${esc(r.team)}" data-code="${esc(r.code)}"
           data-filename="${esc(r.filename)}">
        <div class="search-result-path">
          🏢 ${esc(r.team)} / 📁 ${esc(r.code)} / 📄 ${esc(r.filename)}
        </div>
        <div class="search-result-value">${highlighted}</div>
        <div class="search-result-cell">
          シート: ${esc(r.sheet)} &nbsp;│&nbsp; セル: ${esc(r.cell)}
        </div>
      </div>`;
  });
  $("search-results-body").innerHTML = html.join("");

  // Click → open document
  document.querySelectorAll(".search-result-item").forEach((item) => {
    item.addEventListener("click", () => {
      const { team, code, filename } = item.dataset;
      openDocument(team, code, filename, null);
    });
  });
}

// ── Utility HTML fragments ────────────────────────────────────────────────
function loadingHTML() {
  return '<div class="loading"><div class="spinner"></div> 変換中…</div>';
}
function errorHTML(msg) {
  return `<div class="error-banner">⚠️ ${esc(msg)}</div>`;
}

// ── Event wiring ──────────────────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", () => {
  loadStructure();

  $("refresh-btn").addEventListener("click", loadStructure);

  $("search-btn").addEventListener("click", doSearch);
  $("search-input").addEventListener("keydown", (e) => {
    if (e.key === "Enter") doSearch();
  });

  $("search-close-btn").addEventListener("click", () => {
    $("search-results").style.display = "none";
    if (state.currentDoc) {
      $("doc-toolbar").style.display = "flex";
      showPane(state.currentTab);
    } else {
      $("welcome").style.display = "flex";
    }
  });

  $("view-tab-btn").addEventListener("click", () => showPane("view"));
  $("xml-tab-btn").addEventListener("click", () => showPane("xml"));

  $("edit-toggle-btn").addEventListener("click", () => {
    if (state.editMode) leaveEditMode(); else enterEditMode();
  });

  $("save-btn").addEventListener("click", saveDocument);
  $("export-btn").addEventListener("click", exportDocument);
});
