"use strict";

const CATEGORIES = ["콘크리트", "거푸집", "철근", "잡/기타"];
const ANALYSIS_METRICS = {
  "rebar_per_concrete": { label: "레미콘/철근", unit: "Ton/m³", numerator: "철근", denominator: "콘크리트", digits: 4 },
  "rebar_per_area": { label: "면적/철근", unit: "Ton/m²", numerator: "철근", denominator: "면적", digits: 4 },
  "rebar_per_py": { label: "평수/철근", unit: "Ton/Py", numerator: "철근", denominator: "평수", digits: 4 },
  "form_per_area": { label: "거푸집/면적", unit: "m²/m²", numerator: "거푸집", denominator: "면적", digits: 4 },
  "form_per_py": { label: "거푸집/평수", unit: "m²/Py", numerator: "거푸집", denominator: "평수", digits: 4 }
};

const state = {
  rawItems: [],
  dongs: [],
  floors: [],
  data: {},
  mappings: [],
  areas: {},
  ready: false
};

const $ = (id) => document.getElementById(id);

const fmt = (val, digits = 3) => {
  if (val === 0 || !val || isNaN(val)) return "-";
  return Number(val).toLocaleString(undefined, { maximumFractionDigits: digits });
};

const toNumber = (val) => {
  const n = parseFloat(String(val ?? "").replace(/,/g, ""));
  return isNaN(n) ? 0 : n;
};

function floorSorter(a, b) {
  const getRank = (name) => {
    const s = String(name).toUpperCase().trim();
    if (s.startsWith("B")) return 1000 - (parseInt(s.replace("B", "")) || 0);
    if (s === "FT") return 2000;
    if (s.endsWith("F") || /^\d+$/.test(s)) return 3000 + (parseInt(s.replace("F", "")) || 0);
    if (s.startsWith("PH")) return 4000 + (parseInt(s.replace("PH", "")) || 0);
    return 5000;
  };
  return getRank(a) - getRank(b);
}

function predictCategory(name) {
  const s = String(name).toUpperCase().replace(/\s+/g, "");
  if (/(H|D|HD|SD)\d+/.test(s) || s.includes("철근")) return "철근";
  if (s.includes("MPA") || /\d+-\d+-\d+/.test(s) || (/^\d+$/.test(s) && parseInt(s) >= 150)) return "콘크리트";
  if (["폼", "FORM", "회", "알폼", "갱폼", "합벽"].some(k => s.includes(k)) || /[가-힣]/.test(s)) return "거푸집";
  return "잡/기타";
}

function getSortedDongs() {
  return [...state.dongs].sort((a, b) => String(a).localeCompare(String(b), "ko", { numeric: true }));
}

function getSortedFloors() {
  return [...state.floors].sort(floorSorter);
}

function categoryUnit(cat) {
  if (cat === "철근") return "TON";
  if (cat === "콘크리트") return "M3";
  if (cat === "거푸집") return "M2";
  return "-";
}

function categoryLabel(cat) {
  return cat === "콘크리트" ? "레미콘" : cat;
}

function getActiveCategories() {
  const checked = Array.from(document.querySelectorAll(".category-filter:checked")).map(el => el.value);
  return checked.length ? checked : ["콘크리트", "철근", "거푸집", "잡/기타"];
}

function getThresholdSettings() {
  const metric = $("metric-select")?.value || "rebar_per_concrete";
  const min = toNumber($("threshold-min")?.value);
  const max = toNumber($("threshold-max")?.value);
  const useMin = String($("threshold-min")?.value || "").trim() !== "";
  const useMax = String($("threshold-max")?.value || "").trim() !== "";
  const onlyException = $("only-exception")?.checked || false;
  return { metric, min, max, useMin, useMax, onlyException };
}

function isExceptionValue(value, settings) {
  if (!value || isNaN(value)) return false;
  if (settings.useMin && value < settings.min) return true;
  if (settings.useMax && value > settings.max) return true;
  return false;
}

function buildGrouped(dong) {
  const floors = getSortedFloors();
  const dongData = state.data[dong] || {};
  const grouped = {};

  state.mappings.forEach(m => {
    const qByF = dongData[m.original] || {};
    if (Object.keys(qByF).length === 0) return;
    if (!grouped[m.canonical]) grouped[m.canonical] = { category: m.category, floors: {} };
    floors.forEach(f => {
      grouped[m.canonical].floors[f] = (grouped[m.canonical].floors[f] || 0) + (qByF[f] || 0);
    });
  });

  return grouped;
}

function getCategoryFloorSum(dong, cat, floor) {
  const grouped = buildGrouped(dong);
  return Object.keys(grouped)
    .filter(name => grouped[name].category === cat)
    .reduce((sum, name) => sum + (grouped[name].floors[floor] || 0), 0);
}

function getMetricValue(dong, floor, metricKey) {
  const metric = ANALYSIS_METRICS[metricKey];
  if (!metric) return 0;

  let numerator = 0;
  if (metric.numerator === "철근") numerator = getCategoryFloorSum(dong, "철근", floor);
  if (metric.numerator === "거푸집") numerator = getCategoryFloorSum(dong, "거푸집", floor);

  let denominator = 0;
  if (metric.denominator === "콘크리트") denominator = getCategoryFloorSum(dong, "콘크리트", floor);
  if (metric.denominator === "면적") denominator = state.areas[dong]?.[floor] || 0;
  if (metric.denominator === "평수") denominator = (state.areas[dong]?.[floor] || 0) * 0.3025;

  return denominator > 0 ? numerator / denominator : 0;
}

function getExceptionFloorSet() {
  const settings = getThresholdSettings();
  const set = new Set();
  if (!settings.onlyException || (!settings.useMin && !settings.useMax)) return set;

  getSortedDongs().forEach(dong => {
    getSortedFloors().forEach(floor => {
      const value = getMetricValue(dong, floor, settings.metric);
      if (isExceptionValue(value, settings)) set.add(`${dong}||${floor}`);
    });
  });
  return set;
}

function updateAnalysisMeta() {
  const dongs = getSortedDongs();
  const floors = getSortedFloors();
  const settings = getThresholdSettings();
  let exceptionCount = 0;
  let validCount = 0;

  dongs.forEach(dong => {
    floors.forEach(floor => {
      const value = getMetricValue(dong, floor, settings.metric);
      if (value) validCount++;
      if (isExceptionValue(value, settings)) exceptionCount++;
    });
  });

  const metric = ANALYSIS_METRICS[settings.metric];
  const categoryText = getActiveCategories().join(" / ");
  $("analysis-meta").innerHTML = `검토범위: <b>${dongs.length}</b>개 동 · <b>${floors.length}</b>개 층 · 표시분류: <b>${categoryText}</b> · 기준지표: <b>${metric.label}</b> · 기준 초과/미달: <b>${exceptionCount}</b>개 층 / 유효값 <b>${validCount}</b>개`;
}

$("btn-parse").onclick = async () => {
  const files = Array.from($("file-main").files);
  if (files.length === 0) return alert("파일을 먼저 선택해주세요.");

  state.rawItems = [];
  state.dongs = [];
  state.floors = [];
  state.data = {};
  state.areas = {};

  for (const file of files) {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1, defval: "" });
    parseRows(rows);
  }

  buildMapping();
  renderMapping();
  switchTab("mapping");
};

function parseRows(rows) {
  let curDong = "";
  let lastF = "";
  const r3 = rows[2] || [];
  const r4 = rows[3] || [];

  for (let r = 4; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.length === 0) continue;

    const txt = row.join("|");
    const m = txt.match(/동\s*명\s*:\s*\[([^\]]+)\]/);
    if (m) {
      const raw = m[1].trim();
      if (raw) {
        curDong = raw;
        if (!state.dongs.includes(curDong)) state.dongs.push(curDong);
        state.data[curDong] = state.data[curDong] || {};
      }
      lastF = "";
      continue;
    }
    if (!curDong) continue;

    const fRaw = String(row[0]).trim();
    if (fRaw === "층" || fRaw.includes("계") || fRaw.includes("합") || fRaw.includes("공사명")) {
      lastF = "";
      continue;
    }

    if (fRaw !== "") {
      lastF = /^\d+$/.test(fRaw) ? `${fRaw}F` : fRaw;
      if (!state.floors.includes(lastF)) state.floors.push(lastF);
    }

    if (!lastF) continue;

    for (let c = 1; c < row.length; c++) {
      const val = toNumber(row[c]);
      if (val === 0) continue;

      let name = (fRaw !== "") ? String(r3[c] || "").trim() : String(r4[c] || "").trim();
      if (!name) name = String(r3[c] || r4[c] || "").trim();
      if (!name) continue;

      if (!state.rawItems.includes(name)) state.rawItems.push(name);
      state.data[curDong][name] = state.data[curDong][name] || {};
      state.data[curDong][name][lastF] = (state.data[curDong][name][lastF] || 0) + val;
    }
  }
}

function buildMapping() {
  state.mappings = state.rawItems.map((item, idx) => ({
    id: idx,
    original: item,
    canonical: item,
    category: predictCategory(item)
  }));
}

function renderMapping() {
  $("mapping-list").innerHTML = state.mappings.map(m => {
    const catClass = m.category === "잡/기타" ? "etc" : m.category;
    return `
      <div class="item-row cat-${catClass}">
        <div class="col-num">${m.id + 1}</div>
        <div class="col-orig">${m.original}</div>
        <div class="col-edit"><input class="input" value="${m.canonical}" oninput="updateMapping(${m.id}, 'canonical', this.value)" /></div>
        <div class="col-cat"><select class="input" onchange="updateMapping(${m.id}, 'category', this.value)">${CATEGORIES.map(c => `<option value="${c}" ${m.category === c ? "selected" : ""}>${c}</option>`).join("")}</select></div>
      </div>`;
  }).join("");
}

window.updateMapping = (id, f, v) => {
  state.mappings[id][f] = v;
  if (f === "category") renderMapping();
};

$("btn-apply").onclick = () => {
  renderAreaUI();
  switchTab("area");
};

function renderAreaUI() {
  const dongs = getSortedDongs();
  const floors = getSortedFloors();

  $("area-head").innerHTML = `<tr><th>층 명칭</th>${dongs.map(d => `<th>${d}</th>`).join("")}</tr>`;

  let body = "";
  floors.forEach((f, rIdx) => {
    body += `<tr><td style="font-weight:bold; background:#f4f7fd;">${f}</td>`;
    dongs.forEach((d, cIdx) => {
      const val = state.areas[d]?.[f] || "";
      body += `<td><input type="number" class="area-input" data-r="${rIdx}" data-c="${cIdx}" value="${val}" oninput="updateArea('${d}', '${f}', this.value)" onkeydown="handleAreaNav(event, ${rIdx}, ${cIdx}, ${floors.length}, ${dongs.length})" placeholder="-" /></td>`;
    });
    body += "</tr>";
  });
  $("area-body").innerHTML = body;
}

window.updateArea = (dong, floor, val) => {
  if (!state.areas[dong]) state.areas[dong] = {};
  state.areas[dong][floor] = parseFloat(val) || 0;
};

window.handleAreaNav = (e, r, c, maxR, maxC) => {
  let nr = r;
  let nc = c;
  if (e.key === "ArrowUp") nr = Math.max(0, r - 1);
  else if (e.key === "ArrowDown" || e.key === "Enter") { nr = Math.min(maxR - 1, r + 1); e.preventDefault(); }
  else if (e.key === "ArrowLeft") nc = Math.max(0, c - 1);
  else if (e.key === "ArrowRight") nc = Math.min(maxC - 1, c + 1);
  else return;

  const input = document.querySelector(`.area-input[data-r="${nr}"][data-c="${nc}"]`);
  if (input) { input.focus(); input.select(); }
};

$("btn-download-area").onclick = () => {
  const dongs = getSortedDongs();
  const floors = getSortedFloors();
  const aoa = [["층 명칭", ...dongs]];

  floors.forEach(f => {
    const row = [f];
    dongs.forEach(d => row.push(state.areas[d]?.[f] || ""));
    aoa.push(row);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "면적데이터");
  XLSX.writeFile(wb, "QS_면적입력양식.xlsx");
};

$("file-upload-area").onchange = async (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1, defval: "" });
  if (data.length < 2) return alert("유효한 면적 데이터 양식이 아닙니다.");

  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    const floor = String(data[i][0]).trim();
    if (!floor) continue;
    for (let c = 1; c < headers.length; c++) {
      const dong = String(headers[c]).trim();
      const val = parseFloat(data[i][c]);
      if (!isNaN(val) && state.dongs.includes(dong) && state.floors.includes(floor)) {
        if (!state.areas[dong]) state.areas[dong] = {};
        state.areas[dong][floor] = val;
      }
    }
  }

  renderAreaUI();
  alert("면적 데이터가 성공적으로 불러와졌습니다!");
  e.target.value = "";
};

$("btn-calc-area").onclick = () => {
  state.ready = true;
  $("filter-dong").innerHTML = getSortedDongs().map(d => `<option value="${d}">${d}</option>`).join("");
  renderView();
  switchTab("view");
};

function bindAnalysisControls() {
  const ids = ["filter-dong", "view-mode", "metric-select", "threshold-min", "threshold-max", "only-exception"];
  ids.forEach(id => {
    const el = $(id);
    if (el) el.onchange = renderView;
    if (el && (id === "threshold-min" || id === "threshold-max")) el.oninput = renderView;
  });
  document.querySelectorAll(".category-filter").forEach(el => el.onchange = renderView);
}

document.addEventListener("DOMContentLoaded", bindAnalysisControls);

function renderView() {
  if (!state.ready) return;

  const mode = $("view-mode")?.value || "single-horizontal";
  const selectedDong = $("filter-dong")?.value || getSortedDongs()[0];
  const dongs = mode === "single-horizontal" ? [selectedDong] : getSortedDongs();
  const floors = getSortedFloors();
  const activeCategories = getActiveCategories();
  const exceptionSet = getExceptionFloorSet();
  const settings = getThresholdSettings();

  $("filter-dong").disabled = mode !== "single-horizontal";

  if (mode === "vertical") renderVerticalView(dongs, floors, activeCategories, exceptionSet, settings);
  else if (mode === "metric-summary") renderMetricSummaryView(getSortedDongs(), floors, settings);
  else renderHorizontalView(dongs, floors, activeCategories, exceptionSet, settings);

  updateAnalysisMeta();
}

function renderHorizontalView(dongs, floors, activeCategories, exceptionSet, settings) {
  let headHtml = `<tr><th rowspan="2">동</th><th rowspan="2">아이템</th><th rowspan="2">구분</th><th rowspan="2">단위</th><th colspan="${floors.length}">현재 프로젝트 수량</th><th rowspan="2">합계</th></tr><tr>`;
  floors.forEach(f => headHtml += `<th>${f}</th>`);
  headHtml += "</tr>";
  $("table-head").innerHTML = headHtml;

  let bodyHtml = "";
  dongs.forEach(dong => {
    const grouped = buildGrouped(dong);
    ["콘크리트", "철근", "거푸집", "잡/기타"].forEach(cat => {
      if (!activeCategories.includes(cat)) return;
      const items = Object.keys(grouped).filter(n => grouped[n].category === cat).sort();
      if (items.length === 0) return;

      const visibleFloors = floors.filter(f => !settings.onlyException || exceptionSet.has(`${dong}||${f}`));
      if (settings.onlyException && visibleFloors.length === 0) return;

      let catSum = 0;
      const catClass = cat === "잡/기타" ? "etc" : cat;

      items.forEach(name => {
        const item = grouped[name];
        const total = floors.reduce((s, f) => s + (settings.onlyException && !exceptionSet.has(`${dong}||${f}`) ? 0 : item.floors[f] || 0), 0);
        catSum += total;
        bodyHtml += `<tr class="row-cat-${catClass}"><td>${dong}</td><td>${categoryLabel(cat)}</td><td>${name}</td><td>${categoryUnit(cat)}</td>${floors.map(f => {
          const hidden = settings.onlyException && !exceptionSet.has(`${dong}||${f}`);
          return `<td class="${hidden ? "cell-muted" : ""}">${hidden ? "-" : fmt(item.floors[f], 3)}</td>`;
        }).join("")}<td class="col-total">${fmt(total, 3)}</td></tr>`;
      });

      bodyHtml += `<tr class="row-subtotal"><td>${dong}</td><td colspan="2" style="text-align:right">${categoryLabel(cat)} 합계</td><td>${categoryUnit(cat)}</td>${floors.map(f => {
        const hidden = settings.onlyException && !exceptionSet.has(`${dong}||${f}`);
        const s = items.reduce((sum, n) => sum + (grouped[n].floors[f] || 0), 0);
        return `<td class="${hidden ? "cell-muted" : ""}">${hidden ? "-" : fmt(s, 3)}</td>`;
      }).join("")}<td class="col-total">${fmt(catSum, 3)}</td></tr>`;

      if (cat === "철근") {
        bodyHtml += renderRatioRow(dong, floors, "rebar_per_concrete", exceptionSet, settings);
        bodyHtml += renderRatioRow(dong, floors, "rebar_per_area", exceptionSet, settings);
        bodyHtml += renderRatioRow(dong, floors, "rebar_per_py", exceptionSet, settings);
      }
      if (cat === "거푸집") {
        bodyHtml += renderRatioRow(dong, floors, "form_per_area", exceptionSet, settings);
        bodyHtml += renderRatioRow(dong, floors, "form_per_py", exceptionSet, settings);
      }
    });
  });

  $("table-body").innerHTML = bodyHtml || `<tr><td colspan="${floors.length + 5}">표시할 데이터가 없습니다. 분류 필터 또는 기준치 조건을 확인하세요.</td></tr>`;
}

function renderRatioRow(dong, floors, metricKey, exceptionSet, settings) {
  const metric = ANALYSIS_METRICS[metricKey];
  const isTargetMetric = metricKey === settings.metric;
  const rowClass = isTargetMetric ? "row-ratio row-ratio-target" : "row-ratio";
  let totalNum = 0;
  let totalDiv = 0;
  let html = `<tr class="${rowClass}"><td>${dong}</td><td>지표</td><td>${metric.label}</td><td>${metric.unit}</td>`;

  floors.forEach(f => {
    const hidden = settings.onlyException && !exceptionSet.has(`${dong}||${f}`);
    const value = getMetricValue(dong, f, metricKey);
    const exception = isExceptionValue(value, settings) && metricKey === settings.metric;
    if (!hidden) {
      if (metric.numerator === "철근") totalNum += getCategoryFloorSum(dong, "철근", f);
      if (metric.numerator === "거푸집") totalNum += getCategoryFloorSum(dong, "거푸집", f);
      if (metric.denominator === "콘크리트") totalDiv += getCategoryFloorSum(dong, "콘크리트", f);
      if (metric.denominator === "면적") totalDiv += state.areas[dong]?.[f] || 0;
      if (metric.denominator === "평수") totalDiv += (state.areas[dong]?.[f] || 0) * 0.3025;
    }
    html += `<td class="${hidden ? "cell-muted" : ""} ${exception ? "cell-exception" : ""}">${hidden ? "-" : fmt(value, metric.digits)}</td>`;
  });

  const totalRatio = totalDiv > 0 ? totalNum / totalDiv : 0;
  html += `<td class="col-total">${fmt(totalRatio, metric.digits)}</td></tr>`;
  return html;
}

function renderVerticalView(dongs, floors, activeCategories, exceptionSet, settings) {
  $("table-head").innerHTML = `<tr><th>동</th><th>층</th><th>아이템</th><th>구분</th><th>단위</th><th>수량</th><th>레미콘/철근</th><th>면적/철근</th><th>평수/철근</th><th>거푸집/면적</th><th>거푸집/평수</th></tr>`;
  let bodyHtml = "";

  dongs.forEach(dong => {
    const grouped = buildGrouped(dong);
    floors.forEach(floor => {
      if (settings.onlyException && !exceptionSet.has(`${dong}||${floor}`)) return;
      ["콘크리트", "철근", "거푸집", "잡/기타"].forEach(cat => {
        if (!activeCategories.includes(cat)) return;
        const items = Object.keys(grouped).filter(n => grouped[n].category === cat).sort();
        items.forEach(name => {
          const value = grouped[name].floors[floor] || 0;
          if (!value) return;
          const catClass = cat === "잡/기타" ? "etc" : cat;
          bodyHtml += `<tr class="row-cat-${catClass}"><td>${dong}</td><td>${floor}</td><td>${categoryLabel(cat)}</td><td>${name}</td><td>${categoryUnit(cat)}</td><td>${fmt(value, 3)}</td>${Object.keys(ANALYSIS_METRICS).map(key => {
            const metricValue = getMetricValue(dong, floor, key);
            const exception = key === settings.metric && isExceptionValue(metricValue, settings);
            return `<td class="${exception ? "cell-exception" : ""}">${fmt(metricValue, ANALYSIS_METRICS[key].digits)}</td>`;
          }).join("")}</tr>`;
        });
      });
    });
  });

  $("table-body").innerHTML = bodyHtml || `<tr><td colspan="11">표시할 데이터가 없습니다. 분류 필터 또는 기준치 조건을 확인하세요.</td></tr>`;
}

function renderMetricSummaryView(dongs, floors, settings) {
  $("table-head").innerHTML = `<tr><th>동</th><th>층</th><th>레미콘 합계(M3)</th><th>철근 합계(TON)</th><th>거푸집 합계(M2)</th><th>면적(m²)</th><th>평수(Py)</th><th>레미콘/철근</th><th>면적/철근</th><th>평수/철근</th><th>거푸집/면적</th><th>거푸집/평수</th><th>판정</th></tr>`;
  let bodyHtml = "";

  dongs.forEach(dong => {
    floors.forEach(floor => {
      const concrete = getCategoryFloorSum(dong, "콘크리트", floor);
      const rebar = getCategoryFloorSum(dong, "철근", floor);
      const form = getCategoryFloorSum(dong, "거푸집", floor);
      const area = state.areas[dong]?.[floor] || 0;
      const py = area * 0.3025;
      const metricValue = getMetricValue(dong, floor, settings.metric);
      const exception = isExceptionValue(metricValue, settings);
      if (settings.onlyException && !exception) return;
      if (!concrete && !rebar && !form) return;

      bodyHtml += `<tr class="${exception ? "row-exception" : ""}"><td>${dong}</td><td>${floor}</td><td>${fmt(concrete, 3)}</td><td>${fmt(rebar, 3)}</td><td>${fmt(form, 3)}</td><td>${fmt(area, 3)}</td><td>${fmt(py, 3)}</td>${Object.keys(ANALYSIS_METRICS).map(key => {
        const value = getMetricValue(dong, floor, key);
        const isTarget = key === settings.metric;
        return `<td class="${isTarget && exception ? "cell-exception" : ""}">${fmt(value, ANALYSIS_METRICS[key].digits)}</td>`;
      }).join("")}<td>${exception ? "기준 초과/미달" : "정상"}</td></tr>`;
    });
  });

  $("table-body").innerHTML = bodyHtml || `<tr><td colspan="13">표시할 데이터가 없습니다. 기준치 조건을 확인하세요.</td></tr>`;
}

$("btn-excel").onclick = async () => {
  if (!state.ready) return alert("먼저 분석을 완료해주세요.");
  if (typeof ExcelJS === "undefined") return alert("ExcelJS 라이브러리를 불러오지 못했습니다.");

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("비교양식", {
    views: [{ state: "frozen", ySplit: 4, xSplit: 4, showZeros: false }]
  });

  const floors = getSortedFloors();
  const endCol = 4 + floors.length + 1;
  const maxCol = endCol + 1;

  const cols = [{ width: 10 }, { width: 15 }, { width: 18 }, { width: 10 }];
  floors.forEach(() => cols.push({ width: 9 }));
  cols.push({ width: 13 });
  cols.push({ width: 12 });
  ws.columns = cols;

  const r1 = ws.addRow(["QS 분석용 프로젝트 통합 템플릿"]);
  r1.height = 25;
  ws.mergeCells(1, 1, 2, maxCol);
  const titleCell = ws.getCell(1, 1);
  titleCell.font = { size: 16, bold: true, name: "맑은 고딕" };
  titleCell.alignment = { vertical: "middle", horizontal: "center" };

  const r3Data = ["동", "아이템", "구분", "단위", "현재 프로젝트 수량"];
  for (let i = 0; i < floors.length - 1; i++) r3Data.push("");
  r3Data.push("합계", "비고");

  const r4Data = ["", "", "", ""];
  floors.forEach(f => r4Data.push(f));
  r4Data.push("", "");

  ws.addRow(r3Data).height = 22;
  ws.addRow(r4Data).height = 22;

  ws.mergeCells(3, 1, 4, 1);
  ws.mergeCells(3, 2, 4, 2);
  ws.mergeCells(3, 3, 4, 3);
  ws.mergeCells(3, 4, 4, 4);
  ws.mergeCells(3, 5, 3, endCol - 1);
  ws.mergeCells(3, endCol, 4, endCol);
  ws.mergeCells(3, maxCol, 4, maxCol);

  const borderAll = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
  for (let r = 3; r <= 4; r++) {
    for (let c = 1; c <= maxCol; c++) {
      const cell = ws.getCell(r, c);
      cell.font = { bold: true, size: 10, name: "맑은 고딕", color: { argb: "FFFFFFFF" } };
      cell.alignment = { vertical: "middle", horizontal: "center" };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E78" } };
      cell.border = borderAll;
    }
  }

  const dataBorder = { top: { style: "thin", color: { argb: "FFBFBFBF" } }, left: { style: "thin", color: { argb: "FFBFBFBF" } }, bottom: { style: "thin", color: { argb: "FFBFBFBF" } }, right: { style: "thin", color: { argb: "FFBFBFBF" } } };

  getSortedDongs().forEach(dong => {
    const grouped = buildGrouped(dong);
    const startRow = ws.rowCount + 1;

    ["콘크리트", "철근", "거푸집"].forEach(cat => {
      const items = Object.keys(grouped).filter(n => grouped[n].category === cat).sort();
      if (items.length === 0) return;

      let rowFill = "FFFFFFFF";
      if (cat === "콘크리트") rowFill = "FFEEF4FF";
      else if (cat === "철근") rowFill = "FFF0FCF4";
      else if (cat === "거푸집") rowFill = "FFFFF9EC";

      const catSum = {};
      floors.forEach(f => catSum[f] = 0);
      let totalSum = 0;

      items.forEach(name => {
        const item = grouped[name];
        const rowData = [dong, categoryLabel(cat), name, categoryUnit(cat)];
        let rowTotal = 0;
        floors.forEach(f => {
          rowData.push(item.floors[f] || 0);
          catSum[f] += item.floors[f] || 0;
          rowTotal += item.floors[f] || 0;
        });
        rowData.push(rowTotal, "");
        totalSum += rowTotal;

        const row = ws.addRow(rowData);
        row.height = 18;
        row.outlineLevel = 1;
        for (let c = 1; c <= maxCol; c++) {
          const cell = row.getCell(c);
          cell.border = dataBorder;
          cell.font = { name: "맑은 고딕", size: 10 };
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: rowFill } };
          if (c <= 4) cell.alignment = { vertical: "middle", horizontal: "center" };
          else { cell.alignment = { vertical: "middle", horizontal: "right" }; cell.numFmt = "#,##0.000"; }
        }
      });

      const sumRowData = [dong, categoryLabel(cat), "합계", categoryUnit(cat)];
      floors.forEach(f => sumRowData.push(catSum[f]));
      sumRowData.push(totalSum, "");
      const sumRow = ws.addRow(sumRowData);
      sumRow.height = 18;
      for (let c = 1; c <= maxCol; c++) {
        const cell = sumRow.getCell(c);
        cell.font = { name: "맑은 고딕", size: 10, bold: true };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F2F2" } };
        cell.border = dataBorder;
        if (c <= 4) cell.alignment = { vertical: "middle", horizontal: "center" };
        else { cell.alignment = { vertical: "middle", horizontal: "right" }; cell.numFmt = "#,##0.000"; }
      }
    });

    const endRow = ws.rowCount;
    if (startRow < endRow) {
      ws.mergeCells(startRow, 1, endRow, 1);
      ws.getCell(startRow, 1).alignment = { vertical: "middle", horizontal: "center" };
    }
  });

  const buffer = await wb.xlsx.writeBuffer();
  saveAs(new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), "QS_통합템플릿_리포트.xlsx");
};

function switchTab(id) {
  document.querySelectorAll(".tab, .tab-panel").forEach(el => el.classList.remove("is-active"));
  document.querySelector(`[data-tab="${id}"]`).classList.add("is-active");
  $("tab-" + id).classList.add("is-active");
}

document.querySelectorAll(".tab").forEach(t => t.onclick = () => switchTab(t.dataset.tab));
