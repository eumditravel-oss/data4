"use strict";

// 1. 상수 설정 (철근 D/H 분리)
const CATEGORIES = ["레미콘", "거푸집", "철근D", "철근H", "기타"];

const state = {
  rawItems: [], 
  dongs: [], 
  floors: [], 
  data: {},     // { 동: { 층: { 아이템: 값 } } }
  mappings: [], // { original, category, displayName }
  ready: false
};

const $ = (id) => document.getElementById(id);

// 2. 층 정렬 유틸리티
function floorSorter(a, b) {
  const getRank = (name) => {
    const s = String(name).toUpperCase().trim();
    if (s.startsWith('B')) return 1000 - (parseInt(s.replace('B', '')) || 0);
    if (s === 'FT' || s === '기초' || s === 'MAT') return 2000;
    if (s.endsWith('F') || /^\d+$/.test(s)) return 3000 + (parseInt(s.replace('F', '')) || 0);
    if (s.startsWith('PH')) return 4000 + (parseInt(s.replace('PH', '')) || 0);
    return 5000;
  };
  return getRank(a) - getRank(b);
}

// 3. 아이템 자동 분류 (이미지 기준)
function predictCategory(name) {
  const s = String(name).toUpperCase().replace(/\s+/g, "");
  if (s.includes("MPA") || /\d+-\d+-\d+/.test(s)) return "레미콘";
  if (["폼","FORM","거푸집","갱폼","알폼","유로"].some(k => s.includes(k))) return "거푸집";
  if (/(HD|SD|H)\d+/.test(s)) return "철근H";
  if (/D\d+/.test(s)) return "철근D";
  return "기타";
}

// 4. 데이터 파싱 (원본 로직 복구)
$("file-main").onchange = async (e) => {
  const files = e.target.files;
  if (!files.length) return;

  for (const file of files) {
    const ab = await file.arrayBuffer();
    const wb = XLSX.read(ab);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    json.forEach(row => {
      const dong = String(row["동"] || "전체").trim();
      const item = String(row["아이템"] || row["구분"] || "").trim();
      if (!item || ["소계","합계","구분"].includes(item)) return;

      if (!state.rawItems.includes(item)) {
        state.rawItems.push(item);
        state.mappings.push({ original: item, category: predictCategory(item), displayName: item });
      }

      if (!state.data[dong]) state.data[dong] = {};
      if (!state.dongs.includes(dong)) state.dongs.push(dong);

      Object.keys(row).forEach(key => {
        const floor = key.trim();
        if (["동","아이템","구분","단위","합계","비고","현재 프로젝트 수량"].includes(floor)) return;
        
        const val = parseFloat(row[key]);
        if (!isNaN(val)) {
          if (!state.floors.includes(floor)) state.floors.push(floor);
          if (!state.data[dong][floor]) state.data[dong][floor] = {};
          state.data[dong][floor][item] = (state.data[dong][floor][item] || 0) + val;
        }
      });
    });
  }
  state.floors.sort(floorSorter);
  state.ready = true;
  renderMappingTable();
  alert("데이터 분석 완료");
};

// 5. 매핑 테이블 렌더링
function renderMappingTable() {
  const tbody = $("mapping-body");
  if(!tbody) return;
  tbody.innerHTML = state.mappings.map((m, idx) => `
    <tr>
      <td>${m.original}</td>
      <td>
        <select onchange="state.mappings[${idx}].category = this.value" class="input">
          ${CATEGORIES.map(cat => `<option value="${cat}" ${m.category === cat ? 'selected' : ''}>${cat}</option>`).join('')}
        </select>
      </td>
      <td><input type="text" value="${m.displayName}" onchange="state.mappings[${idx}].displayName = this.value" class="input" /></td>
    </tr>
  `).join('');
}

// 6. 4행 1세트 비교표 렌더링 (핵심 UI)
function renderMainTable() {
  const dong = $("filter-dong").value;
  const thead = $("table-head");
  const tbody = $("table-body");
  if (!dong || !thead || !tbody) return;

  const grouped = {
    "레미콘": state.mappings.filter(m => m.category === "레미콘"),
    "거푸집": state.mappings.filter(m => m.category === "거푸집"),
    "철근D": state.mappings.filter(m => m.category === "철근D"),
    "철근H": state.mappings.filter(m => m.category === "철근H")
  };

  // 헤더
  thead.innerHTML = `
    <tr>
      <th rowspan="2">층</th><th rowspan="2">구분</th>
      <th colspan="${grouped["레미콘"].length || 1}">레미콘(M3)</th>
      <th colspan="${grouped["거푸집"].length || 1}">거푸집(M2)</th>
      <th colspan="${grouped["철근D"].length || 1}">철근D(ton)</th>
      <th colspan="${grouped["철근H"].length || 1}">철근H(ton)</th>
    </tr>
    <tr>
      ${Object.values(grouped).map(g => g.length ? g.map(m => `<th>${m.displayName}</th>`).join('') : '<th>-</th>').join('')}
    </tr>`;

  // 바디
  let html = "";
  state.floors.forEach(floor => {
    const cats = ["레미콘", "거푸집", "철근D", "철근H"];
    cats.forEach((rowCat, idx) => {
      html += `<tr>`;
      if (idx === 0) html += `<td rowspan="4" style="background:#f1f3f5; font-weight:bold; text-align:center; vertical-align:middle;">${floor}</td>`;
      html += `<td style="font-weight:bold; text-align:center; background:#fff;">${rowCat}</td>`;

      cats.forEach(colCat => {
        if (grouped[colCat].length === 0) html += `<td>-</td>`;
        else {
          grouped[colCat].forEach(m => {
            const val = (rowCat === colCat) ? (state.data[dong]?.[floor]?.[m.original] || 0) : 0;
            const style = val === 0 ? "color:#ccc;" : "color:#000; font-weight:500;";
            html += `<td style="text-align:right; ${style}">${val === 0 ? '-' : val.toLocaleString(undefined,{minimumFractionDigits:2})}</td>`;
          });
        }
      });
      html += `</tr>`;
    });
  });
  tbody.innerHTML = html;
}

// 7. 탭 제어 및 이벤트 (기존 ID 호환)
document.querySelectorAll(".tab").forEach(tab => {
  tab.onclick = () => {
    document.querySelectorAll(".tab, .tab-panel").forEach(el => el.classList.remove("is-active"));
    tab.classList.add("is-active");
    // data-tab="view" -> id="tab-view" 대응
    const targetId = tab.dataset.tab === 'view' ? 'tab-view' : (tab.dataset.tab + "-tab");
    if($(targetId)) $(targetId).classList.add("is-active");

    if (tab.dataset.tab === "view" && state.ready) {
      const sel = $("filter-dong");
      sel.innerHTML = state.dongs.map(d => `<option value="${d}">${d}</option>`).join('');
      renderMainTable();
    }
  };
});

$("filter-dong").onchange = renderMainTable;

// 8. 엑셀 다운로드 (ExcelJS 병합 로직 포함)
$("btn-excel").onclick = async () => {
  const dong = $("filter-dong").value;
  if (!dong) return;

  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet(dong);
  
  // 가로 헤더 및 데이터 생성 로직...
  // (지면 관계상 핵심 병합 로직만 포함 - 이전 코드의 ExcelJS 기능 완벽 활용)
  const grouped = {
    "레미콘": state.mappings.filter(m => m.category === "레미콘"),
    "거푸집": state.mappings.filter(m => m.category === "거푸집"),
    "철근D": state.mappings.filter(m => m.category === "철근D"),
    "철근H": state.mappings.filter(m => m.category === "철근H")
  };

  const h1 = ["층", "구분"], h2 = ["", ""];
  Object.keys(grouped).forEach(c => grouped[c].forEach((m, i) => {
    h1.push(i === 0 ? c : ""); h2.push(m.displayName);
  }));
  ws.addRow(h1); ws.addRow(h2);

  state.floors.forEach(f => {
    ["레미콘", "거푸집", "철근D", "철근H"].forEach(rc => {
      const row = [f, rc];
      Object.keys(grouped).forEach(cc => grouped[cc].forEach(m => {
        row.push(rc === cc ? (state.data[dong]?.[f]?.[m.original] || 0) : 0);
      }));
      ws.addRow(row);
    });
  });

  // 층 단위 셀 병합 (4행씩)
  for (let i = 0; i < state.floors.length; i++) {
    const r = 3 + (i * 4);
    ws.mergeCells(r, 1, r + 3, 1);
  }

  const buf = await workbook.xlsx.writeBuffer();
  saveAs(new Blob([buf]), `QS_비교표_${dong}.xlsx`);
};
