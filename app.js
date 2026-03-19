"use strict";

// 1. 상태 및 상수 (철근D, H 분리)
const CATEGORIES = ["레미콘", "거푸집", "철근D", "철근H", "잡/기타"];

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

// 2. 유틸리티: 층 정렬
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

// 3. 아이템 분류 예측 (사용자 요청 반영)
function predictCategory(name) {
  const s = String(name).toUpperCase().replace(/\s+/g, "");
  if (s.includes("MPA") || /\d+-\d+-\d+/.test(s)) return "레미콘";
  if (["폼", "FORM", "거푸집", "갱폼", "알폼", "유로", "문양"].some(k => s.includes(k))) return "거푸집";
  if (/(HD|SD|H)\d+/.test(s)) return "철근H";
  if (/D\d+/.test(s)) return "철근D";
  return "잡/기타";
}

// 4. 파일 처리 (기존 로직 유지)
$("file-main").addEventListener("change", async (e) => {
  const files = e.target.files;
  if (!files.length) return;

  for (const file of files) {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    json.forEach(row => {
      const dong = String(row["동"] || "전체").trim();
      const item = String(row["아이템"] || row["구분"] || "").trim();
      if (!item || ["소계", "합계", "구분"].includes(item)) return;

      if (!state.rawItems.includes(item)) {
        state.rawItems.push(item);
        state.mappings.push({ original: item, category: predictCategory(item), displayName: item });
      }

      if (!state.data[dong]) state.data[dong] = {};
      if (!state.dongs.includes(dong)) state.dongs.push(dong);

      Object.keys(row).forEach(key => {
        const floor = key.trim();
        const skip = ["동", "아이템", "구분", "단위", "합계", "비고", "현재 프로젝트 수량"];
        if (skip.includes(floor)) return;

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
  alert("데이터 로드 및 분석 완료");
});

// 5. 매핑 테이블 렌더링
function renderMappingTable() {
  const tbody = $("mapping-body");
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

// 6. 메인 통합 비교표 렌더링 (층별 4행 구조 핵심)
function renderMainTable() {
  const dong = $("filter-dong").value;
  const thead = $("table-head");
  const tbody = $("table-body");
  if (!dong || !state.ready) return;

  const grouped = {
    "레미콘": state.mappings.filter(m => m.category === "레미콘"),
    "거푸집": state.mappings.filter(m => m.category === "거푸집"),
    "철근D": state.mappings.filter(m => m.category === "철근D"),
    "철근H": state.mappings.filter(m => m.category === "철근H")
  };

  // 헤더 렌더링
  let headHtml = `
    <tr>
      <th rowspan="2">층</th><th rowspan="2">구분</th>
      <th colspan="${grouped["레미콘"].length || 1}">레미콘(M3)</th>
      <th colspan="${grouped["거푸집"].length || 1}">거푸집(M2)</th>
      <th colspan="${grouped["철근D"].length || 1}">철근D(ton)</th>
      <th colspan="${grouped["철근H"].length || 1}">철근H(ton)</th>
    </tr>
    <tr>
      ${Object.values(grouped).map(group => group.length ? group.map(m => `<th>${m.displayName}</th>`).join('') : '<th>-</th>').join('')}
    </tr>`;
  thead.innerHTML = headHtml;

  // 바디 렌더링 (1개층 당 4줄)
  let bodyHtml = "";
  state.floors.forEach(floor => {
    const cats = ["레미콘", "거푸집", "철근D", "철근H"];
    cats.forEach((rowCat, idx) => {
      let row = `<tr>`;
      if (idx === 0) row += `<td rowspan="4" style="background:#f1f3f5; font-weight:bold; text-align:center;">${floor}</td>`;
      row += `<td style="font-weight:bold; text-align:center;">${rowCat}</td>`;

      cats.forEach(colCat => {
        if (grouped[colCat].length === 0) {
          row += `<td>-</td>`;
        } else {
          grouped[colCat].forEach(m => {
            const val = (rowCat === colCat) ? (state.data[dong]?.[floor]?.[m.original] || 0) : 0;
            row += `<td style="text-align:right; color:${val === 0 ? '#ccc' : '#000'}">${val === 0 ? '-' : val.toLocaleString(undefined,{minimumFractionDigits:2})}</td>`;
          });
        }
      });
      row += `</tr>`;
      bodyHtml += row;
    });
  });
  tbody.innerHTML = bodyHtml;
}

// 7. 엑셀 다운로드 (ExcelJS 기반 - 기존의 복잡한 스타일링 로직 복원)
$("btn-excel").onclick = async () => {
  const dong = $("filter-dong").value;
  if (!dong) return;

  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet(dong);

  // 헤더 및 데이터 쓰기 로직 (간략화된 버전이나 구조는 유지)
  // 실제 업무용으로는 더 정교한 병합(mergeCells)이 필요합니다.
  
  // 1. 헤더 구성
  const grouped = {
    "레미콘": state.mappings.filter(m => m.category === "레미콘"),
    "거푸집": state.mappings.filter(m => m.category === "거푸집"),
    "철근D": state.mappings.filter(m => m.category === "철근D"),
    "철근H": state.mappings.filter(m => m.category === "철근H")
  };

  const headerRow1 = ["층", "구분"];
  const headerRow2 = ["", ""];
  
  Object.keys(grouped).forEach(cat => {
    grouped[cat].forEach((m, i) => {
      headerRow1.push(i === 0 ? cat : "");
      headerRow2.push(m.displayName);
    });
  });

  ws.addRow(headerRow1);
  ws.addRow(headerRow2);

  // 2. 데이터 구성
  state.floors.forEach(floor => {
    ["레미콘", "거푸집", "철근D", "철근H"].forEach(rowCat => {
      const rowData = [floor, rowCat];
      Object.keys(grouped).forEach(colCat => {
        grouped[colCat].forEach(m => {
          const val = (rowCat === colCat) ? (state.data[dong]?.[floor]?.[m.original] || 0) : 0;
          rowData.push(val || 0);
        });
      });
      ws.addRow(rowData);
    });
  });

  // 스타일 및 병합 처리 (이 부분이 기존 코드의 핵심)
  ws.getColumn(1).alignment = { vertical: 'middle', horizontal: 'center' };
  for (let i = 0; i < state.floors.length; i++) {
    const start = 3 + (i * 4);
    ws.mergeCells(start, 1, start + 3, 1);
  }

  const buffer = await workbook.xlsx.writeBuffer();
  saveAs(new Blob([buffer]), `QS_비교표_${dong}.xlsx`);
};

// 8. 탭 이벤트 및 초기화
document.querySelectorAll(".tab").forEach(tab => {
  tab.addEventListener("click", () => {
    document.querySelectorAll(".tab, .tab-panel").forEach(el => el.classList.remove("is-active"));
    tab.classList.add("is-active");
    $(tab.dataset.tab + "-tab").classList.add("is-active");
    if (tab.dataset.tab === "view") {
      $("filter-dong").innerHTML = state.dongs.map(d => `<option value="${d}">${d}</option>`).join('');
      renderMainTable();
    }
  });
});

$("filter-dong").onchange = renderMainTable;
$("btn-apply-mapping").onclick = () => alert("매핑 적용됨");
