"use strict";

// 1. 대분류 카테고리 (D, H 구분)
const CATEGORIES = ["레미콘", "거푸집", "철근D", "철근H", "기타"];

const state = {
  rawItems: [], 
  dongs: [], 
  floors: [], 
  data: {},       // 구조: { 동: { 아이템: { 층: 수량 } } } -> 기존 성공 로직 구조
  mappings: [], 
  ready: false
};

const $ = (id) => document.getElementById(id);

// 2. 층 정렬 로직 (기존 성공 로직)
function floorSorter(a, b) {
  const getRank = (name) => {
    const s = String(name).toUpperCase().trim();
    if (s.startsWith('B')) return 1000 - (parseInt(s.replace('B', '')) || 0);
    if (s === 'FT' || s === '기초') return 2000;
    if (s.endsWith('F') || /^\d+$/.test(s)) return 3000 + (parseInt(s.replace('F', '')) || 0);
    if (s.startsWith('PH')) return 4000 + (parseInt(s.replace('PH', '')) || 0);
    return 5000;
  };
  return getRank(a) - getRank(b);
}

// 3. 아이템 분류 예측 (새로운 D, H 구분 로직 적용)
function predictCategory(name) {
  const s = String(name).toUpperCase().replace(/\s+/g, "");
  if (s.includes("MPA") || /\d+-\d+-\d+/.test(s) || (/^\d+$/.test(s) && parseInt(s) >= 150)) return "레미콘";
  if (["폼","FORM","회","알폼","갱폼","합벽","거푸집"].some(k => s.includes(k)) || (/[가-힣]/.test(s) && !s.includes("철근"))) return "거푸집";
  if (/(HD|SD|H)\d+/.test(s)) return "철근H";
  if (/D\d+/.test(s)) return "철근D";
  return "기타";
}

// 4. 🌟 원본 층별집계표 파싱 (기존 성공 로직 100% 복구) 🌟
$("file-main").onchange = async (e) => {
  const files = e.target.files;
  if (!files.length) return;

  state.rawItems = []; state.dongs = []; state.floors = []; state.data = {}; state.mappings = [];

  for (const file of files) {
    const ab = await file.arrayBuffer();
    const wb = XLSX.read(ab, { type: 'array' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    parseRows(rows);
  }
  
  state.floors.sort(floorSorter);
  buildMapping();
  renderMappingTable();
  
  state.ready = true;
  const fileList = $("file-list");
  if(fileList) fileList.innerHTML = `<span style="color:#27ae60;">✅ 데이터 로드 성공! 2번 탭을 확인하세요.</span>`;
  alert("데이터 분석이 완료되었습니다.");
};

// 원본 엑셀의 기괴한 구조(동 명 : [101동])를 읽어내는 특수 파서
function parseRows(rows) {
  let curDong = "", lastF = "";
  const r3 = rows[2] || [], r4 = rows[3] || [];
  
  for (let r = 4; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.length === 0) continue;

    const txt = row.join("|");
    const m = txt.match(/동\s*명\s*:\s*\[([^\]]+)\]/);
    if (m) { 
      const raw = m[1].trim();
      if (raw) {
        curDong = raw;
        if(!state.dongs.includes(curDong)) state.dongs.push(curDong);
        state.data[curDong] = state.data[curDong] || {};
      }
      lastF = ""; continue; 
    }
    if (!curDong) continue;

    const fRaw = String(row[0]).trim();
    if (fRaw === "층" || fRaw.includes("계") || fRaw.includes("합") || fRaw.includes("공사명")) {
      lastF = ""; continue;
    }

    if (fRaw !== "") {
      lastF = /^\d+$/.test(fRaw) ? fRaw + "F" : fRaw;
      if (!state.floors.includes(lastF)) state.floors.push(lastF);
    }
    
    if (!lastF) continue;

    for (let c = 1; c < row.length; c++) {
      const val = parseFloat(String(row[c]).replace(/,/g, ""));
      if (isNaN(val) || val === 0) continue;

      let name = (fRaw !== "") ? String(r3[c] || "").trim() : String(r4[c] || "").trim();
      if (!name) name = String(r3[c] || r4[c] || "").trim();
      if (!name) continue;

      if (!state.rawItems.includes(name)) state.rawItems.push(name);
      state.data[curDong][name] = state.data[curDong][name] || {};
      state.data[curDong][name][lastF] = (state.data[curDong][name][lastF] || 0) + val;
    }
  }
}

// 5. 매핑 데이터 생성 및 렌더링
function buildMapping() {
  state.mappings = state.rawItems.map(item => ({
    original: item,
    displayName: item,
    category: predictCategory(item)
  }));
}

function renderMappingTable() {
  const tbody = $("mapping-body");
  if (!tbody) return;
  tbody.innerHTML = state.mappings.map((m, idx) => `
    <tr>
      <td>${m.original}</td>
      <td>
        <select onchange="state.mappings[${idx}].category = this.value" class="input">
          ${CATEGORIES.map(cat => `<option value="${cat}" ${m.category === cat ? 'selected' : ''}>${cat}</option>`).join('')}
        </select>
      </td>
      <td>
        <input type="text" value="${m.displayName}" onchange="state.mappings[${idx}].displayName = this.value" class="input" style="width:100%;" />
      </td>
    </tr>
  `).join('');
}

// 6. 🌟 새로운 UI: 통합 수량표 (4행 1세트) 렌더링 🌟
function renderMainTable() {
  const dong = $("filter-dong").value;
  const head = $("table-head");
  const body = $("table-body");
  if (!dong || !state.ready || !head || !body) return;

  const groups = {
    "레미콘": state.mappings.filter(m => m.category === "레미콘"),
    "거푸집": state.mappings.filter(m => m.category === "거푸집"),
    "철근D": state.mappings.filter(m => m.category === "철근D"),
    "철근H": state.mappings.filter(m => m.category === "철근H")
  };

  // 1) 헤더 렌더링
  let h1 = `<tr><th rowspan="2">층</th><th rowspan="2">구분</th>`;
  let h2 = `<tr>`;
  
  ["레미콘", "거푸집", "철근D", "철근H"].forEach(cat => {
    const cols = groups[cat];
    h1 += `<th colspan="${cols.length || 1}">${cat}</th>`;
    if (cols.length === 0) h2 += `<th>-</th>`;
    else cols.forEach(m => h2 += `<th>${m.displayName}</th>`);
  });
  head.innerHTML = h1 + `</tr>` + h2 + `</tr>`;

  // 2) 바디 렌더링 (층별로 4줄씩 생성)
  let bHtml = "";
  state.floors.forEach(f => {
    ["레미콘", "거푸집", "철근D", "철근H"].forEach((rowCat, idx) => {
      let row = `<tr class="${idx === 3 ? 'row-group-end' : ''}">`;
      
      if (idx === 0) {
        row += `<td rowspan="4" style="text-align:center; font-weight:bold; background:#f1f3f5; vertical-align:middle;">${f}</td>`;
      }
      row += `<td style="text-align:center; font-weight:bold; background:#fff;">${rowCat}</td>`;

      ["레미콘", "거푸집", "철근D", "철근H"].forEach(colCat => {
        const targetItems = groups[colCat];
        if (targetItems.length === 0) row += `<td>-</td>`;
        else {
          targetItems.forEach(m => {
            // ★ 파싱된 구조에 맞게 데이터 호출: state.data[dong][item][floor]
            const val = (rowCat === colCat) ? (state.data[dong]?.[m.original]?.[f] || 0) : 0;
            const cls = val === 0 ? "val-zero" : "val-active";
            row += `<td style="text-align:right;" class="${cls}">${val === 0 ? '-' : val.toLocaleString(undefined, {minimumFractionDigits:2})}</td>`;
          });
        }
      });
      row += `</tr>`;
      bHtml += row;
    });
  });
  body.innerHTML = bHtml;
}

// 7. 탭 및 필터 이벤트 바인딩
document.querySelectorAll(".tab").forEach(t => {
  t.onclick = () => {
    document.querySelectorAll(".tab, .tab-panel").forEach(el => el.classList.remove("is-active"));
    t.classList.add("is-active");
    
    // index.html의 ID 구조에 대응
    let targetId = t.dataset.tab;
    if (targetId === "view") targetId = "tab-view";
    else targetId = targetId + "-tab";
    
    const targetEl = $(targetId);
    if(targetEl) targetEl.classList.add("is-active");
    
    if (t.dataset.tab === "view" && state.ready) {
      const dongSelect = $("filter-dong");
      const currentVal = dongSelect.value;
      dongSelect.innerHTML = state.dongs.map(d => `<option value="${d}">${d}</option>`).join('');
      if (currentVal && state.dongs.includes(currentVal)) dongSelect.value = currentVal;
      renderMainTable();
    }
  };
});

const applyBtn = $("btn-apply-mapping");
if(applyBtn) applyBtn.onclick = () => alert("매핑이 적용되었습니다. 3단계 탭에서 결과를 확인하세요.");

const filterDong = $("filter-dong");
if(filterDong) filterDong.onchange = renderMainTable;

// 8. 엑셀 다운로드 (화면의 4행 구조를 그대로 엑셀로 추출)
const excelBtn = $("btn-excel");
if(excelBtn) {
  excelBtn.onclick = async () => {
    const dong = $("filter-dong").value;
    if (!dong || !state.ready) return;

    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet(dong, { views: [{ state: 'frozen', xSplit: 2, ySplit: 2 }] });

    const groups = {
      "레미콘": state.mappings.filter(m => m.category === "레미콘"),
      "거푸집": state.mappings.filter(m => m.category === "거푸집"),
      "철근D": state.mappings.filter(m => m.category === "철근D"),
      "철근H": state.mappings.filter(m => m.category === "철근H")
    };

    // 헤더 쓰기
    const h1 = ["층", "구분"], h2 = ["", ""];
    ["레미콘", "거푸집", "철근D", "철근H"].forEach(cat => {
      const cols = groups[cat];
      if(cols.length === 0) { h1.push(cat); h2.push("-"); }
      else cols.forEach((m, i) => { h1.push(i === 0 ? cat : ""); h2.push(m.displayName); });
    });
    
    const row1 = ws.addRow(h1);
    const row2 = ws.addRow(h2);
    
    // 헤더 스타일
    [row1, row2].forEach(r => {
      r.eachCell(c => {
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCE6F1' } };
        c.font = { bold: true };
        c.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
        c.alignment = { vertical: 'middle', horizontal: 'center' };
      });
    });

    // 데이터 쓰기
    state.floors.forEach((f, fIdx) => {
      ["레미콘", "거푸집", "철근D", "철근H"].forEach(rowCat => {
        const rowData = [f, rowCat];
        ["레미콘", "거푸집", "철근D", "철근H"].forEach(colCat => {
          const cols = groups[colCat];
          if(cols.length === 0) rowData.push(0);
          else cols.forEach(m => rowData.push(rowCat === colCat ? (state.data[dong]?.[m.original]?.[f] || 0) : 0));
        });
        
        const row = ws.addRow(rowData);
        row.eachCell((c, colNumber) => {
          c.border = { top:{style:'thin', color:{argb:'FFBFBFBF'}}, left:{style:'thin', color:{argb:'FFBFBFBF'}}, bottom:{style:'thin', color:{argb:'FFBFBFBF'}}, right:{style:'thin', color:{argb:'FFBFBFBF'}} };
          if(colNumber <= 2) c.alignment = { vertical: 'middle', horizontal: 'center' };
          else { c.alignment = { horizontal: 'right' }; c.numFmt = '#,##0.00'; }
        });
      });

      // 층 단위 4줄 셀 병합
      const startR = 3 + (fIdx * 4);
      ws.mergeCells(startR, 1, startR + 3, 1);
    });

    // 열 넓이 자동 조정
    ws.columns.forEach((col, i) => col.width = i < 2 ? 12 : 15);

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), `QS_비교표_${dong}.xlsx`);
  };
}
