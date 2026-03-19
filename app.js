"use strict";

const CATEGORIES = ["레미콘", "거푸집", "철근D", "철근H", "기타"];
const $ = (id) => document.getElementById(id);

let state = {
  data: {},       // { 동: { 층: { 아이템: 수량 } } }
  areas: {},      // { 동: { 층: 면적 } }  <-- 면적 저장 필수 객체
  dongs: [],
  floors: [],
  items: [],      // { original, category, displayName }
  ready: false
};

// 1. 카테고리 자동 분류 (철근D, H 분리)
function predictCategory(name) {
  const n = String(name).toUpperCase().replace(/\s/g, "");
  if (n.includes("MPA") || /\d+-\d+-\d+/.test(n)) return "레미콘";
  if (["폼", "FORM", "거푸집", "갱폼", "알폼", "유로"].some(k => n.includes(k))) return "거푸집";
  if (/(HD|SD|H)\d+/.test(n)) return "철근H";
  if (/D\d+/.test(n)) return "철근D";
  return "기타";
}

// 2. 층 정렬 로직
function sortFloors(arr) {
  const rank = (f) => {
    const s = f.toUpperCase().trim();
    if (s.startsWith('B')) return 1000 - parseInt(s.substring(1));
    if (s === 'FT' || s === '기초') return 2000;
    if (s.endsWith('F')) return 3000 + parseInt(s);
    if (s.startsWith('PH')) return 4000 + parseInt(s.substring(2) || 0);
    return 5000;
  };
  return arr.sort((a, b) => rank(a) - rank(b));
}

// 3. 1단계: 파일 업로드 및 데이터 파싱
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
      if (!item || ["소계", "합계"].includes(item)) return;

      if (!state.items.find(i => i.original === item)) {
        state.items.push({ original: item, category: predictCategory(item), displayName: item });
      }

      if (!state.data[dong]) {
        state.data[dong] = {};
        state.areas[dong] = {}; // 면적 데이터 초기화
      }
      if (!state.dongs.includes(dong)) state.dongs.push(dong);

      Object.keys(row).forEach(key => {
        const skip = ["동", "아이템", "구분", "단위", "합계", "비고"];
        if (skip.includes(key)) return;
        
        const floor = key.trim();
        const val = parseFloat(row[key]);
        if (!isNaN(val)) {
          if (!state.floors.includes(floor)) state.floors.push(floor);
          if (!state.data[dong][floor]) state.data[dong][floor] = {};
          state.data[dong][floor][item] = (state.data[dong][floor][item] || 0) + val;
        }
      });
    });
  }
  state.floors = sortFloors(state.floors);
  state.ready = true;
  $("file-list").innerHTML = "✅ 데이터 분석 완료! 2단계로 넘어가세요.";
  renderMappingTable();
  updateDongSelects();
};

// 동 선택 드롭박스 업데이트
function updateDongSelects() {
  const opts = state.dongs.map(d => `<option value="${d}">${d}</option>`).join('');
  $("filter-dong-area").innerHTML = opts;
  $("filter-dong-view").innerHTML = opts;
  if (state.dongs.length > 0) renderAreaTable();
}

// 4. 2단계: 매핑 렌더링
function renderMappingTable() {
  const tbody = $("mapping-body");
  tbody.innerHTML = state.items.map((m, i) => `
    <tr>
      <td>${m.original}</td>
      <td>
        <select onchange="state.items[${i}].category = this.value" class="input">
          ${CATEGORIES.map(c => `<option value="${c}" ${m.category === c ? 'selected' : ''}>${c}</option>`).join('')}
        </select>
      </td>
      <td><input type="text" value="${m.displayName}" onchange="state.items[${i}].displayName = this.value" class="input"/></td>
    </tr>
  `).join('');
}

// 5. 3단계: 면적 입력 화면 렌더링 (사진 복구 기능)
function renderAreaTable() {
  const dong = $("filter-dong-area").value;
  if (!dong) return;
  const tbody = $("area-body");
  
  tbody.innerHTML = state.floors.map(f => {
    const area = state.areas[dong][f] || 0;
    const py = (area * 0.3025).toFixed(2);
    return `
      <tr>
        <td style="text-align:center; font-weight:bold;">${f}</td>
        <td>
          <input type="number" class="area-input" value="${area || ''}" placeholder="0" 
                 onchange="state.areas['${dong}']['${f}'] = parseFloat(this.value) || 0; renderAreaTable();" />
        </td>
        <td style="text-align:right; color:#e67e22; font-weight:bold;">${py} Py</td>
      </tr>
    `;
  }).join('');
}

$("filter-dong-area").onchange = renderAreaTable;

// 6. 4단계: 통합 비교표 확인 (4행 1세트 + 지표 계산)
function renderMainTable() {
  const dong = $("filter-dong-view").value;
  if (!dong || !state.ready) return;

  const head = $("table-head");
  const body = $("table-body");

  // 대분류별 아이템 분류
  const groups = {
    "레미콘": state.items.filter(i => i.category === "레미콘"),
    "거푸집": state.items.filter(i => i.category === "거푸집"),
    "철근D": state.items.filter(i => i.category === "철근D"),
    "철근H": state.items.filter(i => i.category === "철근H")
  };

  // --- 헤더 생성 ---
  let h1 = `<tr><th rowspan="2" style="width:60px;">층</th><th rowspan="2" style="width:80px;">구분</th>`;
  let h2 = `<tr>`;
  
  // 항목들 가로 나열 (지표 컬럼 추가)
  ["레미콘", "거푸집", "철근D", "철근H"].forEach(cat => {
    const cols = groups[cat];
    h1 += `<th colspan="${cols.length}">${cat} 상세항목</th>`;
    h1 += `<th colspan="4" style="background:#fff4e6;">${cat} 지표 분석</th>`; // 소계, 면적당, 평당, 층면적 등
    
    cols.forEach(i => h2 += `<th>${i.displayName}</th>`);
    h2 += `<th style="background:#fff4e6;">소계</th>`;
    h2 += `<th style="background:#fff4e6;">단위/m²</th>`;
    h2 += `<th style="background:#fff4e6;">단위/Py</th>`;
    h2 += `<th style="background:#fff4e6;">층면적(m²)</th>`;
  });
  head.innerHTML = h1 + `</tr>` + h2 + `</tr>`;

  // --- 바디 생성 (1개 층당 4행) ---
  let bHtml = "";
  state.floors.forEach(f => {
    const area = state.areas[dong][f] || 0;
    const py = area * 0.3025;
    const rowCats = ["레미콘", "거푸집", "철근D", "철근H"];

    rowCats.forEach((rowCat, idx) => {
      let row = `<tr class="${idx === 3 ? 'row-group-end' : ''}">`;
      if (idx === 0) row += `<td rowspan="4" style="vertical-align:middle;">${f}</td>`;
      row += `<td>${rowCat}</td>`;

      rowCats.forEach(colCat => {
        let subTotal = 0;
        
        // 데이터 셀
        groups[colCat].forEach(item => {
          const val = (rowCat === colCat) ? (state.data[dong]?.[f]?.[item.original] || 0) : 0;
          if (rowCat === colCat) subTotal += val;
          const cls = val === 0 ? "val-zero" : "val-active";
          row += `<td style="text-align:right;" class="${cls}">${val === 0 ? '-' : val.toLocaleString(undefined, {minimumFractionDigits:2})}</td>`;
        });

        // 지표 계산 셀 (현재 행 카테고리 === 열 카테고리일 때만 계산)
        if (rowCat === colCat) {
          const perM2 = area > 0 ? (subTotal / area) : 0;
          const perPy = py > 0 ? (subTotal / py) : 0;
          row += `<td class="val-metric">${subTotal > 0 ? subTotal.toLocaleString(undefined, {minimumFractionDigits:2}) : '-'}</td>`;
          row += `<td class="val-metric">${perM2 > 0 ? perM2.toFixed(3) : '-'}</td>`;
          row += `<td class="val-metric">${perPy > 0 ? perPy.toFixed(3) : '-'}</td>`;
          row += `<td class="val-metric">${area > 0 ? area.toLocaleString() : '-'}</td>`;
        } else {
          row += `<td class="val-metric val-zero">-</td><td class="val-metric val-zero">-</td><td class="val-metric val-zero">-</td><td class="val-metric val-zero">-</td>`;
        }
      });
      row += `</tr>`;
      bHtml += row;
    });
  });
  body.innerHTML = bHtml;
}

$("filter-dong-view").onchange = renderMainTable;

// 7. 탭 컨트롤
document.querySelectorAll(".tab").forEach(t => {
  t.onclick = () => {
    document.querySelectorAll(".tab, .tab-panel").forEach(el => el.classList.remove("is-active"));
    t.classList.add("is-active");
    $(`${t.dataset.tab}-tab`).classList.add("is-active");
    
    if (t.dataset.tab === "view") renderMainTable();
  };
});

// 기타 버튼 이벤트
$("btn-apply-mapping").onclick = () => {
  document.querySelector('[data-tab="area"]').click();
};

$("btn-save-area").onclick = () => {
  alert("면적 데이터가 저장되었습니다. 비교표를 확인합니다.");
  document.querySelector('[data-tab="view"]').click();
};

// 8. 엑셀 다운로드 (HTML Table To Excel)
$("btn-excel").onclick = () => {
    const dong = $("filter-dong-view").value;
    if (!dong) return;
    const table = $("main-table");
    const wb = XLSX.utils.table_to_book(table);
    XLSX.writeFile(wb, `QS_통합비교표_${dong}.xlsx`);
};
