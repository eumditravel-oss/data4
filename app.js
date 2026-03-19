"use strict";

const CATEGORIES = ["레미콘", "거푸집", "철근D", "철근H"];
const $ = (id) => document.getElementById(id);

let state = {
  data: {},    // { 동: { 층: { 아이템: 값 } } }
  dongs: [],
  floors: [],
  items: [],   // { name, category }
  ready: false
};

// 1. 아이템 분류 로직
function getCategory(itemName) {
  const n = itemName.toUpperCase().replace(/\s/g, "");
  if (n.includes("MPA") || /\d+-\d+-\d+/.test(n)) return "레미콘";
  if (["폼", "FORM", "거푸집", "갱폼", "알폼"].some(k => n.includes(k))) return "거푸집";
  if (/(HD|SD|H)\d+/.test(n)) return "철근H";
  if (/D\d+/.test(n)) return "철근D";
  return "기타";
}

// 2. 층 정렬
function sortFloors(arr) {
  const rank = (f) => {
    if (f.startsWith('B')) return 1000 - parseInt(f.substring(1));
    if (f === 'FT' || f === '기초') return 2000;
    if (f.endsWith('F')) return 3000 + parseInt(f);
    return 4000;
  };
  return arr.sort((a, b) => rank(a) - rank(b));
}

// 3. 파일 파싱
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

      // 아이템 등록
      if (!state.items.find(i => i.name === item)) {
        state.items.push({ name: item, category: getCategory(item) });
      }

      if (!state.data[dong]) state.data[dong] = {};
      if (!state.dongs.includes(dong)) state.dongs.push(dong);

      Object.keys(row).forEach(key => {
        if (["동", "아이템", "구분", "단위", "합계"].includes(key)) return;
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
  $("file-list").innerHTML = "✅ 분석 완료! 2번 탭으로 이동하세요.";
  alert("데이터 로드 완료");
};

// 4. 테이블 렌더링 (4행 1세트 구조)
function renderTable() {
  const dong = $("filter-dong").value;
  if (!dong) return;

  const head = $("table-head");
  const body = $("table-body");

  // 카테고리별 아이템 필터링
  const groups = {
    "레미콘": state.items.filter(i => i.category === "레미콘"),
    "거푸집": state.items.filter(i => i.category === "거푸집"),
    "철근D": state.items.filter(i => i.category === "철근D"),
    "철근H": state.items.filter(i => i.category === "철근H")
  };

  // 1) 헤더 생성
  let h1 = `<tr><th rowspan="2">층</th><th rowspan="2">구분</th>`;
  let h2 = `<tr>`;
  
  CATEGORIES.forEach(cat => {
    const cols = groups[cat];
    h1 += `<th colspan="${cols.length || 1}">${cat}</th>`;
    if (cols.length === 0) h2 += `<th>-</th>`;
    else cols.forEach(i => h2 += `<th>${i.name}</th>`);
  });
  head.innerHTML = h1 + `</tr>` + h2 + `</tr>`;

  // 2) 바디 생성 (층별 루프)
  let bHtml = "";
  state.floors.forEach(f => {
    CATEGORIES.forEach((rowCat, idx) => {
      let row = `<tr class="${idx === 3 ? 'row-group-end' : ''}">`;
      if (idx === 0) row += `<td rowspan="4" style="text-align:center; font-weight:bold;">${f}</td>`;
      row += `<td style="text-align:center; background:#f9f9f9;">${rowCat}</td>`;

      // 데이터 셀 생성
      CATEGORIES.forEach(colCat => {
        const targetItems = groups[colCat];
        if (targetItems.length === 0) row += `<td>-</td>`;
        else {
          targetItems.forEach(item => {
            // 현재 행의 카테고리와 컬럼의 카테고리가 일치할 때만 값 출력
            const val = (rowCat === colCat) ? (state.data[dong]?.[f]?.[item.name] || 0) : 0;
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

// 5. 탭 및 필터 이벤트
document.querySelectorAll(".tab").forEach(t => {
  t.onclick = () => {
    document.querySelectorAll(".tab, .tab-panel").forEach(el => el.classList.remove("is-active"));
    t.classList.add("is-active");
    $(`${t.dataset.tab}-tab`).classList.add("is-active");
    
    if (t.dataset.tab === "view" && state.ready) {
      $("filter-dong").innerHTML = state.dongs.map(d => `<option value="${d}">${d}</option>`).join('');
      renderTable();
    }
  };
});

$("filter-dong").onchange = renderTable;

// 6. 엑셀 다운로드 (ExcelJS)
$("btn-excel").onclick = async () => {
    const dong = $("filter-dong").value;
    if (!dong) return;
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet(dong);
    
    // 헤더와 데이터를 수동으로 구성 (웹 테이블 구조와 동일하게)
    // ... (상세 엑셀 스타일링 로직 생략 가능 - 필요시 추가 구현)
    alert("엑셀 생성이 시작됩니다.");
    const table = $("main-table");
    const wb = XLSX.utils.table_to_book(table);
    XLSX.writeFile(wb, `QS_비교표_${dong}.xlsx`);
};
