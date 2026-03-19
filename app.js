"use strict";

const CATEGORIES = ["콘크리트", "거푸집", "철근", "잡/기타"];

const state = {
  rawItems: [], dongs: [], floors: [], data: {}, mappings: [], areas: {}, ready: false
};

const $ = (id) => document.getElementById(id);

function floorSorter(a, b) {
  const getRank = (name) => {
    const s = String(name).toUpperCase().trim();
    if (s.startsWith('B')) return 1000 - (parseInt(s.replace('B', '')) || 0);
    if (s === 'FT') return 2000;
    if (s.endsWith('F') || /^\d+$/.test(s)) return 3000 + (parseInt(s.replace('F', '')) || 0);
    if (s.startsWith('PH')) return 4000 + (parseInt(s.replace('PH', '')) || 0);
    return 5000;
  };
  return getRank(a) - getRank(b);
}

function predictCategory(name) {
  const s = String(name).toUpperCase().replace(/\s+/g, "");
  if (/(H|D|HD|SD)\d+/.test(s) || s.includes("철근")) return "철근";
  if (s.includes("MPA") || /\d+-\d+-\d+/.test(s) || (/^\d+$/.test(s) && parseInt(s) >= 150)) return "콘크리트";
  if (["폼","FORM","회","알폼","갱폼","합벽"].some(k => s.includes(k)) || /[가-힣]/.test(s)) return "거푸집";
  return "잡/기타";
}

$("btn-parse").onclick = async () => {
  const files = Array.from($('file-main').files);
  if (files.length === 0) return alert("파일을 먼저 선택해주세요.");
  
  state.rawItems = []; state.dongs = []; state.floors = []; state.data = {}; state.areas = {};

  for (const file of files) {
    const rows = XLSX.utils.sheet_to_json(XLSX.read(await file.arrayBuffer(), {type:'array'}).Sheets[XLSX.read(await file.arrayBuffer(), {type:'array'}).SheetNames[0]], {header:1, defval:""});
    parseRows(rows);
  }
  buildMapping(); renderMapping(); switchTab('mapping');
};

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

function buildMapping() {
  state.mappings = state.rawItems.map((item, idx) => ({ id: idx, original: item, canonical: item, category: predictCategory(item) }));
}

function renderMapping() {
  $("mapping-list").innerHTML = state.mappings.map(m => {
    const catClass = m.category === '잡/기타' ? 'etc' : m.category;
    return `
    <div class="item-row cat-${catClass}">
      <div class="col-num">${m.id + 1}</div>
      <div class="col-orig">${m.original}</div>
      <div class="col-edit"><input class="input" value="${m.canonical}" oninput="updateMapping(${m.id},'canonical',this.value)"/></div>
      <div class="col-cat"><select class="input" onchange="updateMapping(${m.id},'category',this.value)">${CATEGORIES.map(c=>`<option value="${c}" ${m.category===c?'selected':''}>${c}</option>`).join("")}</select></div>
    </div>`;
  }).join("");
}
window.updateMapping = (id, f, v) => { state.mappings[id][f] = v; if (f === 'category') renderMapping(); };

$("btn-apply").onclick = () => {
  renderAreaUI();
  switchTab('area');
};

function renderAreaUI() {
  const dongs = state.dongs.sort();
  const floors = state.floors.sort(floorSorter);
  
  let head = `<tr><th>층 명칭</th>${dongs.map(d=>`<th>${d}</th>`).join("")}</tr>`;
  $("area-head").innerHTML = head;

  let body = "";
  floors.forEach((f, rIdx) => {
    body += `<tr><td style="font-weight:bold; background:#f4f7fd;">${f}</td>`;
    dongs.forEach((d, cIdx) => {
      const val = state.areas[d]?.[f] || "";
      body += `<td><input type="number" class="area-input" data-r="${rIdx}" data-c="${cIdx}" value="${val}" oninput="updateArea('${d}','${f}',this.value)" onkeydown="handleAreaNav(event, ${rIdx}, ${cIdx}, ${floors.length}, ${dongs.length})" placeholder="0" /></td>`;
    });
    body += `</tr>`;
  });
  $("area-body").innerHTML = body;
}

window.updateArea = (dong, floor, val) => {
  if (!state.areas[dong]) state.areas[dong] = {};
  state.areas[dong][floor] = parseFloat(val) || 0;
};

window.handleAreaNav = (e, r, c, maxR, maxC) => {
  let nr = r, nc = c;
  if (e.key === 'ArrowUp') nr = Math.max(0, r - 1);
  else if (e.key === 'ArrowDown' || e.key === 'Enter') { nr = Math.min(maxR - 1, r + 1); e.preventDefault(); }
  else if (e.key === 'ArrowLeft') nc = Math.max(0, c - 1);
  else if (e.key === 'ArrowRight') nc = Math.min(maxC - 1, c + 1);
  else return;
  const input = document.querySelector(`.area-input[data-r="${nr}"][data-c="${nc}"]`);
  if (input) { input.focus(); input.select(); }
};

$("btn-download-area").onclick = () => {
  const dongs = state.dongs.sort();
  const floors = state.floors.sort(floorSorter);
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
  const data = XLSX.utils.sheet_to_json(XLSX.read(buffer, { type: 'array' }).Sheets[XLSX.read(buffer, { type: 'array' }).SheetNames[0]], { header: 1, defval: "" });
  
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
  $("filter-dong").innerHTML = state.dongs.sort().map(d => `<option value="${d}">${d}</option>`).join("");
  renderView(); switchTab('view');
};

$("filter-dong").onchange = renderView;

function renderView() {
  if (!state.ready) return;
  const dong = $("filter-dong").value;
  const floors = state.floors.sort(floorSorter);
  const dongData = state.data[dong] || {};
  const grouped = {};

  state.mappings.forEach(m => {
    const qByF = dongData[m.original] || {}; if (Object.keys(qByF).length === 0) return;
    if (!grouped[m.canonical]) grouped[m.canonical] = { category: m.category, floors: {} };
    floors.forEach(f => grouped[m.canonical].floors[f] = (grouped[m.canonical].floors[f] || 0) + (qByF[f] || 0));
  });

  let headHtml = `<tr><th rowspan="2">동</th><th rowspan="2">아이템</th><th rowspan="2">구분</th><th rowspan="2">단위</th><th colspan="${floors.length}">현재 프로젝트 수량</th><th rowspan="2">합계</th></tr><tr>`;
  floors.forEach(f => headHtml += `<th>${f}</th>`); headHtml += "</tr>";
  $("table-head").innerHTML = headHtml;

  let bodyHtml = "";
  ["콘크리트", "철근", "거푸집", "잡/기타"].forEach(cat => {
    const items = Object.keys(grouped).filter(n => grouped[n].category === cat).sort();
    if (items.length === 0) return;
    
    let catSum = 0;
    const catClass = cat === '잡/기타' ? 'etc' : cat;

    items.forEach(name => {
      const item = grouped[name];
      const total = floors.reduce((s,f)=>s+item.floors[f],0);
      catSum += total;
      bodyHtml += `<tr class="row-cat-${catClass}"><td>${dong}</td><td>${cat==='콘크리트'?'레미콘':cat}</td><td>${name}</td><td>${cat==='철근'?'TON':(cat==='콘크리트'?'M3':'M2')}</td>${floors.map(f=>`<td>${item.floors[f].toLocaleString(undefined,{maximumFractionDigits:3})}</td>`).join("")}<td class="col-total">${total.toLocaleString(undefined,{maximumFractionDigits:3})}</td></tr>`;
    });

    bodyHtml += `<tr class="row-subtotal"><td colspan="3" style="text-align:right">합계</td><td>${cat==='철근'?'TON':(cat==='콘크리트'?'M3':'M2')}</td>${floors.map(f => {
      const s = items.reduce((sum, n) => sum + grouped[n].floors[f], 0);
      return `<td>${s.toLocaleString(undefined,{maximumFractionDigits:3})}</td>`;
    }).join("")}<td class="col-total">${catSum.toLocaleString(undefined,{maximumFractionDigits:3})}</td></tr>`;

    const renderRatioRow = (title, unit, numFn, divFn) => {
      let html = `<tr class="row-ratio"><td colspan="3" style="text-align:right">${title}</td><td>${unit}</td>`;
      let totalNum = 0, totalDiv = 0;
      floors.forEach(f => {
        const nVal = numFn(f); const dVal = divFn(f);
        totalNum += nVal; totalDiv += dVal;
        html += `<td>${dVal > 0 ? (nVal/dVal).toFixed(4) : '-'}</td>`;
      });
      html += `<td class="col-total">${totalDiv > 0 ? (totalNum/totalDiv).toFixed(4) : '-'}</td></tr>`;
      return html;
    };

    if(cat === '철근') {
      const numFn = (f) => Object.keys(grouped).filter(n=>grouped[n].category==='철근').reduce((s,n)=>s+grouped[n].floors[f],0);
      bodyHtml += renderRatioRow("지표 (톤당 루베)", "Ton/m³", numFn, (f) => Object.keys(grouped).filter(n=>grouped[n].category==='콘크리트').reduce((s,n)=>s+grouped[n].floors[f],0));
      bodyHtml += renderRatioRow("지표 (톤당 면적)", "Ton/m²", numFn, (f) => state.areas[dong]?.[f] || 0);
      bodyHtml += renderRatioRow("지표 (톤당 평수)", "Ton/Py", numFn, (f) => (state.areas[dong]?.[f] || 0) * 0.3025);
    }
    if(cat === '거푸집') {
      const numFn = (f) => Object.keys(grouped).filter(n=>grouped[n].category==='거푸집').reduce((s,n)=>s+grouped[n].floors[f],0);
      bodyHtml += renderRatioRow("지표 (거푸집/면적)", "m²/m²", numFn, (f) => state.areas[dong]?.[f] || 0);
      bodyHtml += renderRatioRow("지표 (거푸집/평수)", "m²/Py", numFn, (f) => (state.areas[dong]?.[f] || 0) * 0.3025);
    }
  });
  $("table-body").innerHTML = bodyHtml;
}

/* ★ 엑셀 내보내기 (argb 오타 수정 완료) ★ */
$("btn-excel").onclick = async () => {
  if (!state.ready) return alert("먼저 분석을 완료해주세요.");
  if (typeof ExcelJS === 'undefined') return alert("ExcelJS 라이브러리를 불러오지 못했습니다.");

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('비교양식', { views: [{ state: 'frozen', ySplit: 4, xSplit: 4 }] });

  const floors = state.floors.sort(floorSorter);
  const endCol = 4 + floors.length + 1; // "합계" 컬럼
  const maxCol = endCol + 1;            // "비고" 컬럼

  const cols = [{ width: 10 }, { width: 15 }, { width: 18 }, { width: 10 }];
  floors.forEach(() => cols.push({ width: 9 }));
  cols.push({ width: 13 }); 
  cols.push({ width: 12 });
  ws.columns = cols;

  const r1 = ws.addRow(["QS 분석용 프로젝트 통합 템플릿"]); r1.height = 25;
  ws.mergeCells(1, 1, 2, maxCol); // 1~2행 병합
  const titleCell = ws.getCell(1, 1);
  titleCell.font = { size: 16, bold: true, name: '맑은 고딕' };
  titleCell.alignment = { vertical: 'middle', horizontal: 'center' };

  const r3Data = ["동", "아이템", "구분", "단위", "현재 프로젝트 수량"];
  for(let i=0; i<floors.length-1; i++) r3Data.push(""); 
  r3Data.push("합계", "비고");
  
  const r4Data = ["", "", "", ""];
  floors.forEach(f => r4Data.push(f));
  r4Data.push("", "");

  const r3 = ws.addRow(r3Data); r3.height = 22;
  const r4 = ws.addRow(r4Data); r4.height = 22;

  ws.mergeCells(3, 1, 4, 1); ws.mergeCells(3, 2, 4, 2); ws.mergeCells(3, 3, 4, 3); ws.mergeCells(3, 4, 4, 4);
  ws.mergeCells(3, 5, 3, endCol - 1); 
  ws.mergeCells(3, endCol, 4, endCol); 
  ws.mergeCells(3, maxCol, 4, maxCol); 

  const borderAll = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
  for(let r=3; r<=4; r++) {
    for(let c=1; c<=maxCol; c++) {
      const cell = ws.getCell(r, c);
      // [수정] argb 사용, (R:31 G:78 B:120 -> 1F4E78)
      cell.font = { bold: true, size: 10, name: '맑은 고딕', color: { argb: 'FFFFFFFF' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } }; 
      cell.border = borderAll;
    }
  }

  const dataBorder = { top:{style:'thin', color:{argb:'FFBFBFBF'}}, left:{style:'thin', color:{argb:'FFBFBFBF'}}, bottom:{style:'thin', color:{argb:'FFBFBFBF'}}, right:{style:'thin', color:{argb:'FFBFBFBF'}} };

  state.dongs.sort().forEach(dong => {
    const dongData = state.data[dong] || {};
    const grouped = {};
    state.mappings.forEach(m => {
      const qByF = dongData[m.original] || {}; if (Object.keys(qByF).length === 0) return;
      if (!grouped[m.canonical]) grouped[m.canonical] = { category: m.category, floors: {} };
      floors.forEach(f => grouped[m.canonical].floors[f] = (grouped[m.canonical].floors[f] || 0) + (qByF[f] || 0));
    });

    const startRow = ws.rowCount + 1;

    ["콘크리트", "철근", "거푸집"].forEach(cat => {
      const items = Object.keys(grouped).filter(n => grouped[n].category === cat).sort();
      if (items.length === 0) return;

      let rowFill = 'FFFFFFFF';
      if (cat === '콘크리트') rowFill = 'FFEEF4FF';
      else if (cat === '철근') rowFill = 'FFF0FCF4';
      else if (cat === '거푸집') rowFill = 'FFFFF9EC';

      const catSum = {}; floors.forEach(f => catSum[f] = 0);
      let totalSum = 0;

      items.forEach(name => {
        const item = grouped[name];
        const rowData = [dong, cat==='콘크리트'?'레미콘':cat, name, cat==='철근'?'TON':(cat==='콘크리트'?'M3':'M2')];
        let rowTotal = 0;
        floors.forEach(f => { rowData.push(item.floors[f] || 0); catSum[f] += (item.floors[f] || 0); rowTotal += (item.floors[f] || 0); });
        rowData.push(rowTotal); 
        rowData.push(""); 

        const row = ws.addRow(rowData);
        row.height = 18; row.outlineLevel = 1; totalSum += rowTotal;
        
        for(let c=1; c<=maxCol; c++) {
          const cell = row.getCell(c);
          cell.border = dataBorder; cell.font = { name: '맑은 고딕', size: 10 };
          // [수정] argb 사용
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowFill } };
          if (c <= 4) cell.alignment = { vertical: 'middle', horizontal: 'center' };
          else { cell.alignment = { vertical: 'middle', horizontal: 'right' }; cell.numFmt = '#,##0.000'; }
        }
      });

      const sumRowData = [dong, cat==='콘크리트'?'레미콘':cat, "합계", cat==='철근'?'TON':(cat==='콘크리트'?'M3':'M2')];
      floors.forEach(f => sumRowData.push(catSum[f])); 
      sumRowData.push(totalSum); 
      sumRowData.push(""); 
      
      const sumRow = ws.addRow(sumRowData); 
      sumRow.height = 18; sumRow.outlineLevel = 0; 
      for(let c=1; c<=maxCol; c++) {
        const cell = sumRow.getCell(c);
        cell.font = { name: '맑은 고딕', size: 10, bold: true };
        // [수정] argb 사용
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } }; 
        cell.border = dataBorder;
        if (c <= 4) cell.alignment = { vertical: 'middle', horizontal: 'center' };
        else { cell.alignment = { vertical: 'middle', horizontal: 'right' }; cell.numFmt = '#,##0.000'; }
      }

      const renderExcelRatio = (title, unit, numFn, divFn) => {
        const ratioRowData = [dong, "지표", title, unit];
        let totalNum = 0, totalDiv = 0;
        
        floors.forEach(f => {
          const nVal = numFn(f); const dVal = divFn(f);
          totalNum += nVal; totalDiv += dVal;
          ratioRowData.push(dVal > 0 ? (nVal / dVal) : 0);
        });
        
        ratioRowData.push(totalDiv > 0 ? (totalNum / totalDiv) : 0); 
        ratioRowData.push(""); 
        
        const ratioRow = ws.addRow(ratioRowData);
        ratioRow.height = 18;
        
        for(let c=1; c<=maxCol; c++) {
          const cell = ratioRow.getCell(c);
          // [수정] argb 사용, (R:150 G:54 B:52 -> 963634)
          cell.font = { name: '맑은 고딕', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF963634' } };
          cell.border = dataBorder;
          if (c <= 4) cell.alignment = { vertical: 'middle', horizontal: 'center' };
          else { cell.alignment = { vertical: 'middle', horizontal: 'right' }; cell.numFmt = '#,##0.0000'; }
        }
      };

      if (cat === '철근') {
        const numFn = (f) => Object.keys(grouped).filter(n=>grouped[n].category==='철근').reduce((s,n)=>s+grouped[n].floors[f],0);
        renderExcelRatio("레미콘/철근", "Ton/m³", numFn, (f) => Object.keys(grouped).filter(n=>grouped[n].category==='콘크리트').reduce((s,n)=>s+grouped[n].floors[f],0));
        renderExcelRatio("면적/철근", "Ton/m²", numFn, (f) => state.areas[dong]?.[f] || 0);
        renderExcelRatio("평수/철근", "Ton/Py", numFn, (f) => (state.areas[dong]?.[f] || 0) * 0.3025);
      }
      
      if (cat === '거푸집') {
        const numFn = (f) => Object.keys(grouped).filter(n=>grouped[n].category==='거푸집').reduce((s,n)=>s+grouped[n].floors[f],0);
        renderExcelRatio("거푸집/면적", "m²/m²", numFn, (f) => state.areas[dong]?.[f] || 0);
        renderExcelRatio("거푸집/평수", "m²/Py", numFn, (f) => (state.areas[dong]?.[f] || 0) * 0.3025);
      }
    });

    const endRow = ws.rowCount;
    if (startRow < endRow) {
      ws.mergeCells(startRow, 1, endRow, 1);
      ws.getCell(startRow, 1).alignment = { vertical: 'middle', horizontal: 'center' };
    }
  });

  const buffer = await wb.xlsx.writeBuffer();
  saveAs(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), "QS_통합템플릿_리포트.xlsx");
};

function switchTab(id) {
  document.querySelectorAll(".tab, .tab-panel").forEach(el => el.classList.remove("is-active"));
  document.querySelector(`[data-tab="${id}"]`).classList.add("is-active");
  $("tab-" + id).classList.add("is-active");
}
document.querySelectorAll(".tab").forEach(t => t.onclick = () => switchTab(t.dataset.tab));
