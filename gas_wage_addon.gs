// ═══════════════════════════════════════════════════════════
// DC FLEET — WAGES SYNC (기존 GAS에 추가할 코드)
// ═══════════════════════════════════════════════════════════
// 아래 코드를 기존 GAS doGet/doPost 함수의 switch(action) 안에 추가하세요.
// 그리고 setupWagesSheet() 함수와 헬퍼들을 파일 하단에 붙여넣으세요.
// ═══════════════════════════════════════════════════════════

// ──────────────────────────────────────────
// doGet switch에 추가할 cases:
// ──────────────────────────────────────────
/*
  case 'get_wages':
    return jsonRes(getWages(e.parameter));

  case 'get_wages_driver':
    return jsonRes(getWagesForDriver(e.parameter.driver || ''));
*/

// ──────────────────────────────────────────
// doPost switch에 추가할 cases:
// ──────────────────────────────────────────
/*
  case 'add_wage':
    return jsonRes(addWage(body.data));

  case 'update_wage':
    return jsonRes(updateWage(body.rowIndex, body.data));

  case 'delete_wage':
    return jsonRes(deleteWage(body.rowIndex));

  case 'replace_wages':
    return jsonRes(replaceWages(body.rows));
*/

// ═══════════════════════════════════════════════════════════
// 아래 함수들을 GAS 파일 하단에 붙여넣기:
// ═══════════════════════════════════════════════════════════

function getOrCreateWagesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Wages');
  if (!sh) {
    sh = ss.insertSheet('Wages');
    // 헤더 행 설정
    const headers = ['Driver', 'WeekStart', 'Date', 'Amount', 'Method', 'Note', 'RowID'];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    // 헤더 스타일
    sh.getRange(1, 1, 1, headers.length)
      .setBackground('#1e293b')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sh.setColumnWidth(1, 120); // Driver
    sh.setColumnWidth(2, 110); // WeekStart
    sh.setColumnWidth(3, 100); // Date
    sh.setColumnWidth(4, 80);  // Amount
    sh.setColumnWidth(5, 80);  // Method
    sh.setColumnWidth(6, 200); // Note
    sh.setColumnWidth(7, 80);  // RowID
  }
  return sh;
}

function getWages(params) {
  try {
    const sh = getOrCreateWagesSheet();
    const data = sh.getDataRange().getValues();
    if (data.length <= 1) return { ok: true, rows: [] };
    const headers = data[0].map(h => String(h).trim());
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const row = {};
      headers.forEach((h, j) => { row[h] = data[i][j] !== undefined ? String(data[i][j]) : ''; });
      row._rowIndex = i + 1; // 1-based sheet row
      // Date 정규화
      const ds = row.Date || '';
      if (/^\d{4}-\d{2}-\d{2}$/.test(ds)) {
        row._isoDate = ds;
      } else if (/^\d{2}\/\d{2}\/\d{4}$/.test(ds)) {
        const [dd, mm, yyyy] = ds.split('/');
        row._isoDate = `${yyyy}-${mm}-${dd}`;
      } else {
        row._isoDate = ds;
      }
      if (row.Driver || row.WeekStart) rows.push(row);
    }
    // 드라이버 필터
    if (params && params.driver) {
      const drv = params.driver;
      return { ok: true, rows: rows.filter(r => r.Driver === drv) };
    }
    return { ok: true, rows };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

function getWagesForDriver(driverName) {
  return getWages({ driver: driverName });
}

function addWage(data) {
  try {
    const sh = getOrCreateWagesSheet();
    const rowId = Date.now().toString();
    const lastRow = sh.getLastRow();
    const newRow = [
      data.Driver   || '',
      data.WeekStart|| '',
      data.Date     || '',
      parseFloat(data.Amount) || 0,
      data.Method   || '현금',
      data.Note     || '',
      rowId,
    ];
    sh.appendRow(newRow);
    return { ok: true, row: lastRow + 1, rowId };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

function updateWage(rowIndex, data) {
  try {
    const sh = getOrCreateWagesSheet();
    const ri = parseInt(rowIndex);
    if (!ri || ri < 2) return { ok: false, msg: '유효하지 않은 행 번호' };
    const lastRow = sh.getLastRow();
    if (ri > lastRow) return { ok: false, msg: '행이 존재하지 않음' };
    sh.getRange(ri, 1, 1, 6).setValues([[
      data.Driver   || '',
      data.WeekStart|| '',
      data.Date     || '',
      parseFloat(data.Amount) || 0,
      data.Method   || '현금',
      data.Note     || '',
    ]]);
    return { ok: true };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

function deleteWage(rowIndex) {
  try {
    const sh = getOrCreateWagesSheet();
    const ri = parseInt(rowIndex);
    if (!ri || ri < 2) return { ok: false, msg: '유효하지 않은 행 번호' };
    sh.deleteRow(ri);
    return { ok: true };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

function replaceWages(rows) {
  try {
    const sh = getOrCreateWagesSheet();
    // 헤더 제외 모든 행 삭제
    const lastRow = sh.getLastRow();
    if (lastRow > 1) sh.deleteRows(2, lastRow - 1);
    // 새 행 추가
    if (rows && rows.length > 0) {
      const newData = rows.map(r => [
        r.Driver   || '',
        r.WeekStart|| '',
        r.Date     || '',
        parseFloat(r.Amount) || 0,
        r.Method   || '현금',
        r.Note     || '',
        r.RowID    || Date.now().toString(),
      ]);
      sh.getRange(2, 1, newData.length, 7).setValues(newData);
    }
    return { ok: true, count: rows ? rows.length : 0 };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}
