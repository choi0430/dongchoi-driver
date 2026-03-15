
// ═══════════════════════════════════════════════════════════════
// DONG CHOI PTY LTD — Google Apps Script (완전판)
// 모든 앱 ↔ 구글 시트 동기화
// ═══════════════════════════════════════════════════════════════

const SS = SpreadsheetApp.getActiveSpreadsheet();

// ── 시트 이름 상수
const SHEETS = {
  DAILY:    'Daily_Report',
  PRE:      'Pre_Departure',
  EOS:      'End_of_Shift',
  DRIVERS:  'M_Drivers',
  VEHICLES: 'M_Vehicles',
  AGENCIES: 'M_Agencies',
  COURSES:  'M_Courses',
  PRICE:    'M_PriceDriver',
  HOTELS:   'M_Hotels',
  SUB:      'M_Sub_Rates',
  LEDGER:   'Ledger',
  WAGES:    'Wages',
  NOTICES:  'M_Notices',
};

// ── 공통: 시트 → JSON 배열
function sheetToJSON(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      const v = row[i];
      // Date 오브젝트는 dd/mm/yyyy 문자열로 변환 (UTC 직렬화 시 타임존 오류 방지)
      obj[h] = (v instanceof Date) ? fmtDate(v) : (v !== undefined && v !== null ? v : '');
    });
    return obj;
  }).filter(row => headers.some(h => row[h] !== ''));
}

// ── 공통: 날짜 포맷 (dd/mm/yyyy)
function fmtDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const d = val;
    return String(d.getDate()).padStart(2,'0') + '/' +
           String(d.getMonth()+1).padStart(2,'0') + '/' + d.getFullYear();
  }
  return String(val);
}

// ── GET 라우터
function doGet(e) {
  const action = e.parameter.action || '';
  const sheet  = e.parameter.sheet  || '';
  try {
    let result;
    switch(action) {
      case 'ping':         result = {ok:true, ts: new Date().toISOString()}; break;
      case 'get_reports':  result = getReports(sheet); break;
      case 'get_master':   result = getMaster(sheet); break;
      case 'get_all_masters': result = getAllMasters(); break;
      case 'get_active_regos': result = getActiveRegos(); break;
      case 'get_wages':    result = getWages(); break;
      case 'get_ledger':   result = getLedger(); break;
      case 'get_sub_rates': result = getSubRates(); break;
      case 'get_price_sub': result = getPriceSub(); break;
      case 'get_notices':  result = getNotices(); break;
      default: result = {ok:false, error:'Unknown action: '+action};
    }
    return json(result);
  } catch(err) {
    return json({ok:false, error: err.message});
  }
}

// ── POST 라우터
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action || '';
    let result;
    switch(action) {
      // ── 리포트 저장
      case 'save_report':       result = saveReport(payload);       break;
      case 'save_predeparture': result = savePredeparture(payload); break;
      case 'save_endofshift':   result = saveEndofshift(payload);   break;
      // ── 마스터 관리
      case 'add_master':        result = addMaster(payload);    break;
      case 'update_master':     result = updateMaster(payload); break;
      case 'delete_master':     result = deleteMaster(payload); break;
      case 'replace_master':    result = replaceMaster(payload);break;
      case 'init_masters':      result = initMasters(payload);  break;
      // ── 드라이버 PIN
      case 'update_driver_pin': result = updateDriverPin(payload); break;
      // ── 급여
      case 'add_wage':          result = addWage(payload);     break;
      case 'update_wage':       result = updateWage(payload);  break;
      case 'delete_wage':       result = deleteWage(payload);  break;
      case 'replace_wages':     result = replaceWages(payload);break;
      // ── 정산
      case 'add_ledger':        result = addLedger(payload);    break;
      case 'update_ledger':     result = updateLedger(payload); break;
      case 'delete_ledger':     result = deleteLedger(payload); break;
      case 'replace_ledger':    result = replaceLedger(payload);break;
      // ── 기타
      case 'replace_sub_rates': result = replaceSubRates(payload); break;
      case 'replace_price_sub': result = replacePriceSub(payload); break;
      // ── 공지사항
      case 'save_notices':      result = saveNotices(payload);  break;
      default: result = {ok:false, error:'Unknown action: '+action};
    }
    return json(result);
  } catch(err) {
    return json({ok:false, error: err.message});
  }
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════
// GET 구현
// ═══════════════════════════════════════

function getReports(sheetName) {
  const s = SS.getSheetByName(sheetName);
  if (!s) return {ok:true, rows:[]};
  return {ok:true, rows: sheetToJSON(s)};
}

function getMaster(sheetName) {
  const s = SS.getSheetByName(sheetName);
  if (!s) return {ok:true, rows:[]};
  return {ok:true, rows: sheetToJSON(s)};
}

function getAllMasters() {
  const result = {ok:true};
  // M_Drivers
  const drvS = SS.getSheetByName(SHEETS.DRIVERS);
  result.drivers = drvS ? sheetToJSON(drvS) : [];
  // M_Vehicles
  const vehS = SS.getSheetByName(SHEETS.VEHICLES);
  result.vehicles = vehS ? sheetToJSON(vehS) : [];
  // M_Agencies
  const agtS = SS.getSheetByName(SHEETS.AGENCIES);
  result.agencies = agtS ? sheetToJSON(agtS) : [];
  // M_Courses
  const crsS = SS.getSheetByName(SHEETS.COURSES);
  result.courses = crsS ? sheetToJSON(crsS) : [];
  // M_PriceDriver
  const prcS = SS.getSheetByName(SHEETS.PRICE);
  result.priceDriver = prcS ? sheetToJSON(prcS) : [];
  // Ledger
  const ledS = SS.getSheetByName(SHEETS.LEDGER);
  result.ledger = ledS ? sheetToJSON(ledS) : [];
  // Wages
  const wageS = SS.getSheetByName(SHEETS.WAGES);
  result.wages = wageS ? sheetToJSON(wageS) : [];
  // M_Hotels
  const htlS = SS.getSheetByName(SHEETS.HOTELS);
  result.hotels = htlS ? sheetToJSON(htlS) : [];
  // Notices
  const ntcS = SS.getSheetByName(SHEETS.NOTICES);
  result.notices = ntcS ? sheetToJSON(ntcS) : [];
  // data 래퍼 (기존 앱 호환)
  result.data = {
    M_Drivers:    result.drivers,
    M_Vehicles:   result.vehicles,
    M_Agencies:   result.agencies,
    M_Courses:    result.courses,
    M_PriceDriver:result.priceDriver,
    M_Hotels:     result.hotels,
  };
  return result;
}

function getActiveRegos() {
  const s = SS.getSheetByName(SHEETS.VEHICLES);
  if (!s) return {ok:true, regos:[]};
  const rows = sheetToJSON(s);
  const regos = rows
    .filter(r => String(r.Active||'').toUpperCase() === 'Y')
    .map(r => ({
      rego:         r.Rego,
      seats:        r.Seats,
      make:         r.Make,
      model:        r.Model,
      transmission: r.Transmission||'Auto'
    }));
  return {ok:true, regos};
}

function getWages() {
  const s = SS.getSheetByName(SHEETS.WAGES);
  if (!s) return {ok:true, rows:[]};
  return {ok:true, rows: sheetToJSON(s)};
}

function getLedger() {
  const s = SS.getSheetByName(SHEETS.LEDGER);
  if (!s) return {ok:true, rows:[]};
  return {ok:true, rows: sheetToJSON(s)};
}

function getSubRates() {
  const s = SS.getSheetByName(SHEETS.SUB);
  if (!s) return {ok:true, rows:[]};
  return {ok:true, rows: sheetToJSON(s)};
}

function getPriceSub() {
  const s = SS.getSheetByName(SHEETS.PRICE);
  if (!s) return {ok:true, rows:[]};
  return {ok:true, rows: sheetToJSON(s)};
}

function getNotices() {
  const s = SS.getSheetByName(SHEETS.NOTICES);
  if (!s) return {ok:true, rows:[]};
  return {ok:true, rows: sheetToJSON(s)};
}

// ═══════════════════════════════════════
// 리포트 저장
// ═══════════════════════════════════════

function saveReport(p) {
  let s = SS.getSheetByName(SHEETS.DAILY);
  if (!s) {
    s = SS.insertSheet(SHEETS.DAILY);
    s.appendRow(['Driver','Date','Rego','Seats','Agency','Attraction',
      'Pickup','Dropoff','KM_Start','KM_End','Time_Start','Time_End',
      'Guide','Tour_Code','SVC_Label','SVC_Charge','Hotel_Surcharge',
      'Dist_Surcharge','OT','Early','Night_Type','Night_DR','Night_Owner',
      'Trailer','Wash','Meal','Toll','Fuel','Total_TA','DR_Cost',
      'Remarks','Submitted']);
  }
  const d = p.data || p;
  s.appendRow([
    d.Driver||'', d.Date||'', d.Rego||'', d.Seats||'',
    d.Agency||'', d.Attraction||'', d.Pickup||'', d.Dropoff||'',
    d.KM_Start||'', d.KM_End||'', d.Time_Start||'', d.Time_End||'',
    d.Guide||'', d.Tour_Code||'', d.SVC_Label||'', d.SVC_Charge||0,
    d.Hotel_Surcharge||0, d.Dist_Surcharge||0, d.OT||0, d.Early||0,
    d.Night_Type||'', d.Night_DR||0, d.Night_Owner||0,
    d.Trailer||0, d.Wash||0, d.Meal||0, d.Toll||0, d.Fuel||0,
    d.Total_TA||0, d.DR_Cost||0,
    d.Remarks||'', d.Submitted||''
  ]);
  return {ok:true};
}

function savePredeparture(p) {
  let s = SS.getSheetByName(SHEETS.PRE);
  if (!s) {
    s = SS.insertSheet(SHEETS.PRE);
    s.appendRow(['Driver','Date','Rego','Seats','Start_KM','Fuel',
      'Start_Time','Check_Results','Remarks','Submitted']);
  }
  const d = p.data || p;
  s.appendRow([
    d.Driver||'', d.Date||'', d.Rego||'', d.Seats||'',
    d.Start_KM||'', d.Fuel||'', d.Start_Time||'',
    d.Check_Results||'', d.Remarks||'', d.Submitted||''
  ]);
  return {ok:true};
}

function saveEndofshift(p) {
  let s = SS.getSheetByName(SHEETS.EOS);
  if (!s) {
    s = SS.insertSheet(SHEETS.EOS);
    s.appendRow(['Driver','Date','Rego','End_KM','End_Time',
      'Check_Results','Damage_Report','Remarks','Submitted']);
  }
  const d = p.data || p;
  s.appendRow([
    d.Driver||'', d.Date||'', d.Rego||'', d.End_KM||'', d.End_Time||'',
    d.Check_Results||'', d.Damage_Report||'', d.Remarks||'', d.Submitted||''
  ]);
  return {ok:true};
}

// ═══════════════════════════════════════
// 마스터 관리 (공통)
// ═══════════════════════════════════════

function addMaster(p) {
  const s = SS.getSheetByName(p.sheet);
  if (!s) return {ok:false, error:'Sheet not found: '+p.sheet};
  const headers = s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
  const row = headers.map(h => p.data[h] !== undefined ? p.data[h] : '');
  s.appendRow(row);
  return {ok:true};
}

function updateMaster(p) {
  const s = SS.getSheetByName(p.sheet);
  if (!s) return {ok:false, error:'Sheet not found'};
  const data = s.getDataRange().getValues();
  const headers = data[0];
  const keyCol = headers.indexOf(p.keyField || headers[0]);
  const keyVal = p.keyValue || (p.data && p.data[headers[keyCol]]);
  for (let i=1; i<data.length; i++) {
    if (String(data[i][keyCol]) === String(keyVal)) {
      headers.forEach((h,j) => {
        if (p.data[h] !== undefined) s.getRange(i+1,j+1).setValue(p.data[h]);
      });
      return {ok:true};
    }
  }
  // 없으면 추가
  s.appendRow(headers.map(h => p.data[h] !== undefined ? p.data[h] : ''));
  return {ok:true};
}

function deleteMaster(p) {
  const s = SS.getSheetByName(p.sheet);
  if (!s) return {ok:false, error:'Sheet not found'};
  const data = s.getDataRange().getValues();
  const headers = data[0];
  const keyCol = headers.indexOf(p.keyField || headers[0]);
  for (let i=data.length-1; i>=1; i--) {
    if (String(data[i][keyCol]) === String(p.keyValue)) {
      s.deleteRow(i+1);
      return {ok:true};
    }
  }
  return {ok:false, error:'Row not found'};
}

function replaceMaster(p) {
  const s = SS.getSheetByName(p.sheet);
  if (!s) return {ok:false, error:'Sheet not found'};
  const headers = s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
  // 데이터 행 모두 삭제 후 재입력
  if (s.getLastRow() > 1) s.deleteRows(2, s.getLastRow()-1);
  (p.rows||[]).forEach(row => s.appendRow(headers.map(h => row[h] !== undefined ? row[h] : '')));
  return {ok:true};
}

function initMasters(p) {
  // 여러 시트 초기화 (배열)
  (p.sheets||[]).forEach(item => {
    let s = SS.getSheetByName(item.name);
    if (!s) s = SS.insertSheet(item.name);
    s.clearContents();
    if (item.headers) s.appendRow(item.headers);
    (item.rows||[]).forEach(row => {
      s.appendRow(Array.isArray(row) ? row : item.headers.map(h=>row[h]||''));
    });
  });
  return {ok:true};
}

function updateDriverPin(p) {
  const s = SS.getSheetByName(SHEETS.DRIVERS);
  if (!s) return {ok:false, error:'M_Drivers sheet not found'};
  const data = s.getDataRange().getValues();
  const headers = data[0];
  const nameKRCol = headers.indexOf('Name_KR');
  const nameENCol = headers.indexOf('Name_EN');
  const pinCol    = headers.indexOf('PIN');
  if (pinCol < 0) return {ok:false, error:'PIN column not found'};
  for (let i=1; i<data.length; i++) {
    const nKR = String(data[i][nameKRCol]||'');
    const nEN = String(data[i][nameENCol]||'');
    if (nKR === p.name || nEN === p.name) {
      s.getRange(i+1, pinCol+1).setValue(p.pin);
      return {ok:true};
    }
  }
  return {ok:false, error:'Driver not found: '+p.name};
}

// ═══════════════════════════════════════
// 급여 (Wages)
// ═══════════════════════════════════════

function ensureWagesSheet() {
  let s = SS.getSheetByName(SHEETS.WAGES);
  if (!s) {
    s = SS.insertSheet(SHEETS.WAGES);
    s.appendRow(['Driver','WeekStart','Date','Amount','Method','Note']);
  }
  return s;
}

function getWagesSheet() { return ensureWagesSheet(); }

function addWage(p) {
  const s = ensureWagesSheet();
  const d = p.data;
  s.appendRow([d.Driver||'', d.WeekStart||'', d.Date||'',
    d.Amount||0, d.Method||'현금', d.Note||'']);
  const row = s.getLastRow();
  return {ok:true, row};
}

function updateWage(p) {
  const s = ensureWagesSheet();
  const ri = p.rowIndex;
  if (!ri || ri < 2) return {ok:false, error:'Invalid rowIndex'};
  const d = p.data;
  s.getRange(ri,1,1,6).setValues([[
    d.Driver||'', d.WeekStart||'', d.Date||'',
    d.Amount||0, d.Method||'현금', d.Note||''
  ]]);
  return {ok:true};
}

function deleteWage(p) {
  const s = ensureWagesSheet();
  const ri = p.rowIndex;
  if (!ri || ri < 2) return {ok:false, error:'Invalid rowIndex'};
  s.deleteRow(ri);
  return {ok:true};
}

function replaceWages(p) {
  const s = ensureWagesSheet();
  if (s.getLastRow() > 1) s.deleteRows(2, s.getLastRow()-1);
  (p.rows||[]).forEach(r => {
    s.appendRow([r.Driver||'', r.WeekStart||'', r.Date||'',
      r.Amount||0, r.Method||'현금', r.Note||'']);
  });
  return {ok:true};
}

// ═══════════════════════════════════════
// 정산 (Ledger)
// ═══════════════════════════════════════

function ensureLedgerSheet() {
  let s = SS.getSheetByName(SHEETS.LEDGER);
  if (!s) {
    s = SS.insertSheet(SHEETS.LEDGER);
    s.appendRow(['ID','Date','Agency','Description','Amount','Type','Note']);
  }
  return s;
}

function addLedger(p) {
  const s = ensureLedgerSheet();
  const d = p.data;
  s.appendRow([d.ID||'', d.Date||'', d.Agency||'', d.Description||'',
    d.Amount||0, d.Type||'', d.Note||'']);
  return {ok:true, row: s.getLastRow()};
}

function updateLedger(p) {
  const s = ensureLedgerSheet();
  const ri = p.rowIndex;
  if (!ri || ri < 2) return {ok:false, error:'Invalid rowIndex'};
  const d = p.data;
  s.getRange(ri,1,1,7).setValues([[
    d.ID||'', d.Date||'', d.Agency||'', d.Description||'',
    d.Amount||0, d.Type||'', d.Note||''
  ]]);
  return {ok:true};
}

function deleteLedger(p) {
  const s = ensureLedgerSheet();
  const ri = p.rowIndex;
  if (!ri || ri < 2) return {ok:false, error:'Invalid rowIndex'};
  s.deleteRow(ri);
  return {ok:true};
}

function replaceLedger(p) {
  const s = ensureLedgerSheet();
  if (s.getLastRow() > 1) s.deleteRows(2, s.getLastRow()-1);
  (p.rows||[]).forEach(r => {
    s.appendRow([r.ID||'', r.Date||'', r.Agency||'', r.Description||'',
      r.Amount||0, r.Type||'', r.Note||'']);
  });
  return {ok:true};
}

// ═══════════════════════════════════════
// 기타
// ═══════════════════════════════════════

function replaceSubRates(p) {
  let s = SS.getSheetByName(SHEETS.SUB);
  if (!s) { s = SS.insertSheet(SHEETS.SUB); }
  s.clearContents();
  if (p.rows && p.rows.length > 0) {
    const headers = Object.keys(p.rows[0]);
    s.appendRow(headers);
    p.rows.forEach(r => s.appendRow(headers.map(h => r[h]||'')));
  }
  return {ok:true};
}

function replacePriceSub(p) {
  let s = SS.getSheetByName(SHEETS.PRICE);
  if (!s) { s = SS.insertSheet(SHEETS.PRICE); }
  s.clearContents();
  if (p.rows && p.rows.length > 0) {
    const headers = Object.keys(p.rows[0]);
    s.appendRow(headers);
    p.rows.forEach(r => s.appendRow(headers.map(h => r[h]||'')));
  }
  return {ok:true};
}

// ═══════════════════════════════════════
// 공지사항 (M_Notices)
// ═══════════════════════════════════════

function ensureNoticesSheet() {
  let s = SS.getSheetByName(SHEETS.NOTICES);
  if (!s) {
    s = SS.insertSheet(SHEETS.NOTICES);
    s.appendRow(['ID','Title','Content','Type','Date','Active']);
  }
  return s;
}

function saveNotices(p) {
  const s = ensureNoticesSheet();
  // 전체 교체
  if (s.getLastRow() > 1) s.deleteRows(2, s.getLastRow()-1);
  (p.rows||[]).forEach(r => {
    s.appendRow([
      r.id||r.ID||String(Date.now()),
      r.title||r.Title||'',
      r.content||r.Content||'',
      r.type||r.Type||'info',
      r.date||r.Date||'',
      r.active===false||r.Active==='false' ? 'false' : 'true'
    ]);
  });
  return {ok:true};
}
