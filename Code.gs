// ═══════════════════════════════════════════════════
// DC Fleet — Google Sheets Backend (Apps Script)
// Spreadsheet: Dong Choi Pty Ltd - Driver Reports
// ═══════════════════════════════════════════════════

const SHEET_ID = '1kUU-_-IFJkKd97O-Im-A6xojsafYG-0njVyRKmSLKeE';

// ── 리포트 시트 헤더 ──
const REPORT_HEADERS = {
  'Daily_Report':   ['Submitted','Driver','Date','Rego','Seats','Agency','Attraction','Pickup','Dropoff',
                     'KM_Start','KM_End','Time_Start','Time_End','Guide','Tour_Code',
                     'SVC_Label','SVC_Charge','Hotel_Surcharge','Dist_Surcharge',
                     'OT','Trailer','Total_TA','DR_Cost','Toll','Toll_Personal',
                     'Fuel','Fuel_Personal','Early','Night_Type','Night_DR','Night_Owner',
                     'Wash','Meal','Tip','Etc','Remarks'],
  'Pre_Departure':  ['Submitted','Driver','Date','Rego','Seats','Start_KM','Fuel','Start_Time',
                     'Check_Results','Remarks'],
  'End_of_Shift':   ['Submitted','Driver','Date','Rego','End_KM','End_Time','Fuel_End','Remarks'],
  'MOT_Report':     ['Submitted','Driver','Date','Time','Rego','Location','Officer','Type',
                     'Result','NoticeNum','Fine','Notes','FailedItems','Checks']
};

// ── 마스터 시트 헤더 ──
const MASTER_HEADERS = {
  'M_Vehicles': ['Rego','Make','Model','Manufacture_Date','Capacity','Owner','Rego_Date','HVIS_Date',
                 'Current_KM','Last_Service_KM','Service_Interval','VIN','Engine_Number',
                 'Accreditation','Current_Status'],
  'M_Drivers':  ['Name_EN','Name_KR','DriverID','Mobile_1','NEXT_OF_KIN','License_Class',
                 'License_No','License_Expiry','Authority_No','Authority_Expiry',
                 'Address','Suburb','Bank_Name','BSB','Account_Number','PIN'],
  'M_Clients':  ['Name','ClientID','Mobile','Address','Bank_Name','BSB','Account_Number'],
  'M_Guides':   ['GuideID','Guide_Name','Mobile','Agency','Email','Remarks'],
  'M_Hotels':   ['Hotel_Name','Phone','Address','Surcharge_Area'],
  'M_PriceClient': ['Agency','Course','max_hours','seats_21_rate','seats_21_ot',
                    'seats_25_rate','seats_25_ot','seats_40_rate','seats_40_ot',
                    'seats_50_rate','seats_50_ot'],
  'M_PriceDriver': ['Course','max_hours','seats_21_base','seats_21_ot',
                    'seats_25_base','seats_25_ot','seats_40_base','seats_40_ot',
                    'seats_50_base','seats_50_ot'],
  // ── 신규 ──
  'M_PriceSub': ['SubCo','Course','max_hours','seats_21_rate','seats_21_ot',
                   'seats_25_rate','seats_25_ot','seats_40_rate','seats_40_ot',
                   'seats_50_rate','seats_50_ot'],
  'Sub_Rates':  ['Rego','Tour','seats_21','seats_25','seats_40','seats_50'],
  'Ledger':     ['WeekStart','Date','Rego','Tour','TA','SubTotal','MyDr','Extra',
                 'OT','Trailer','Hotel','Note'],
  'Wages':      ['Driver','WeekStart','Date','Amount','Method','Note'],
  'Notices':    ['ID','Title','Content','Type','Date','Active']
};

// ── 탭 색상 ──
const TAB_COLORS = {
  'M_Vehicles':'#d97706','M_Drivers':'#1a56db','M_Clients':'#7e3af2',
  'M_Guides':'#0e9f6e','M_Hotels':'#e02424','M_PriceClient':'#0694a2',
  'M_PriceDriver':'#057a55','M_PriceSub':'#7c3aed','Sub_Rates':'#b45309','Ledger':'#1e40af','Wages':'#065f46',
  'MOT_Report':'#be185d',
  'Notices':'#0369a1'
};

function cors(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════
// GET 핸들러
// ═══════════════════════════════════════════════════
function doGet(e) {
  try {
    const action = e.parameter.action || 'ping';
    if (action === 'ping')            return cors({ok:true, msg:'DC Fleet API ready'});
    if (action === 'get_reports')     return cors(getReports(e.parameter.sheet, e.parameter.driver));
    if (action === 'get_master')      return cors(getMaster(e.parameter.sheet));
    if (action === 'get_all_masters') return cors(getAllMasters());
    // ── 신규 ──
    if (action === 'get_sub_rates')   return cors(getSubRatesSheet());
    if (action === 'get_price_sub')   return cors(getPriceSubSheet());
    if (action === 'get_ledger')      return cors(getLedgerSheet());
    if (action === 'get_wages')       return cors(getWagesSheet(e.parameter.driver));
    if (action === 'get_mot_reports') return cors(getReports('MOT_Report', e.parameter.driver));
    if (action === 'get_notices')     return cors(getNoticesSheet());
    if (action === 'get_active_regos') return cors(getActiveRegos());

    return cors({ok:false, msg:'Unknown action: ' + action});
  } catch(err) {
    return cors({ok:false, error: err.toString()});
  }
}

// ═══════════════════════════════════════════════════
// POST 핸들러
// ═══════════════════════════════════════════════════
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action;

    // 리포트 저장
    if (action === 'save_report')        return cors(saveReport('Daily_Report',  payload.data));
    if (action === 'update_report')      return cors(updateReport(payload.sheet, payload.rowIndex, payload.data));
    if (action === 'delete_report')      return cors(deleteReport(payload.sheet, payload.rowIndex));
    if (action === 'save_predeparture')  return cors(saveReport('Pre_Departure', payload.data));
    if (action === 'save_endofshift')    return cors(saveReport('End_of_Shift',  payload.data));
    if (action === 'submit_mot')         return cors(saveReport('MOT_Report',    payload.data));

    // 마스터 CRUD
    if (action === 'add_master')         return cors(addMasterRow(payload.sheet, payload.data));
    if (action === 'add_price_client_agency') return cors(addPriceClientAgency(payload.agency, payload.rows));
    if (action === 'update_master')      return cors(updateMasterRow(payload.sheet, payload.rowIndex, payload.data));
    if (action === 'delete_master')      return cors(deleteMasterRow(payload.sheet, payload.rowIndex));
    if (action === 'replace_master')     return cors(replaceMasterSheet(payload.sheet, payload.rows));
    if (action === 'init_masters')       return cors(initAllMasters());

    // ── 신규: 서브 요금표 ──
    if (action === 'replace_sub_rates')  return cors(replaceMasterSheet('Sub_Rates', payload.rows));
    if (action === 'replace_price_sub')  return cors(replaceMasterSheet('M_PriceSub', payload.rows));

    // ── 신규: 정산 내역 (Ledger) ──
    if (action === 'add_ledger')         return cors(addMasterRow('Ledger', payload.data));
    if (action === 'update_ledger')      return cors(updateMasterRow('Ledger', payload.rowIndex, payload.data));
    if (action === 'delete_ledger')      return cors(deleteMasterRow('Ledger', payload.rowIndex));
    if (action === 'replace_ledger')     return cors(replaceMasterSheet('Ledger', payload.rows));

    // ── 신규: 드라이버 급여 지급 ──
    if (action === 'add_wage')           return cors(addMasterRow('Wages', payload.data));
    if (action === 'update_wage')        return cors(updateMasterRow('Wages', payload.rowIndex, payload.data));
    if (action === 'delete_wage')        return cors(deleteMasterRow('Wages', payload.rowIndex));
    if (action === 'replace_wages')      return cors(replaceMasterSheet('Wages', payload.rows));

    if (action === 'save_notices')       return cors(replaceMasterSheet('Notices', payload.rows));

    // ── 드라이버 정보 수정 ──
    if (action === 'update_driver_pin')  return cors(updateDriverPin(payload.driverName, payload.pin));
    if (action === 'update_driver_info') return cors(updateDriverInfo(payload.driverName, payload.data));

    return cors({ok:false, msg:'Unknown action: ' + action});
  } catch(err) {
    return cors({ok:false, error: err.toString()});
  }
}

// ═══════════════════════════════════════════════════
// 리포트 저장
// ═══════════════════════════════════════════════════
function saveReport(sheetName, data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  const headers = REPORT_HEADERS[sheetName];
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1,1,1,headers.length).setValues([headers])
      .setBackground('#1a56db').setFontColor('white').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  const row = headers.map(h => data[h] !== undefined ? data[h] : '');
  sheet.appendRow(row);
  return {ok:true, sheet:sheetName, row:sheet.getLastRow()};
}

// ═══════════════════════════════════════════════════
// 마스터 읽기
// ═══════════════════════════════════════════════════

function updateReport(sheetName, rowIndex, data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return {ok: false, msg: sheetName + ' 시트 없음'};
  const headers = REPORT_HEADERS[sheetName];
  if (!headers) return {ok: false, msg: 'Unknown sheet: ' + sheetName};
  const row = headers.map(h => data[h] !== undefined ? data[h] : '');
  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  return {ok: true};
}

function deleteReport(sheetName, rowIndex) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return {ok: false, msg: sheetName + ' 시트 없음'};
  sheet.deleteRow(rowIndex);
  return {ok: true};
}

function getMaster(sheetName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return {ok:false, msg:sheetName+' 시트 없음'};
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return {ok:true, sheet:sheetName, rows:[]};
  const headers = data[0];
  const rows = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h,i) => { obj[h] = row[i]; });
    return obj;
  });
  return {ok:true, sheet:sheetName, rows};
}

function getAllMasters() {
  const sheets = ['M_Vehicles','M_Drivers','M_Clients','M_Guides','M_Hotels','M_PriceClient','M_PriceDriver','M_PriceSub'];
  const result = {};
  sheets.forEach(name => {
    const r = getMaster(name);
    result[name] = r.ok ? r.rows : [];
  });
  return {ok:true, data:result};
}

// ═══════════════════════════════════════════════════
// 신규: Sub_Rates, Ledger, Wages 읽기
// ═══════════════════════════════════════════════════
function getSubRatesSheet() {
  const r = getMaster('Sub_Rates');
  return {ok:r.ok, rows:r.rows||[]};
}

function getPriceSubSheet() {
  const r = getMaster('M_PriceSub');
  return {ok:r.ok, rows:r.rows||[]};
}

function getLedgerSheet() {
  const r = getMaster('Ledger');
  return {ok:r.ok, rows:r.rows||[]};
}

function getWagesSheet(driver) {
  const r = getMaster('Wages');
  let rows = r.rows || [];
  if (driver) rows = rows.filter(row => row.Driver === driver);
  return {ok:r.ok, rows};
}

// ═══════════════════════════════════════════════════
// 마스터 쓰기
// ═══════════════════════════════════════════════════
function ensureSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    const headers = MASTER_HEADERS[sheetName];
    if (headers) {
      const color = TAB_COLORS[sheetName] || '#1a56db';
      sheet.getRange(1,1,1,headers.length).setValues([headers])
        .setBackground(color).setFontColor('white').setFontWeight('bold');
      sheet.setFrozenRows(1);
      sheet.setTabColor(color);
    }
  }
  return sheet;
}


// ── 새 여행사 M_PriceClient 일괄 생성 ──
function addPriceClientAgency(agencyName, rows) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ensureSheet(ss, 'M_PriceClient');
  const headers = MASTER_HEADERS['M_PriceClient'];

  // 기존 Agency+Course 중복 체크
  const lastRow = sheet.getLastRow();
  const existing = new Set();
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    data.forEach(r => { if (r[0] && r[1]) existing.add(r[0] + '||' + r[1]); });
  }

  const newRows = (rows || []).filter(r => !existing.has(r.Agency + '||' + r.Course));
  if (newRows.length === 0) return {ok: true, added: 0, msg: '중복 없음'};

  const values = newRows.map(r => headers.map(h => r[h] !== undefined ? r[h] : ''));
  sheet.getRange(lastRow + 1, 1, values.length, headers.length).setValues(values);
  return {ok: true, added: newRows.length};
}

function addMasterRow(sheetName, data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ensureSheet(ss, sheetName);
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => data[h] !== undefined ? data[h] : '');
  sheet.appendRow(row);
  return {ok:true, row:sheet.getLastRow()};
}

function updateMasterRow(sheetName, rowIndex, data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ensureSheet(ss, sheetName);
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => data[h] !== undefined ? data[h] : '');
  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  return {ok:true};
}

function deleteMasterRow(sheetName, rowIndex) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return {ok:false, msg:'시트 없음'};
  sheet.deleteRow(rowIndex);
  return {ok:true};
}

function replaceMasterSheet(sheetName, rows) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ensureSheet(ss, sheetName);
  const headers = MASTER_HEADERS[sheetName];
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  if (rows && rows.length > 0) {
    const data = rows.map(obj => headers.map(h => obj[h] !== undefined ? obj[h] : ''));
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  }
  return {ok:true, count: rows ? rows.length : 0};
}

// ═══════════════════════════════════════════════════
// 리포트 읽기
// ═══════════════════════════════════════════════════
function getReports(sheetName, driver) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return {ok:false, msg:sheetName+' 시트 없음'};
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return {ok:true, rows:[]};
  const headers = data[0];
  // 날짜 셀(Date 오브젝트)을 dd/mm/yyyy 문자열로 변환
  const DATE_FIELDS = ['Date','Submitted','License_Expiry','Authority_Expiry','Rego_Date','HVIS_Date'];
  function fmtCell(h, v) {
    if (DATE_FIELDS.indexOf(h) !== -1 && v instanceof Date && !isNaN(v)) {
      // AEST(UTC+10) 기준으로 변환
      const local = new Date(v.getTime() + 10 * 3600 * 1000);
      const dd = String(local.getUTCDate()).padStart(2,'0');
      const mm = String(local.getUTCMonth()+1).padStart(2,'0');
      const yyyy = local.getUTCFullYear();
      return dd + '/' + mm + '/' + yyyy;
    }
    return v;
  }
  let rows = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h,i) => { obj[h] = fmtCell(h, row[i]); });
    return obj;
  });
  if (driver) rows = rows.filter(r => r.Driver === driver);
  return {ok:true, rows};
}

// ═══════════════════════════════════════════════════
// 마스터 시트 초기화 (신규 시트 포함)
// ═══════════════════════════════════════════════════
function initAllMasters() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const results = [];
  Object.keys(MASTER_HEADERS).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (sheet) { results.push({sheet:name, status:'skipped'}); return; }
    sheet = ss.insertSheet(name);
    const headers = MASTER_HEADERS[name];
    const color = TAB_COLORS[name] || '#1a56db';
    sheet.getRange(1,1,1,headers.length).setValues([headers])
      .setBackground(color).setFontColor('white').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setTabColor(color);
    results.push({sheet:name, status:'created'});
  });
  return {ok:true, results};
}



// ═══════════════════════════════════════════════════
// 드라이버 PIN / 정보 수정
// ═══════════════════════════════════════════════════
function updateDriverPin(driverName, pin) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('M_Drivers');
  if (!sheet) return {ok: false, msg: 'M_Drivers 시트 없음'};
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameENIdx = headers.indexOf('Name_EN');
  const nameKRIdx = headers.indexOf('Name_KR');
  const pinIdx    = headers.indexOf('PIN');
  if (pinIdx === -1) return {ok: false, msg: 'PIN 컬럼 없음'};
  for (let r = 1; r < data.length; r++) {
    if (data[r][nameENIdx] === driverName || data[r][nameKRIdx] === driverName) {
      sheet.getRange(r + 1, pinIdx + 1).setValue(pin);
      return {ok: true};
    }
  }
  return {ok: false, msg: '드라이버 없음: ' + driverName};
}

function updateDriverInfo(driverName, data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('M_Drivers');
  if (!sheet) return {ok: false, msg: 'M_Drivers 시트 없음'};
  const sheetData = sheet.getDataRange().getValues();
  const headers = sheetData[0];
  const nameENIdx = headers.indexOf('Name_EN');
  const nameKRIdx = headers.indexOf('Name_KR');

  const fieldMap = {
    nameEN: 'Name_EN', nameKR: 'Name_KR', mobile: 'Mobile_1',
    licClass: 'License_Class', licNo: 'License_No', licExp: 'License_Expiry',
    authNo: 'Authority_No', authExp: 'Authority_Expiry',
    nokName: 'NEXT_OF_KIN', address: 'Address', suburb: 'Suburb'
  };

  for (let r = 1; r < sheetData.length; r++) {
    if (sheetData[r][nameENIdx] === driverName || sheetData[r][nameKRIdx] === driverName) {
      Object.entries(data).forEach(([key, val]) => {
        const col = fieldMap[key];
        if (col) {
          const colIdx = headers.indexOf(col);
          if (colIdx !== -1) sheet.getRange(r + 1, colIdx + 1).setValue(val);
        }
      });
      return {ok: true};
    }
  }
  return {ok: false, msg: '드라이버 없음: ' + driverName};
}

// ═══════════════════════════════════════════════════
// 공지사항 & 활성 운행 조회
// ═══════════════════════════════════════════════════
function getNoticesSheet() {
  const r = getMaster('Notices');
  return {ok: r.ok, rows: r.rows || []};
}

function getActiveRegos() {
  // Pre_Departure는 있지만 End_of_Shift가 없는 운행 = 현재 운행 중
  // ★ 오늘 날짜 기준으로만 체크 (과거 기록 제외)
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const preSheet = ss.getSheetByName('Pre_Departure');
  const eosSheet = ss.getSheetByName('End_of_Shift');
  if (!preSheet) return {ok: true, regos: []};

  const preData = preSheet.getDataRange().getValues();
  if (preData.length < 2) return {ok: true, regos: []};
  const preHeaders = preData[0];

  // 오늘 날짜 문자열 (Sydney 시간 기준) - "DD/MM/YYYY" 또는 "YYYY-MM-DD" 양쪽 대응
  const now = new Date();
  const sydOffset = 10 * 60; // AEST UTC+10 (DST는 +11이지만 보수적으로 처리)
  const utc = now.getTime() + now.getTimezoneOffset() * 60000;
  const syd = new Date(utc + sydOffset * 60000);
  const yy = syd.getFullYear();
  const mm = String(syd.getMonth()+1).padStart(2,'0');
  const dd = String(syd.getDate()).padStart(2,'0');
  const todayISO = yy+'-'+mm+'-'+dd;           // "2025-03-13"
  const todayDMY = dd+'/'+mm+'/'+yy;           // "13/03/2025"
  const todayMDY = mm+'/'+dd+'/'+yy;           // "03/13/2025"

  function isToday(val) {
    const s = String(val).trim().replace(/\s+.*/,''); // 시간 부분 제거
    return s === todayISO || s === todayDMY || s === todayMDY;
  }

  const preRows = preData.slice(1).map(row => {
    const obj = {};
    preHeaders.forEach((h,i) => obj[h] = row[i]);
    return obj;
  }).filter(r => isToday(r.Date)); // ★ 오늘 것만

  // EoS 데이터 수집 (오늘 것만)
  const eosSet = new Set();
  if (eosSheet && eosSheet.getLastRow() > 1) {
    const eosData = eosSheet.getDataRange().getValues();
    const eosH = eosData[0];
    eosData.slice(1).forEach(row => {
      const obj = {};
      eosH.forEach((h,i) => obj[h] = row[i]);
      if (isToday(obj.Date)) {
        // Rego+Date 조합 (Driver는 빈값일 수 있으므로 Rego+Date로만 판단)
        eosSet.add(String(obj.Rego).trim()+'|'+String(obj.Date).trim());
      }
    });
  }

  // Pre_Departure는 있지만 End_of_Shift 없는 것 = 운행 중
  const active = [];
  const seen = new Set(); // 동일 Rego 중복 방지
  preRows.forEach(r => {
    const regoKey = String(r.Rego).trim()+'|'+String(r.Date).trim();
    if (!eosSet.has(regoKey) && !seen.has(regoKey)) {
      seen.add(regoKey);
      const driverName = String(r.Driver || '').trim();
      active.push({
        driver: driverName || '알 수 없음',
        rego: String(r.Rego).trim(),
        date: String(r.Date).trim()
      });
    }
  });

  return {ok: true, regos: active};
}

/**
 * M_Guides, M_Drivers, M_Hotels 시트의 전화번호 컬럼 앞자리 0 복원
 */
function fixPhoneNumbers() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const targets = [
    { sheet: 'M_Guides',  col: 'Mobile' },
    { sheet: 'M_Drivers', col: 'Phone'  },
    { sheet: 'M_Hotels',  col: 'Phone'  },
  ];
  let totalFixed = 0;
  targets.forEach(({ sheet: sheetName, col: colName }) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) { Logger.log(sheetName + ' 시트 없음, 건너뜀'); return; }
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
    const headers = data[0];
    const colIdx = headers.indexOf(colName);
    if (colIdx === -1) { Logger.log(sheetName + '.' + colName + ' 컬럼 없음'); return; }
    let fixed = 0;
    for (let r = 1; r < data.length; r++) {
      const val = data[r][colIdx];
      if (val === '' || val === null) continue;
      let s = String(val).replace(/\.0+$/, '').replace(/\s/g, '').replace(/[^0-9]/g, '');
      if (s.length === 9) {
        s = '0' + s;
        sheet.getRange(r + 1, colIdx + 1).setValue(s).setNumberFormat('@');
        fixed++;
      } else if (s.length === 10 && !String(val).startsWith('0')) {
        sheet.getRange(r + 1, colIdx + 1).setValue(s).setNumberFormat('@');
        fixed++;
      }
    }
    Logger.log(sheetName + '.' + colName + ': ' + fixed + '개 수정');
    totalFixed += fixed;
  });
  Logger.log('총 ' + totalFixed + '개 전화번호 수정 완료');
  SpreadsheetApp.getUi().alert('완료! ' + totalFixed + '개 전화번호 앞자리 0 복원됨');
}
