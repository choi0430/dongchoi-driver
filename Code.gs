// ═══════════════════════════════════════════════════
// DC Fleet — Google Sheets Backend (Apps Script)
// Spreadsheet: Dong Choi Pty Ltd - Driver Reports
// ═══════════════════════════════════════════════════

const SHEET_ID = '1kUU-_-IFJkKd97O-Im-A6xojsafYG-0njVyRKmSLKeE';

// ── 리포트 시트 헤더 ──
const REPORT_HEADERS = {
  'Daily_Report':   ['Submitted','Driver','Date','Rego','Agency','Attraction','Pickup','Dropoff',
                     'KM_Start','KM_End','Time_Start','Time_End','Guide','Tour_Code',
                     'SVC_Label','SVC_Charge','Hotel_Surcharge','Dist_Surcharge',
                     'OT','Trailer','Total_TA','DR_Cost','Remarks'],
  'Pre_Departure':  ['Submitted','Driver','Date','Rego','Seats','Start_KM','Fuel','Start_Time',
                     'Check_Results','Remarks'],
  'End_of_Shift':   ['Submitted','Driver','Date','Rego','End_KM','End_Time','Fuel_End','Remarks']
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
                    'seats_50_base','seats_50_ot']
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
    if (action === 'save_predeparture')  return cors(saveReport('Pre_Departure', payload.data));
    if (action === 'save_endofshift')    return cors(saveReport('End_of_Shift',  payload.data));

    // 마스터 CRUD
    if (action === 'add_master')         return cors(addMasterRow(payload.sheet, payload.data));
    if (action === 'update_master')      return cors(updateMasterRow(payload.sheet, payload.rowIndex, payload.data));
    if (action === 'delete_master')      return cors(deleteMasterRow(payload.sheet, payload.rowIndex));
    if (action === 'replace_master')     return cors(replaceMasterSheet(payload.sheet, payload.rows));
    if (action === 'init_masters')       return cors(initAllMasters());

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

// 모든 마스터 한 번에 읽기
function getAllMasters() {
  const sheets = ['M_Vehicles','M_Drivers','M_Clients','M_Guides','M_Hotels','M_PriceClient','M_PriceDriver'];
  const result = {};
  sheets.forEach(name => {
    const r = getMaster(name);
    result[name] = r.ok ? r.rows : [];
  });
  return {ok:true, data:result};
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
      sheet.getRange(1,1,1,headers.length).setValues([headers])
        .setBackground('#1a56db').setFontColor('white').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
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

// 시트 전체 내용을 rows 배열로 교체 (PC/PD 등 복잡한 구조 저장용)
function replaceMasterSheet(sheetName, rows) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ensureSheet(ss, sheetName);
  const headers = MASTER_HEADERS[sheetName];
  // 기존 데이터 행 삭제
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  // 새 데이터 삽입
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
  let rows = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h,i) => obj[h] = row[i]);
    return obj;
  });
  if (driver) rows = rows.filter(r => r.Driver === driver);
  return {ok:true, rows};
}

// ═══════════════════════════════════════════════════
// 마스터 시트 초기화 (최초 1회)
// ═══════════════════════════════════════════════════
function initAllMasters() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const colors = {
    'M_Vehicles':'#d97706','M_Drivers':'#1a56db','M_Clients':'#7e3af2',
    'M_Guides':'#0e9f6e','M_Hotels':'#e02424','M_PriceClient':'#0694a2','M_PriceDriver':'#057a55'
  };
  const results = [];
  Object.keys(MASTER_HEADERS).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (sheet) { results.push({sheet:name, status:'skipped'}); return; }
    sheet = ss.insertSheet(name);
    const headers = MASTER_HEADERS[name];
    sheet.getRange(1,1,1,headers.length).setValues([headers])
      .setBackground(colors[name]||'#1a56db').setFontColor('white').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setTabColor(colors[name]||'#1a56db');
    results.push({sheet:name, status:'created'});
  });
  return {ok:true, results};
}

/**
 * M_Guides, M_Drivers, M_Hotels 시트의 전화번호 컬럼 앞자리 0 복원
 * Apps Script 편집기에서 한 번만 실행하세요.
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
        // 셀에 텍스트로 저장 (앞에 ' 붙여서 강제 텍스트)
        sheet.getRange(r + 1, colIdx + 1).setValue(s).setNumberFormat('@');
        fixed++;
      } else if (s.length === 10 && !String(val).startsWith('0')) {
        // 10자리인데 텍스트 아닌 경우도 텍스트로 재저장
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
