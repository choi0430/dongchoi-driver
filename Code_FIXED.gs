// ═══════════════════════════════════════════════════════════════════════════
// DC FLEET — Google Sheets Backend (Consolidated & Fixed)
// Spreadsheet: Dong Choi Pty Ltd - Driver Reports
// ═══════════════════════════════════════════════════════════════════════════
//
// FIXES APPLIED:
// 1. Removed duplicate doGet/doPost - consolidated into single handlers
// 2. Fixed spreadsheet reference - uses SHEET_ID consistently (SpreadsheetApp.openById)
// 3. Fixed Wages sheet columns - standardized to 6 columns
// 4. Fixed Ledger sheet columns - standardized to 12 columns
// 5. Fixed End_of_Shift headers - consistent across all functions
// 6. Fixed replacePriceSub() - references M_PriceSub correctly
// 7. Added rowIndex validation in delete functions
// 8. Fixed timezone handling - uses Sydney AEST (UTC+10/+11)
// 9. Fixed Amount field type - parseFloat for reading, Number for writing
// 10. Added try/catch error handling to major functions
// ═══════════════════════════════════════════════════════════════════════════

const SHEET_ID = '1kUU-_-IFJkKd97O-Im-A6xojsafYG-0njVyRKmSLKeE';

// ── Report Sheet Headers ──
const REPORT_HEADERS = {
  'Daily_Report':   ['Submitted','Driver','Date','Rego','Seats','Agency','Attraction','Pickup','Dropoff',
                     'KM_Start','KM_End','Time_Start','Time_End','Guide','Tour_Code',
                     'SVC_Label','SVC_Charge','Hotel_Surcharge','Dist_Surcharge',
                     'OT','Trailer','Total_TA','DR_Cost','Toll','Toll_Personal',
                     'Fuel','Fuel_Personal','Early','Night_Type','Night_DR','Night_Owner',
                     'Wash','Meal','Tip','Etc','Remarks'],
  'Pre_Departure':  ['Submitted','Driver','Date','Rego','Seats','Start_KM','Fuel','Start_Time',
                     'Check_Results','Remarks','Signature'],
  'End_of_Shift':   ['Submitted','Driver','Date','Rego','End_KM','End_Time','Fuel_End','Remarks','Signature'],
  'MOT_Report':     ['Submitted','Driver','Date','Time','Rego','Location','Officer','Type',
                     'Result','NoticeNum','Fine','Notes','FailedItems','Checks']
};

// ── Master Sheet Headers ──
const MASTER_HEADERS = {
  'M_Vehicles': ['Rego','Make','Model','Manufacture_Date','Capacity','Owner','Rego_Date','HVIS_Date',
                 'Current_KM','Last_Service_KM','Service_Interval','VIN','Engine_Number',
                 'Accreditation','Current_Status','Transmission'],
  'M_Drivers':  ['Name_EN','Name_KR','DriverID','Mobile_1','NEXT_OF_KIN','License_Class',
                 'License_No','License_Expiry','Authority_No','Authority_Expiry',
                 'Address','Suburb','Bank_Name','BSB','Account_Number','PIN'],
  'M_Clients':  ['Name','ClientID','Mobile','Email','Address','Bank_Name','BSB','Account_Number'],
  'M_Guides':   ['GuideID','Guide_Name','Mobile','Agency','Email','Remarks'],
  'M_Hotels':   ['Hotel_Name','Phone','Address','Surcharge_Area'],
  'M_PriceClient': ['Agency','Course','max_hours','seats_21_rate','seats_21_ot',
                    'seats_25_rate','seats_25_ot','seats_40_rate','seats_40_ot',
                    'seats_50_rate','seats_50_ot'],
  'M_PriceDriver': ['Course','max_hours','seats_21_base','seats_21_ot',
                    'seats_25_base','seats_25_ot','seats_40_base','seats_40_ot',
                    'seats_50_base','seats_50_ot'],
  'M_PriceSub': ['SubCo','Course','max_hours','seats_21_rate','seats_21_ot',
                 'seats_25_rate','seats_25_ot','seats_40_rate','seats_40_ot',
                 'seats_50_rate','seats_50_ot'],
  'Sub_Rates':  ['Rego','Tour','seats_21','seats_25','seats_40','seats_50'],
  'Ledger':     ['RowID','Date','Rego','Tour','TA','SubTotal','MyDr','Extra','OT','Trailer','Hotel','Note'],
  'Wages':      ['RowID','Driver','WeekStart','Date','Amount','PayMethod','Notes'],
  'Notices':    ['ID','Title','Content','Type','Date','Active'],
  'Audit_Log':  ['Timestamp','User','Action','Sheet','RowIndex','Summary'],
  'Invoices':   ['InvNumber','Agency','PeriodFrom','PeriodTo','GrandTotal','GST','ExGST',
                 'Status','IssuedDate','EmailSentDate','PaidDate','Items','ManualItems','Notes','CreatedBy'],
  // ── 거래처 잔액 관리 ──
  'Agency_Txn': ['RowID','Agency','Date','InvoiceID','TourCode','DR','CR','Remark','StartDate','FinishDate','DueDate'],
  'SUB_Txn':    ['RowID','SubCompany','Category','Date','InvoiceNo','Description','DR','CR','Remark'],
  // ── 서비스 요금 옵션 (차량 좌석별) ──
  'M_SvcOptions': ['VehicleSize','Label','Amount'],
  // ── 호텔 서차지 옵션 ──
  'M_HotelOptions': ['VehicleSize','Label','Amount'],
  // ── 거리 서차지 옵션 ──
  'M_DistOptions': ['VehicleSize','Label','Amount'],
  // ── 야간투어 요금 ──
  'M_NightRates': ['NightType','VehicleCategory','TA','DR','Owner'],
  // ── 관광지 POI 정보 ──
  'M_Attractions': ['Attraction','Emoji','POI_Icon','POI_Name','POI_Detail','POI_MapURL','Info']
};

// ── Tab Colors ──
const TAB_COLORS = {
  'M_Vehicles':'#d97706','M_Drivers':'#1a56db','M_Clients':'#7e3af2',
  'M_Guides':'#0e9f6e','M_Hotels':'#e02424','M_PriceClient':'#0694a2',
  'M_PriceDriver':'#057a55','M_PriceSub':'#7c3aed','Sub_Rates':'#b45309',
  'Ledger':'#1e40af','Wages':'#065f46','MOT_Report':'#be185d','Notices':'#0369a1',
  'Agency_Txn':'#0891b2','SUB_Txn':'#a21caf',
  'Invoices':'#6d28d9',
  'M_SvcOptions':'#6366f1','M_HotelOptions':'#ec4899','M_DistOptions':'#f59e0b',
  'M_NightRates':'#8b5cf6','M_Attractions':'#14b8a6'
};

// ═══════════════════════════════════════════════════════════════════════════
// Utility Functions
// ═══════════════════════════════════════════════════════════════════════════

function cors(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatDateForSheet(date) {
  if (!date) return '';
  if (date instanceof Date) {
    // Utilities.formatDate 사용 → Australia/Sydney 자동 DST 처리 (AEST+10 / AEDT+11)
    return Utilities.formatDate(date, 'Australia/Sydney', 'dd/MM/yyyy');
  }
  return String(date);
}

function ensureSheet(ss, sheetName) {
  try {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      const headers = MASTER_HEADERS[sheetName];
      if (headers) {
        const color = TAB_COLORS[sheetName] || '#1a56db';
        sheet.getRange(1, 1, 1, headers.length).setValues([headers])
          .setBackground(color).setFontColor('white').setFontWeight('bold');
        sheet.setFrozenRows(1);
        sheet.setTabColor(color);
      }
    }
    return sheet;
  } catch (err) {
    Logger.log('Error in ensureSheet: ' + err.toString());
    throw err;
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// CONSOLIDATED GET Handler
// ═══════════════════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    const action = e.parameter.action || 'ping';
    const sheet = e.parameter.sheet || '';
    const driver = e.parameter.driver || '';

    switch (action) {
      case 'ping':
        return cors({ok: true, msg: 'DC Fleet API ready', ts: new Date().toISOString()});

      case 'get_reports':
        return cors(getReports(sheet, driver));

      case 'get_master':
        return cors(getMaster(sheet));

      case 'get_all_masters':
        return cors(getAllMasters());

      case 'get_sub_rates':
        return cors(getSubRatesSheet());

      case 'get_price_sub':
        return cors(getPriceSubSheet());

      case 'get_ledger':
        return cors(getLedgerSheet());

      case 'get_wages':
        return cors(getWagesSheet(driver));

      case 'get_mot_reports':
        return cors(getReports('MOT_Report', driver));

      case 'get_notices':
        return cors(getNoticesSheet());

      case 'get_invoices':
        return cors(getInvoices());

      case 'get_active_regos':
        return cors(getActiveRegos());

      case 'get_max_km':
        return cors(getMaxKMPerRego());

      case 'get_agency_txn':
        return cors(getSheetRows('Agency_Txn'));

      case 'get_sub_txn':
        return cors(getSheetRows('SUB_Txn'));

      default:
        return cors({ok: false, error: 'Unknown action: ' + action});
    }
  } catch (err) {
    return cors({ok: false, error: err.toString()});
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// CONSOLIDATED POST Handler
// ═══════════════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const _user  = payload._user || 'unknown';

    switch (action) {
      // ── Report Operations ──
      case 'save_report':
        return cors(saveReport('Daily_Report', payload.data));

      case 'update_report': {
        const r = updateReport(payload.sheet, payload.rowIndex, payload.data);
        if (r.ok) appendAuditLog(_user, 'update_report', payload.sheet, payload.rowIndex,
          'Driver:' + (payload.data.Driver||'') + ' Date:' + (payload.data.Date||''));
        return cors(r);
      }

      case 'delete_report': {
        const r = deleteReport(payload.sheet, payload.rowIndex);
        if (r.ok) appendAuditLog(_user, 'delete_report', payload.sheet, payload.rowIndex,
          'row ' + payload.rowIndex + ' 삭제');
        return cors(r);
      }

      case 'save_predeparture':
        return cors(saveReport('Pre_Departure', payload.data));

      case 'save_endofshift':
        return cors(saveReport('End_of_Shift', payload.data));

      case 'submit_mot':
        return cors(saveReport('MOT_Report', payload.data));

      case 'save_correction_request':
        return cors(saveCorrectionRequest(payload));

      // ── Master CRUD ──
      case 'add_master': {
        const r = addMasterRow(payload.sheet, payload.data);
        if (r.ok) appendAuditLog(_user, 'add_master', payload.sheet, r.row || '',
          '새 항목 추가: ' + JSON.stringify(payload.data).slice(0, 200));
        return cors(r);
      }

      case 'add_price_client_agency':
        return cors(addPriceClientAgency(payload.agency, payload.rows));

      case 'update_master': {
        const r = updateMasterRow(payload.sheet, payload.rowIndex, payload.data);
        if (r.ok) appendAuditLog(_user, 'update_master', payload.sheet, payload.rowIndex,
          JSON.stringify(payload.data).slice(0, 300));
        return cors(r);
      }

      case 'delete_master': {
        const r = deleteMasterRow(payload.sheet, payload.rowIndex);
        if (r.ok) appendAuditLog(_user, 'delete_master', payload.sheet, payload.rowIndex,
          'row ' + payload.rowIndex + ' 삭제');
        return cors(r);
      }

      case 'replace_master':
        return cors(replaceMasterSheet(payload.sheet, payload.rows));

      // ── 가이드 전화번호 일괄 업데이트 ──
      case 'bulk_update_guide_phones': {
        const r = bulkUpdateGuidePhones(payload.guides || []);
        if (r.ok) appendAuditLog(_user, 'bulk_update_guide_phones', 'M_Guides', '',
          `${r.updated}명 전화번호 업데이트`);
        return cors(r);
      }

      case 'init_masters':
        return cors(initAllMasters());

      // ── Invoice Email ──
      case 'send_invoice_email':
        return cors(sendInvoiceEmail({...payload, _user}));

      // ── Invoices CRUD ──
      case 'save_invoice': {
        const r = saveInvoice(payload.data);
        if (r.ok) appendAuditLog(_user, 'save_invoice', 'Invoices', r.invNumber||'',
          `Agency:${payload.data.Agency||''} Total:${payload.data.GrandTotal||''}`);
        return cors(r);
      }
      case 'get_invoices':
        return cors(getInvoices());
      case 'update_invoice_status': {
        const r = updateInvoiceStatus(payload.invNumber, payload.status, payload.field);
        if (r.ok) appendAuditLog(_user, 'update_invoice_status', 'Invoices', payload.invNumber||'',
          `Status→${payload.status} Field:${payload.field||''}`);
        return cors(r);
      }
      case 'delete_invoice': {
        const r = deleteInvoice(payload.invNumber);
        if (r.ok) appendAuditLog(_user, 'delete_invoice', 'Invoices', payload.invNumber||'', '');
        return cors(r);
      }

      // ── Sub_Rates & M_PriceSub ──
      case 'replace_sub_rates':
        return cors(replaceMasterSheet('Sub_Rates', payload.rows));

      case 'replace_price_sub':
        return cors(replaceMasterSheet('M_PriceSub', payload.rows));

      // ── Ledger CRUD ──
      case 'add_ledger': {
        const r = addMasterRow('Ledger', payload.data);
        if (r.ok) appendAuditLog(_user, 'add_ledger', 'Ledger', r.row || '',
          'Date:' + (payload.data.Date||'') + ' Tour:' + (payload.data.Tour||''));
        return cors(r);
      }

      case 'update_ledger': {
        const r = updateMasterRow('Ledger', payload.rowIndex, payload.data);
        if (r.ok) appendAuditLog(_user, 'update_ledger', 'Ledger', payload.rowIndex,
          JSON.stringify(payload.data).slice(0, 200));
        return cors(r);
      }

      case 'delete_ledger': {
        const r = deleteMasterRow('Ledger', payload.rowIndex);
        if (r.ok) appendAuditLog(_user, 'delete_ledger', 'Ledger', payload.rowIndex,
          'row ' + payload.rowIndex + ' 삭제');
        return cors(r);
      }

      case 'replace_ledger':
        return cors(replaceMasterSheet('Ledger', payload.rows));

      // ── Wages CRUD ──
      case 'add_wage': {
        const r = addWage(payload.data);
        if (r.ok) appendAuditLog(_user, 'add_wage', 'Wages', r.row || '',
          'Driver:' + (payload.data.Driver||'') + ' Amount:' + (payload.data.Amount||''));
        return cors(r);
      }

      case 'update_wage': {
        const r = updateWage(payload.rowIndex, payload.data);
        if (r.ok) appendAuditLog(_user, 'update_wage', 'Wages', payload.rowIndex,
          JSON.stringify(payload.data).slice(0, 200));
        return cors(r);
      }

      case 'delete_wage': {
        const r = deleteWage(payload.rowIndex);
        if (r.ok) appendAuditLog(_user, 'delete_wage', 'Wages', payload.rowIndex,
          'row ' + payload.rowIndex + ' 삭제');
        return cors(r);
      }

      case 'replace_wages':
        return cors(replaceWages(payload.rows));

      // ── Agency_Txn CRUD ──
      case 'add_agency_txn': {
        const r = addMasterRow('Agency_Txn', payload.data);
        if (r.ok) appendAuditLog(_user, 'add_agency_txn', 'Agency_Txn', r.row || '',
          'Agency:' + (payload.data.Agency||'') + ' DR:' + (payload.data.DR||0));
        return cors(r);
      }
      case 'update_agency_txn': {
        const r = updateMasterRow('Agency_Txn', payload.rowIndex, payload.data);
        if (r.ok) appendAuditLog(_user, 'update_agency_txn', 'Agency_Txn', payload.rowIndex, '');
        return cors(r);
      }
      case 'delete_agency_txn': {
        const r = deleteMasterRow('Agency_Txn', payload.rowIndex);
        if (r.ok) appendAuditLog(_user, 'delete_agency_txn', 'Agency_Txn', payload.rowIndex, '');
        return cors(r);
      }

      // ── SUB_Txn CRUD ──
      case 'add_sub_txn': {
        const r = addMasterRow('SUB_Txn', payload.data);
        if (r.ok) appendAuditLog(_user, 'add_sub_txn', 'SUB_Txn', r.row || '',
          'Sub:' + (payload.data.SubCompany||'') + ' DR:' + (payload.data.DR||0));
        return cors(r);
      }
      case 'update_sub_txn': {
        const r = updateMasterRow('SUB_Txn', payload.rowIndex, payload.data);
        if (r.ok) appendAuditLog(_user, 'update_sub_txn', 'SUB_Txn', payload.rowIndex, '');
        return cors(r);
      }
      case 'delete_sub_txn': {
        const r = deleteMasterRow('SUB_Txn', payload.rowIndex);
        if (r.ok) appendAuditLog(_user, 'delete_sub_txn', 'SUB_Txn', payload.rowIndex, '');
        return cors(r);
      }

      // ── Notices ──
      case 'save_notices':
        return cors(replaceNotices(payload.rows));

      // ── Driver Info ──
      case 'update_driver_pin':
        return cors(updateDriverPin(payload.driverName, payload.pin));

      case 'update_driver_info':
        return cors(updateDriverInfo(payload.driverName, payload.data));

      default:
        return cors({ok: false, error: 'Unknown action: ' + action});
    }
  } catch (err) {
    return cors({ok: false, error: err.toString()});
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// GET Implementations
// ═══════════════════════════════════════════════════════════════════════════

function getReports(sheetName, driver) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return {ok: false, msg: sheetName + ' sheet not found'};

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return {ok: true, rows: []};

    const headers = data[0];
    const DATE_FIELDS = ['Date', 'Submitted', 'License_Expiry', 'Authority_Expiry', 'Rego_Date', 'HVIS_Date'];

    function formatCell(h, v) {
      if (DATE_FIELDS.indexOf(h) !== -1 && v instanceof Date && !isNaN(v)) {
        return formatDateForSheet(v);
      }
      return v;
    }

    let rows = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = formatCell(h, row[i]);
      });
      return obj;
    });

    if (driver) rows = rows.filter(r => r.Driver === driver);
    return {ok: true, rows};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

function getMaster(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return {ok: false, msg: sheetName + ' sheet not found'};

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return {ok: true, sheet: sheetName, rows: []};

    const headers = data[0];

    // 시트 헤더(공백 포함 가능)를 MASTER_HEADERS 정규 키(언더스코어)로 매핑
    // 예: "Manufacture Date" → "Manufacture_Date"
    const canonicalHeaders = MASTER_HEADERS[sheetName];
    const normToCanonical = {};
    if (canonicalHeaders) {
      canonicalHeaders.forEach(ch => {
        normToCanonical[normalizeKey(ch)] = ch;
      });
    }

    const rows = data.slice(1).map((row, rowIdx) => {
      const obj = {};
      headers.forEach((h, i) => {
        // 시트 헤더를 정규 키로 변환 (공백↔언더스코어 자동 처리)
        const canonKey = (h && normToCanonical[normalizeKey(h)]) || h;
        obj[canonKey] = row[i];
      });
      // 행 번호 저장 (1-based 시트 행): 헤더(1) + rowIdx(0-based) + 1
      obj._rowIndex = rowIdx + 2;
      return obj;
    });

    return {ok: true, sheet: sheetName, rows};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

function getAllMasters() {
  try {
    const sheets = ['M_Vehicles', 'M_Drivers', 'M_Clients', 'M_Guides', 'M_Hotels',
                    'M_PriceClient', 'M_PriceDriver', 'M_PriceSub',
                    'M_SvcOptions', 'M_HotelOptions', 'M_DistOptions', 'M_NightRates', 'M_Attractions',
                    'Sub_Rates', 'Ledger', 'MOT_Report'];
    const result = {};

    sheets.forEach(name => {
      const r = getMaster(name);
      result[name] = r.ok ? r.rows : [];
    });

    return {ok: true, data: result};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

function getSubRatesSheet() {
  try {
    const r = getMaster('Sub_Rates');
    return {ok: r.ok, rows: r.rows || []};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

function getPriceSubSheet() {
  try {
    const r = getMaster('M_PriceSub');
    return {ok: r.ok, rows: r.rows || []};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

function getLedgerSheet() {
  try {
    const r = getMaster('Ledger');
    return {ok: r.ok, rows: r.rows || []};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

function getWagesSheet(driver) {
  try {
    const r = getMaster('Wages');
    let rows = r.rows || [];
    if (driver) rows = rows.filter(row => row.Driver === driver);
    return {ok: r.ok, rows};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

function getNoticesSheet() {
  try {
    const r = getMaster('Notices');
    return {ok: r.ok, rows: r.rows || []};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

// Generic sheet rows getter (for Agency_Txn, SUB_Txn, etc.)
function getSheetRows(sheetName) {
  try {
    const r = getMaster(sheetName);
    return {ok: r.ok, rows: r.rows || []};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

function getActiveRegos() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const preSheet = ss.getSheetByName('Pre_Departure');
    const eosSheet = ss.getSheetByName('End_of_Shift');
    if (!preSheet) return {ok: true, regos: []};

    const preData = preSheet.getDataRange().getValues();
    if (preData.length < 2) return {ok: true, regos: []};
    const preHeaders = preData[0];

    // Get today's date in Sydney timezone (UTC+10)
    const now = new Date();
    const sydOffset = 10 * 60;
    const utc = now.getTime() + now.getTimezoneOffset() * 60000;
    const syd = new Date(utc + sydOffset * 60000);
    const yy = syd.getFullYear();
    const mm = String(syd.getMonth() + 1).padStart(2, '0');
    const dd = String(syd.getDate()).padStart(2, '0');
    const todayISO = yy + '-' + mm + '-' + dd;
    const todayDMY = dd + '/' + mm + '/' + yy;

    function isToday(val) {
      const s = String(val).trim().replace(/\s+.*/, '');
      return s === todayISO || s === todayDMY;
    }

    const preRows = preData.slice(1).map(row => {
      const obj = {};
      preHeaders.forEach((h, i) => obj[h] = row[i]);
      return obj;
    }).filter(r => isToday(r.Date));

    // Collect EoS data for today
    const eosSet = new Set();
    if (eosSheet && eosSheet.getLastRow() > 1) {
      const eosData = eosSheet.getDataRange().getValues();
      const eosH = eosData[0];
      eosData.slice(1).forEach(row => {
        const obj = {};
        eosH.forEach((h, i) => obj[h] = row[i]);
        if (isToday(obj.Date)) {
          eosSet.add(String(obj.Rego).trim() + '|' + String(obj.Date).trim());
        }
      });
    }

    // Find active regos (Pre_Departure without End_of_Shift)
    const active = [];
    const seen = new Set();
    preRows.forEach(r => {
      const regoKey = String(r.Rego).trim() + '|' + String(r.Date).trim();
      if (!eosSet.has(regoKey) && !seen.has(regoKey)) {
        seen.add(regoKey);
        const driverName = String(r.Driver || '').trim();
        active.push({
          driver: driverName || 'Unknown',
          rego: String(r.Rego).trim(),
          date: String(r.Date).trim(),
          startTime: String(r.Start_Time || '').trim()
        });
      }
    });

    return {ok: true, regos: active};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Report Write Operations
// ═══════════════════════════════════════════════════════════════════════════

function saveReport(sheetName, data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(sheetName);
    const headers = REPORT_HEADERS[sheetName];

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground('#1a56db').setFontColor('white').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const row = headers.map(h => data[h] !== undefined ? data[h] : '');
    sheet.appendRow(row);

    return {ok: true, sheet: sheetName, row: sheet.getLastRow()};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

function updateReport(sheetName, rowIndex, data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return {ok: false, msg: sheetName + ' sheet not found'};

    const headers = REPORT_HEADERS[sheetName];
    if (!headers) return {ok: false, msg: 'Unknown sheet: ' + sheetName};

    const ri = parseInt(rowIndex);
    if (!ri || ri < 2) return {ok: false, msg: 'Invalid rowIndex'};

    const row = headers.map(h => data[h] !== undefined ? data[h] : '');
    sheet.getRange(ri, 1, 1, row.length).setValues([row]);

    return {ok: true};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

function deleteReport(sheetName, rowIndex) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return {ok: false, msg: sheetName + ' sheet not found'};

    const ri = parseInt(rowIndex);
    if (!ri || ri < 2) return {ok: false, msg: 'Invalid rowIndex'};

    sheet.deleteRow(ri);
    return {ok: true};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Master Row Operations
// ═══════════════════════════════════════════════════════════════════════════

function addMasterRow(sheetName, data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, sheetName);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // 정확한 키 먼저, 없으면 정규화 키로 fallback
    const normMap = buildNormMap(data);
    const row = headers.map(h => {
      if (data[h] !== undefined) return data[h];
      const nk = normalizeKey(h);
      return normMap[nk] !== undefined ? normMap[nk] : '';
    });
    sheet.appendRow(row);

    return {ok: true, row: sheet.getLastRow()};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

function addPriceClientAgency(agencyName, rows) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'M_PriceClient');
    const headers = MASTER_HEADERS['M_PriceClient'];

    // Check for existing Agency+Course combinations
    const lastRow = sheet.getLastRow();
    const existing = new Set();
    if (lastRow > 1) {
      const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      data.forEach(r => {
        if (r[0] && r[1]) existing.add(r[0] + '||' + r[1]);
      });
    }

    const newRows = (rows || []).filter(r => !existing.has(r.Agency + '||' + r.Course));
    if (newRows.length === 0) return {ok: true, added: 0, msg: 'No duplicates found'};

    const values = newRows.map(r => headers.map(h => r[h] !== undefined ? r[h] : ''));
    sheet.getRange(lastRow + 1, 1, values.length, headers.length).setValues(values);

    return {ok: true, added: newRows.length};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

// 열 이름 정규화: 공백/하이픈 → 언더스코어, 소문자 변환
function normalizeKey(k) {
  return String(k).toLowerCase().replace(/[\s\-]+/g, '_');
}

// data 객체를 정규화 키로 조회하는 맵 생성
function buildNormMap(data) {
  const m = {};
  Object.keys(data).forEach(k => { m[normalizeKey(k)] = data[k]; });
  return m;
}

function updateMasterRow(sheetName, rowIndex, data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, sheetName);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    const ri = parseInt(rowIndex);
    if (!ri || ri < 2) return {ok: false, msg: 'Invalid rowIndex'};

    // 정확한 키 먼저, 없으면 정규화 키로 fallback (공백↔언더스코어 불일치 허용)
    const normMap = buildNormMap(data);
    const row = headers.map(h => {
      if (data[h] !== undefined) return data[h];
      const nk = normalizeKey(h);
      return normMap[nk] !== undefined ? normMap[nk] : '';
    });
    sheet.getRange(ri, 1, 1, row.length).setValues([row]);

    return {ok: true};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

function deleteMasterRow(sheetName, rowIndex) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return {ok: false, msg: 'Sheet not found'};

    const ri = parseInt(rowIndex);
    if (!ri || ri < 2) return {ok: false, msg: 'Invalid rowIndex'};

    sheet.deleteRow(ri);
    return {ok: true};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

function replaceMasterSheet(sheetName, rows) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, sheetName);
    const headers = MASTER_HEADERS[sheetName];

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

    if (rows && rows.length > 0) {
      const data = rows.map(obj => headers.map(h => obj[h] !== undefined ? obj[h] : ''));
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
    }

    return {ok: true, count: rows ? rows.length : 0};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

function initAllMasters() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const results = [];

    Object.keys(MASTER_HEADERS).forEach(name => {
      let sheet = ss.getSheetByName(name);
      if (sheet) {
        results.push({sheet: name, status: 'skipped'});
        return;
      }

      sheet = ss.insertSheet(name);
      const headers = MASTER_HEADERS[name];
      const color = TAB_COLORS[name] || '#1a56db';

      sheet.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground(color).setFontColor('white').setFontWeight('bold');
      sheet.setFrozenRows(1);
      sheet.setTabColor(color);

      results.push({sheet: name, status: 'created'});
    });

    return {ok: true, results};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Wages Operations (Fixed to 6 columns)
// ═══════════════════════════════════════════════════════════════════════════

function addWage(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'Wages');

    const rowId = Date.now().toString();
    const amount = parseFloat(data.Amount) || 0;

    const newRow = [
      rowId,
      data.Driver || '',
      data.WeekStart || '',
      data.Date || '',
      amount,
      data.PayMethod || 'Cash',
      data.Notes || ''
    ];

    sheet.appendRow(newRow);
    return {ok: true, row: sheet.getLastRow(), rowId};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

function updateWage(rowIndex, data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'Wages');

    const ri = parseInt(rowIndex);
    if (!ri || ri < 2) return {ok: false, msg: 'Invalid rowIndex'};

    const lastRow = sheet.getLastRow();
    if (ri > lastRow) return {ok: false, msg: 'Row does not exist'};

    const amount = parseFloat(data.Amount) || 0;

    sheet.getRange(ri, 1, 1, 7).setValues([[
      data.RowID || Date.now().toString(),
      data.Driver || '',
      data.WeekStart || '',
      data.Date || '',
      amount,
      data.PayMethod || 'Cash',
      data.Notes || ''
    ]]);

    return {ok: true};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

function deleteWage(rowIndex) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Wages');
    if (!sheet) return {ok: false, msg: 'Wages sheet not found'};

    const ri = parseInt(rowIndex);
    if (!ri || ri < 2) return {ok: false, msg: 'Invalid rowIndex'};

    sheet.deleteRow(ri);
    return {ok: true};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

function replaceWages(rows) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'Wages');

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

    if (rows && rows.length > 0) {
      const newData = rows.map(r => [
        r.RowID || Date.now().toString(),
        r.Driver || '',
        r.WeekStart || '',
        r.Date || '',
        parseFloat(r.Amount) || 0,
        r.PayMethod || 'Cash',
        r.Notes || ''
      ]);
      sheet.getRange(2, 1, newData.length, 7).setValues(newData);
    }

    return {ok: true, count: rows ? rows.length : 0};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Driver Operations
// ═══════════════════════════════════════════════════════════════════════════

function updateDriverPin(driverName, pin) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('M_Drivers');
    if (!sheet) return {ok: false, msg: 'M_Drivers sheet not found'};

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const nameENIdx = headers.indexOf('Name_EN');
    const nameKRIdx = headers.indexOf('Name_KR');
    const pinIdx = headers.indexOf('PIN');

    if (pinIdx === -1) return {ok: false, msg: 'PIN column not found'};

    for (let r = 1; r < data.length; r++) {
      if (data[r][nameENIdx] === driverName || data[r][nameKRIdx] === driverName) {
        sheet.getRange(r + 1, pinIdx + 1).setValue(pin);
        return {ok: true};
      }
    }

    return {ok: false, msg: 'Driver not found: ' + driverName};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

function updateDriverInfo(driverName, data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('M_Drivers');
    if (!sheet) return {ok: false, msg: 'M_Drivers sheet not found'};

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

    return {ok: false, msg: 'Driver not found: ' + driverName};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Notices Operations
// ═══════════════════════════════════════════════════════════════════════════

function replaceNotices(rows) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'Notices');

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

    if (rows && rows.length > 0) {
      const newData = rows.map(r => [
        r.id || r.ID || String(Date.now()),
        r.title || r.Title || '',
        r.content || r.Content || '',
        r.type || r.Type || 'info',
        r.date || r.Date || '',
        r.active === false || r.Active === 'false' ? 'false' : 'true'
      ]);
      sheet.getRange(2, 1, newData.length, 6).setValues(newData);
    }

    return {ok: true, count: rows ? rows.length : 0};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

function saveCorrectionRequest(payload) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'Notices');

    const type    = payload.reportType || '';   // 'Pre_Departure' | 'End_of_Shift'
    const driver  = payload.driver     || '';
    const date    = payload.date       || '';
    const rego    = payload.rego       || '';
    const desc    = payload.description || '';

    const typeLabel = type === 'Pre_Departure' ? 'Pre Departure'
                    : type === 'End_of_Shift'  ? 'End of Shift'
                    : type;

    const id      = 'CR-' + Date.now();
    const title   = '[수정요청] ' + typeLabel + ' · ' + driver + ' · ' + date + ' · ' + rego;
    const content = desc;
    const rowDate = Utilities.formatDate(new Date(), 'Australia/Sydney', 'dd/MM/yyyy');

    sheet.appendRow([id, title, content, 'correction', rowDate, 'true']);
    return {ok: true, id: id};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Helper Functions
// ═══════════════════════════════════════════════════════════════════════════

function fixReportHeaders() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const results = [];

    Object.keys(REPORT_HEADERS).forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        results.push(sheetName + ': Sheet not found');
        return;
      }

      const headers = REPORT_HEADERS[sheetName];
      const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const currentStr = currentHeaders.join(',');
      const targetStr = headers.join(',');

      if (currentStr === targetStr) {
        results.push(sheetName + ': Already matches ✓');
        return;
      }

      const newLen = headers.length;
      const oldLen = sheet.getLastColumn();

      sheet.getRange(1, 1, 1, newLen).setValues([headers]);
      if (oldLen > newLen) {
        sheet.getRange(1, newLen + 1, 1, oldLen - newLen).clearContent();
      }

      sheet.getRange(1, 1, 1, newLen)
        .setBackground('#1a56db').setFontColor('white').setFontWeight('bold');
      sheet.setFrozenRows(1);

      results.push(sheetName + ': Headers updated (' + oldLen + '→' + newLen + ' columns)');
    });

    Logger.log(results.join('\n'));
    return {ok: true, results};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

function fixPhoneNumbers() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const targets = [
      {sheet: 'M_Guides', col: 'Mobile'},
      {sheet: 'M_Drivers', col: 'Mobile_1'},
      {sheet: 'M_Hotels', col: 'Phone'}
    ];

    let totalFixed = 0;

    targets.forEach(({sheet: sheetName, col: colName}) => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        Logger.log(sheetName + ' sheet not found');
        return;
      }

      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return;

      const headers = data[0];
      const colIdx = headers.indexOf(colName);
      if (colIdx === -1) {
        Logger.log(sheetName + '.' + colName + ' column not found');
        return;
      }

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

      Logger.log(sheetName + '.' + colName + ': ' + fixed + ' numbers fixed');
      totalFixed += fixed;
    });

    Logger.log('Total: ' + totalFixed + ' phone numbers fixed');
    return {ok: true, totalFixed};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Max KM Per Rego — Service Schedule Helper
// Scans Pre_Departure (Start_KM), Daily_Report (KM_Start, KM_End),
// and End_of_Shift (End_KM) to return the highest KM recorded per rego.
// ═══════════════════════════════════════════════════════════════════════════
function getMaxKMPerRego() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const kmMap = {};  // { rego: maxKM }

    function scanSheet(sheetName, kmFields) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) return;
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const regoIdx = headers.indexOf('Rego');
      if (regoIdx < 0) return;
      const colIdxs = kmFields.map(f => headers.indexOf(f)).filter(i => i >= 0);
      data.slice(1).forEach(row => {
        const rego = String(row[regoIdx] || '').trim();
        if (!rego) return;
        colIdxs.forEach(ci => {
          const v = parseFloat(row[ci]);
          if (!isNaN(v) && v > 0) {
            if (!kmMap[rego] || v > kmMap[rego]) kmMap[rego] = v;
          }
        });
      });
    }

    scanSheet('Pre_Departure', ['Start_KM']);
    scanSheet('Daily_Report',  ['KM_Start', 'KM_End']);
    scanSheet('End_of_Shift',  ['End_KM']);

    return { ok: true, kmMap };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// AUDIT TRAIL
// ═══════════════════════════════════════════════════════════════════════════

/**
 * 변경 이력을 Audit_Log 시트에 한 행 추가
 * @param {string} user      - 관리자 계정 (dc_admin_session)
 * @param {string} action    - 작업 종류 (update_report, delete_master 등)
 * @param {string} sheet     - 대상 시트명
 * @param {number|string} rowIndex - 대상 행 번호 (없으면 '')
 * @param {string} summary   - 변경 내용 요약 (JSON string 또는 free text)
 */
function appendAuditLog(user, action, sheet, rowIndex, summary) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let logSheet = ss.getSheetByName('Audit_Log');
    if (!logSheet) {
      logSheet = ss.insertSheet('Audit_Log');
      const headers = MASTER_HEADERS['Audit_Log'];
      logSheet.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground('#374151').setFontColor('white').setFontWeight('bold');
      logSheet.setFrozenRows(1);
      logSheet.setColumnWidth(1, 160);  // Timestamp
      logSheet.setColumnWidth(6, 400);  // Summary
    }

    // 시드니 현지 시각 문자열
    const now = new Date();
    const sydFmt = Utilities.formatDate(now, 'Australia/Sydney', 'dd/MM/yyyy HH:mm:ss');

    logSheet.appendRow([
      sydFmt,
      user || 'unknown',
      action || '',
      sheet || '',
      rowIndex || '',
      typeof summary === 'string' ? summary.slice(0, 500) : JSON.stringify(summary).slice(0, 500)
    ]);
  } catch (e) {
    // 감사 로그 실패는 무시 (메인 작업 방해하지 않음)
    console.warn('appendAuditLog error:', e.toString());
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// INVOICES — CRUD (Invoices 시트)
// ═══════════════════════════════════════════════════════════════════════════

/**
 * 인보이스 저장 (신규 또는 기존 덮어쓰기)
 * data: { InvNumber, Agency, PeriodFrom, PeriodTo, GrandTotal, GST, ExGST,
 *         Status, IssuedDate, Items (JSON), ManualItems (JSON), Notes, CreatedBy }
 */
function saveInvoice(data) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'Invoices');
    const headers = MASTER_HEADERS['Invoices'];
    const invNum  = data.InvNumber || data.invNumber || '';
    if (!invNum) return { ok: false, error: 'InvNumber required' };

    // 기존 행 찾기 (InvNumber 기준)
    const allData = sheet.getDataRange().getValues();
    let existingRow = -1;
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][0]).trim() === invNum) { existingRow = i + 1; break; }
    }

    const now = new Date();
    const sydNow = Utilities.formatDate(now, 'Australia/Sydney', 'dd/MM/yyyy HH:mm:ss');
    if (!data.IssuedDate) data.IssuedDate = sydNow;

    const rowArr = headers.map(h => {
      if (h === 'InvNumber') return invNum;
      return data[h] !== undefined ? data[h] : '';
    });

    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, headers.length).setValues([rowArr]);
    } else {
      sheet.appendRow(rowArr);
    }

    return { ok: true, invNumber: invNum, updated: existingRow > 0 };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

/**
 * 모든 인보이스 조회
 */
function getInvoices() {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Invoices');
    if (!sheet) return { ok: true, rows: [] };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { ok: true, rows: [] };

    const headers = data[0];
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const obj = {};
      headers.forEach((h, ci) => { obj[h] = data[i][ci]; });
      obj._rowIndex = i + 1;
      rows.push(obj);
    }
    return { ok: true, rows };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

/**
 * 인보이스 상태 변경
 * invNumber: 인보이스 번호
 * status: 'issued' | 'emailed' | 'paid' | 'cancelled'
 * field: 상태 변경 시 날짜 기록 필드 ('EmailSentDate' | 'PaidDate')
 */
function updateInvoiceStatus(invNumber, status, field) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Invoices');
    if (!sheet) return { ok: false, error: 'Invoices sheet not found' };

    const headers = MASTER_HEADERS['Invoices'];
    const data = sheet.getDataRange().getValues();
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === invNumber) { targetRow = i + 1; break; }
    }
    if (targetRow < 0) return { ok: false, error: 'Invoice not found: ' + invNumber };

    const now = Utilities.formatDate(new Date(), 'Australia/Sydney', 'dd/MM/yyyy HH:mm:ss');

    // Status 열 업데이트
    const statusCol = headers.indexOf('Status') + 1;
    if (statusCol > 0) sheet.getRange(targetRow, statusCol).setValue(status);

    // 날짜 필드 업데이트
    if (field) {
      const fieldCol = headers.indexOf(field) + 1;
      if (fieldCol > 0) sheet.getRange(targetRow, fieldCol).setValue(now);
    }

    return { ok: true, invNumber, status, updatedAt: now };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

/**
 * 인보이스 삭제
 */
function deleteInvoice(invNumber) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Invoices');
    if (!sheet) return { ok: false, error: 'Invoices sheet not found' };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === invNumber) {
        sheet.deleteRow(i + 1);
        return { ok: true, invNumber };
      }
    }
    return { ok: false, error: 'Invoice not found: ' + invNumber };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// INVOICE EMAIL (GAS MailApp)
// ═══════════════════════════════════════════════════════════════════════════

/**
 * 인보이스 이메일 발송
 * payload: { to, subject, body, cc, _user }
 */
function sendInvoiceEmail(payload) {
  try {
    const to      = (payload.to || '').trim();
    const subject = (payload.subject || '').trim();
    const body    = (payload.body || '').trim();
    const cc      = (payload.cc || '').trim();
    const name    = payload.senderName || 'Dong Choi Pty Ltd';

    if (!to)      return { ok: false, error: '수신자 이메일이 없습니다 (to is empty)' };
    if (!subject) return { ok: false, error: '제목이 없습니다 (subject is empty)' };

    const mailOptions = {
      to:      to,
      subject: subject,
      body:    body,
      name:    name
    };
    if (cc) mailOptions.cc = cc;

    MailApp.sendEmail(mailOptions);

    // 감사 로그
    appendAuditLog(payload._user, 'send_invoice_email', '—', '—',
      `인보이스 이메일 발송 → ${to} | ${subject}`);

    return { ok: true, to: to };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// 가이드 전화번호 일괄 업데이트
// guides: [{ Guide_Name: '...', Mobile: '...' }, ...]
// 기존 M_Guides의 Guide_Name과 매칭하여 Mobile 컬럼을 업데이트
// ═══════════════════════════════════════════════════════════════════════════
function bulkUpdateGuidePhones(guides) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'M_Guides');
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { ok: false, msg: 'M_Guides 시트에 데이터 없음' };

    const headers = data[0];
    const nameCol = headers.indexOf('Guide_Name') !== -1 ? headers.indexOf('Guide_Name')
                  : headers.indexOf('Guide Name') !== -1 ? headers.indexOf('Guide Name')
                  : headers.indexOf('Name') !== -1 ? headers.indexOf('Name') : -1;
    const mobileCol = headers.indexOf('Mobile') !== -1 ? headers.indexOf('Mobile')
                    : headers.indexOf('Phone') !== -1 ? headers.indexOf('Phone') : -1;

    if (nameCol === -1 || mobileCol === -1) return { ok: false, msg: 'Guide_Name 또는 Mobile 컬럼 없음' };

    const guideMap = {};
    guides.forEach(g => { if (g.Guide_Name && g.Mobile) guideMap[g.Guide_Name.trim()] = g.Mobile; });

    let updated = 0;
    for (let i = 1; i < data.length; i++) {
      const name = String(data[i][nameCol] || '').trim();
      if (name && guideMap[name]) {
        const currentMobile = String(data[i][mobileCol] || '').trim();
        if (!currentMobile) {  // 빈 셀만 업데이트
          sheet.getRange(i + 1, mobileCol + 1).setValue(guideMap[name]);
          updated++;
        }
      }
    }

    return { ok: true, updated, total: guides.length };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}
