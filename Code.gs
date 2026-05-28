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
const DRIVE_ROOT_FOLDER = 'DongChoi_DriverDocs'; // Google Drive 루트 폴더명

// ── Report Sheet Headers ──
const REPORT_HEADERS = {
  'Daily_Report':   ['Submitted','Driver','Date','Rego','Seats','Agency','Attraction','Pickup','Dropoff',
                     'KM_Start','KM_End','Time_Start','Time_End','Guide','Tour_Code',
                     'SVC_Label','SVC_Charge','Hotel_Surcharge','Dist_Surcharge',
                     'OT','Trailer','Trailer_Number','Total_TA','DR_Cost','Toll','Toll_Personal',
                     'Fuel','Fuel_Personal','Early','Night_Type','Night_DR','Night_Owner',
                     'Wash','Meal','Tip','Etc','Etc_Desc','Remarks',
                     'SubPaid_Owner','SubPaid_OwnerAt','SubPaid_OwnerBy',
                     'SubPaid_Driver','SubPaid_DriverAt'],
  'Pre_Departure':  ['Submitted','Driver','Date','Rego','Seats','Start_KM','Fuel','Start_Time',
                     'Check_Results','Remarks','Signature','Trailer_Number'],
  'End_of_Shift':   ['Submitted','Driver','Date','Rego','Start_KM','End_KM','End_Time','Fuel_End','Damage','Check_Results','Daily_Reports','Remarks','Signature'],
  'MOT_Report':     ['Submitted','Driver','Date','Time','Rego','Location','Officer','Type',
                     'Result','NoticeNum','Fine','Notes','FailedItems','Checks']
};

// ── Master Sheet Headers ──
const MASTER_HEADERS = {
  'M_Vehicles': ['Rego','Make','Model','Manufacture_Date','Capacity','Owner','Rego_Date','HVIS_Date',
                 'Current_KM','Last_Service_KM','Service_Interval','VIN','Engine_Number',
                 'Accreditation','Current_Status','Transmission','Active',
                 'Photo_Front','Photo_Back','Photo_Left','Photo_Right'],
  'M_Drivers':  ['Name_EN','Name_KR','Initials','DriverID','Mobile_1','NEXT_OF_KIN','Mobile_2','License_Class',
                 'License_No','License_Expiry','Authority_No','Authority_Expiry','WWC_No','WWC_Expiry',
                 'Address','Suburb','Bank_Name','BSB','Account_Number','PIN','Owner','Active'],
  'M_Clients':  ['Name','ClientID','ABN','Mobile','Email','Email_CC','Address','Bank_Name','BSB','Account_Number','Payment_Terms'],
  'M_Guides':   ['GuideID','Guide_Name','Mobile','Agency','Email','Remarks'],
  'M_Hotels':   ['Hotel_Name','Phone','Address','Surcharge_Area','Short_Name'],
  'M_Trailers': ['Trailer_Number','Owner','Capacity','Rego_Date','ESafety_Date','Notes','Active'],
  'M_PriceClient': ['Agency','Course','max_hours','seats_21_rate','seats_21_ot',
                    'seats_25_rate','seats_25_ot','seats_40_rate','seats_40_ot',
                    'seats_50_rate','seats_50_ot'],
  'M_PriceDriver': ['Course','max_hours','seats_21_base','seats_21_ot',
                    'seats_25_base','seats_25_ot','seats_40_base','seats_40_ot',
                    'seats_50_base','seats_50_ot'],
  'M_PriceSub': ['SubCo','Course','max_hours','seats_21_rate','seats_21_ot',
                 'seats_25_rate','seats_25_ot','seats_40_rate','seats_40_ot',
                 'seats_50_rate','seats_50_ot'],
  'Sub_Rates':  ['SubCo','Tour','seats_21','seats_25','seats_40','seats_50'],
  'Ledger':     ['RowID','Date','Rego','Tour','TA','SubTotal','MyDr','Extra','OT','Trailer','Hotel','Note'],
  'Wages':      ['RowID','Driver','WeekStart','Date','Amount','PayMethod','Notes'],
  'Notices':    ['ID','Title','Content','Type','Date','Active'],
  'Audit_Log':  ['Timestamp','User','Action','Sheet','RowIndex','Summary'],
  'Invoices':   ['InvNumber','Agency','PeriodFrom','PeriodTo','GrandTotal','GST','ExGST',
                 'Status','IssuedDate','EmailSentDate','PaidDate','Items','ManualItems','Notes','CreatedBy'],
  // ── 드라이버 근무/휴무 로스터 ──
  'Driver_Roster': ['Driver','Date','Status','Updated_At','Source'],
  // ── 거래처 잔액 관리 ──
  'Agency_Txn': ['RowID','Agency','Date','InvoiceID','TourCode','Guide','Type','DR','CR','Remark','StartDate','FinishDate','DueDate'],
  'SUB_Txn':    ['RowID','SubCompany','Category','Date','InvoiceNo','TourCode','Description','DR','CR','Remark'],
  // ── 서비스 요금 옵션 (차량 좌석별) ──
  'M_SvcOptions': ['VehicleSize','Label','Amount'],
  // ── 호텔 서차지 옵션 ──
  'M_HotelOptions': ['VehicleSize','Label','Amount'],
  // ── 거리 서차지 옵션 ──
  'M_DistOptions': ['VehicleSize','Label','Amount'],
  // ── 야간투어 요금 ──
  'M_NightRates': ['NightType','VehicleCategory','TA','DR','Owner'],
  // ── 관광지 POI 정보 ──
  'M_Attractions': ['Attraction','Emoji','POI_Icon','POI_Name','POI_Detail','POI_MapURL','Info'],
  // ── 결함 리포트 ──
  'Defect_Reports': ['ID','Rego','Category','Location','Description','Severity','KM','Driver','Status','SubmittedAt','AdminNote'],
  // ── 차량 데미지 마커 ──
  'Bus_Damage': ['Rego','Markers','UpdatedAt','UpdatedBy'],
  // ── HVIS 부킹 관리 ──
  'HVIS_Bookings': ['ID','Rego','InspDate','InspTime','Location','CustomerNo','BookingNo','VehicleType','OwnerName','BookingDate','Status'],
  // ── 정비 기록 ──
  'Maint_Records': ['ID','Rego','Date','KM','Type','Description','Workshop','Cost','NextServiceKM'],
  // ── 인보이스 서차지 오버라이드 ──
  'Invoice_Overrides': ['RowKey','Value'],
  // ── 회사 정보 (single-row config) ──
  'Company_Profile': ['Key','Value'],
  // ── 인보이스 공제 항목 ──
  'Invoice_Deductions': ['ID','Agency','Period','Type','Amount','Note'],
  // ── 인보이스 수동 항목 ──
  'Invoice_Manual_Items': ['ID','Agency','Period','Date','Rego','Tour','Seats','TourCode','Note','Amount','OT','Hotel','Dist','Trailer','Toll','Start','End','Driver','Guide','Pickup','Dropoff'],
  // ── 인증 토큰 (세션 관리) ──
  'Active_Tokens': ['Token','User','Role','IssuedAt','ExpiresAt','LastUsed','UserAgent'],
  // ── 로그인 실패 추적 (rate limiting) ──
  'Auth_Failures': ['Name','FailCount','FirstFail','LastFail','LockedUntil'],
  // ── 운행 일정 (Schedule, 중기 자동화 핵심 데이터) ──
  // Status: scheduled / in_progress / completed / invoiced / paid / cancelled
  // TourPlan: JSON string [{date, course, ot, hotel, pickup, dropoff, note}]
  // BillingEntity: 인보이스 발행사 ('DC' = 자사 발행, 또는 'EG TRAVEL PTY LTD' 같은 파트너사명)
  'Schedule':   ['TourID','Agency','TourCode','StartDate','EndDate','Pax','Seats','Trailer',
                 'Guide','GuidePhone','Driver','Rego','FlightIn','FlightOut','Hotel',
                 'TourPlan','Notes','Status','InvoiceID','Quote','BillingEntity','CreatedAt','UpdatedAt'],
  // 외주 지급 오버라이드 — BillingEntity 자동 판단 결과를 수동으로 뒤집을 때 사용
  // Action: 'INCLUDE' (강제 포함) | 'EXCLUDE' (강제 제외)
  'PayoutOverrides': ['TourCode','SubCompany','Action','UpdatedAt','UpdatedBy'],
  // EG TRAVEL 자동 리포트 발송 이력 — 중복 발송 방지용 (특히 종료된 투어 섹션)
  // ReportType: 'daily' | 'weekly' | 'manual'
  // TourCodes: 이번 발송에 포함된 종료 투어코드 목록 (콤마 구분)
  'EG_Report_Log': ['SentAt','ReportType','PeriodFrom','PeriodTo','Recipients','TourCodes','Subject','Status','Notes']
};

// ═══════════════════════════════════════════════════════════════
// BillingEntity가 DC(자사)인지 판정 — 다양한 표기 모두 허용
// ═══════════════════════════════════════════════════════════════
// 잡히는 표기 (모두 true):
//   '', null, undefined (빈값 = 기본 자사)
//   'DC', 'dc', 'Dc', 'D.C.', 'D.C', 'D C' (점/공백 변이)
//   'Dong Choi', 'DONG CHOI PTY LTD', 'dongchoi', 'Dong  Choi  Pty  Ltd'
//   '동초이', '동최' (한글 표기 — 향후 확장 대비)
// 잡히지 않는 표기 (false):
//   'EG TRAVEL PTY LTD', 'TOUR HOJURO PTY LTD' 등 다른 회사명
function isBillingEntityDC_(be){
  if (be === null || be === undefined) return true;
  var s = String(be).replace(/^\s+|\s+$/g,'');  // trim
  if (!s) return true;
  var norm = s.replace(/[.\s\-_·]+/g,'').toUpperCase();
  if (norm === 'DC') return true;
  if (norm.indexOf('DONGCHOI') >= 0) return true;
  if (s.indexOf('동초이') >= 0 || s.indexOf('동최') >= 0) return true;
  return false;
}

// ── Tab Colors ──
const TAB_COLORS = {
  'M_Vehicles':'#d97706','M_Drivers':'#1a56db','M_Clients':'#7e3af2',
  'M_Guides':'#0e9f6e','M_Hotels':'#e02424','M_PriceClient':'#0694a2',
  'M_PriceDriver':'#057a55','M_PriceSub':'#7c3aed','Sub_Rates':'#b45309',
  'Ledger':'#1e40af','Wages':'#065f46','MOT_Report':'#be185d','Notices':'#0369a1',
  'Agency_Txn':'#0891b2','SUB_Txn':'#a21caf',
  'Invoices':'#6d28d9',
  'M_SvcOptions':'#6366f1','M_HotelOptions':'#ec4899','M_DistOptions':'#f59e0b',
  'M_NightRates':'#8b5cf6','M_Attractions':'#14b8a6',
  'Defect_Reports':'#dc2626','Bus_Damage':'#ea580c',
  'Maint_Records':'#059669','Invoice_Overrides':'#7c3aed','Company_Profile':'#0284c7',
  'Invoice_Deductions':'#db2777','Invoice_Manual_Items':'#9333ea',
  'Active_Tokens':'#374151',
  'Auth_Failures':'#991b1b'
};

// ═══════════════════════════════════════════════════════════════════════════
// Utility Functions
// ═══════════════════════════════════════════════════════════════════════════

function cors(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════════════════════════
// AUTHENTICATION MODULE (토큰 기반 인증)
// ═══════════════════════════════════════════════════════════════════════════
//
// 흐름:
//   1) action=login: 이름 + PIN → 서버에서 M_Drivers 조회 → 검증 → 토큰 발급
//   2) 이후 모든 요청: token 파라미터 필수 (login, ping, get_company_profile_public 제외)
//   3) 관리자 전용 action은 role='admin' 토큰만 허용
//   4) 만료된 토큰은 자동 삭제
//
// 유효기간:
//   - 드라이버: 7일
//   - 관리자:  24시간
//
// M_Drivers의 PIN은 절대로 클라이언트에 응답으로 나가지 않음 (strip_pin_from_master)
// ═══════════════════════════════════════════════════════════════════════════

// 관리자 계정 이름 (M_Drivers의 Name_KR 또는 Name_EN와 일치)
const ADMIN_NAMES = ['Branden Choi', 'Branden', '최동철', 'Dong Cheol Choi'];

const TOKEN_TTL_DRIVER_MS = 7 * 24 * 60 * 60 * 1000;   // 7일
const TOKEN_TTL_ADMIN_MS  = 1 * 24 * 60 * 60 * 1000;   // 24시간

// ── 보안 상수 ──
// PIN 해시 식별 prefix (이 prefix가 있으면 해시된 값으로 인식)
const PIN_HASH_PREFIX = 'h1$';
// PIN 해시 salt (시스템 고유값 — 변경 시 모든 PIN 재설정 필요)
// PIN 해시 salt (시스템 고유값 — 변경 시 모든 PIN 재설정 필요)
// ★ 보안: Script Properties에서 조회 (코드에 평문 노출 방지)
//   설정 방법: Apps Script 에디터 → 프로젝트 설정 (⚙️) → 스크립트 속성 → 추가
//     속성: PIN_HASH_SECRET
//     값: DC_FLEET_2026_K7p9Qx2L  (또는 더 강력한 새 secret)
//   설정 안 하면 폴백 값 사용 (기존 PIN 호환 유지)
const PIN_HASH_SECRET_FALLBACK = 'DC_FLEET_2026_K7p9Qx2L';
function _getPinSecret() {
  try {
    const v = PropertiesService.getScriptProperties().getProperty('PIN_HASH_SECRET');
    if (v && v.length > 0) return v;
  } catch(e) {}
  return PIN_HASH_SECRET_FALLBACK;
}
// Rate limiting: 5회 실패 → 15분 락
const AUTH_MAX_FAILS = 5;
const AUTH_LOCK_MS = 15 * 60 * 1000;
// 실패 카운트 리셋 윈도우: 첫 실패 후 30분 내 시도만 누적
const AUTH_FAIL_WINDOW_MS = 30 * 60 * 1000;

// 인증 없이 호출 가능한 액션 (로그인 및 공개 메타데이터)
const PUBLIC_ACTIONS = ['ping', 'login', 'logout', 'get_login_names'];

// 관리자 전용 액션 (드라이버 토큰 거부)
const ADMIN_ONLY_ACTIONS = [
  'update_report', 'delete_report',
  'add_master', 'update_master', 'delete_master', 'replace_master',
  'bulk_update_guide_phones', 'init_masters',
  'send_invoice_email', 'save_invoice', 'update_invoice_status', 'delete_invoice',
  'replace_sub_rates', 'replace_price_sub',
  'add_ledger', 'update_ledger', 'delete_ledger', 'replace_ledger',
  'add_wage', 'update_wage', 'delete_wage', 'replace_wages',
  'add_agency_txn', 'update_agency_txn', 'delete_agency_txn',
  'add_sub_txn', 'update_sub_txn', 'delete_sub_txn',
  'save_notices',
  // update_driver_info는 본인 정보 수정에 한해 드라이버도 허용 (doPost에서 driverName 강제)
  'update_defect_status',
  'review_leave_request', 'update_roster_cell',
  'save_hvis_booking', 'delete_hvis_booking',
  'save_maint_record', 'delete_maint_record',
  'save_invoice_override', 'delete_invoice_override', 'bulk_save_invoice_overrides',
  'save_company_profile',
  // ── 운행 일정 관리 (Schedule) ──
  'save_schedule', 'delete_schedule', 'update_schedule_status',
  // ── EG TRAVEL 자동 리포트 발송 ──
  'send_eg_daily_report', 'send_eg_weekly_report', 'setup_eg_report_triggers',
  // 관리자가 주로 쓰지만 드라이버도 가끔 필요할 수 있는 조회는 제외:
  // get_invoices, get_agency_txn, get_sub_txn 등은 일단 드라이버도 허용
  // 추후 엄격하게 할 수 있음
];

// 관리자 전용 GET 액션
const ADMIN_ONLY_GET_ACTIONS = [
  'get_agency_txn', 'get_sub_txn', 'get_agency_balances',
  'get_invoices', 'get_all_leave_requests',
  'get_ledger',
  // get_defect_reports, get_roster: 드라이버는 본인 것만 조회 (case 핸들러에서 effectiveDriver 강제)
  'get_admin_bundle', 'get_audit_log',
  // ── 운행 일정 ──
  'get_schedule', 'get_schedule_stats',
  // ── EG 리포트 미리보기 ──
  'preview_eg_report'
];

function _getAuthSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ensureSheet(ss, 'Active_Tokens');
}

function _getAuthFailSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ensureSheet(ss, 'Auth_Failures');
}

// ── PIN 해시 (SHA-256, salt=secret + name) ─────────────────────────────────
// 결과 형식: 'h1$' + base64url(sha256(secret + ':' + name + ':' + pin))
function _hashPin(pin, name) {
  const input = _getPinSecret() + ':' + String(name || '').trim() + ':' + String(pin || '');
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8);
  return PIN_HASH_PREFIX + Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

// 저장된 PIN(평문 또는 해시)과 입력 PIN 비교
// - 저장값이 'h1$'로 시작 → 해시 비교
// - 그 외 → 평문 비교 (마이그레이션 호환)
function _verifyPin(storedPin, inputPin, name) {
  const stored = String(storedPin || '');
  const input = String(inputPin || '');
  if (!stored || !input) return false;
  if (stored.indexOf(PIN_HASH_PREFIX) === 0) {
    return stored === _hashPin(input, name);
  }
  // 평문 비교 (마이그레이션 전 호환)
  return stored === input;
}

// ── Rate Limiting ─────────────────────────────────────────────────────────
// 반환: {locked: bool, remainingMs?: number, failCount?: number}
function _checkAuthLock(name) {
  try {
    const sheet = _getAuthFailSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return {locked: false};
    const now = new Date().getTime();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === name) {
        const lockedUntilStr = String(data[i][4] || '');
        if (lockedUntilStr) {
          const lockedUntil = new Date(lockedUntilStr).getTime();
          if (!isNaN(lockedUntil) && lockedUntil > now) {
            return {locked: true, remainingMs: lockedUntil - now, failCount: Number(data[i][1] || 0)};
          }
        }
        return {locked: false, failCount: Number(data[i][1] || 0), rowIndex: i + 1};
      }
    }
    return {locked: false};
  } catch (err) {
    return {locked: false};  // fail-open: 시트 오류 시 정상 진행
  }
}

// 로그인 실패 기록
function _recordAuthFail(name) {
  try {
    const sheet = _getAuthFailSheet();
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    const nowMs = now.getTime();
    let foundRow = -1;
    let firstFail = now.toISOString();
    let failCount = 0;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === name) {
        foundRow = i + 1;
        const ff = new Date(String(data[i][2] || ''));
        // 윈도우 만료 시 카운트 리셋
        if (!isNaN(ff.getTime()) && (nowMs - ff.getTime()) < AUTH_FAIL_WINDOW_MS) {
          failCount = Number(data[i][1] || 0);
          firstFail = String(data[i][2] || now.toISOString());
        }
        break;
      }
    }
    failCount += 1;
    const lockedUntil = (failCount >= AUTH_MAX_FAILS) ? new Date(nowMs + AUTH_LOCK_MS).toISOString() : '';
    const rowData = [name, failCount, firstFail, now.toISOString(), lockedUntil];
    if (foundRow > 0) {
      sheet.getRange(foundRow, 1, 1, 5).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    return {failCount: failCount, locked: !!lockedUntil};
  } catch (err) {
    return {failCount: 0, locked: false};
  }
}

// 로그인 성공 시 실패 기록 삭제
function _clearAuthFails(name) {
  try {
    const sheet = _getAuthFailSheet();
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === name) {
        sheet.deleteRow(i + 1);
      }
    }
  } catch (err) {
    // ignore
  }
}

function _generateToken() {
  // 256-bit 랜덤 문자열 (base64 url-safe)
  const bytes = new Array(32);
  for (let i = 0; i < 32; i++) bytes[i] = Math.floor(Math.random() * 256);
  // Apps Script 에서 byte array → base64
  const blob = Utilities.newBlob(bytes);
  return Utilities.base64EncodeWebSafe(blob.getBytes()).replace(/=+$/, '');
}

function _loginAction(payload) {
  try {
    const nameInput = String(payload.name || '').trim();
    const pinInput  = String(payload.pin || '').trim();
    const userAgent = String(payload.ua || '').slice(0, 100);
    if (!nameInput || !pinInput) {
      return {ok: false, error: 'name and pin required'};
    }

    // ── Rate limit 체크 ──
    const lockState = _checkAuthLock(nameInput);
    if (lockState.locked) {
      const mins = Math.ceil(lockState.remainingMs / 60000);
      return {ok: false, error: 'locked', reason: 'too_many_attempts',
              lockMinutes: mins,
              message: '로그인 시도가 너무 많습니다. ' + mins + '분 후 다시 시도하세요.'};
    }

    // M_Drivers에서 사용자 찾기 (Name_KR 또는 Name_EN 매칭)
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('M_Drivers');
    if (!sheet) return {ok: false, error: 'M_Drivers not found'};
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return {ok: false, error: 'no drivers'};
    const headers = data[0].map(String);
    const nameKrIdx = headers.indexOf('Name_KR');
    const nameEnIdx = headers.indexOf('Name_EN');
    const pinIdx    = headers.indexOf('PIN');
    const activeIdx = headers.indexOf('Active');
    if (pinIdx === -1) return {ok: false, error: 'PIN column missing'};

    let matched = null;
    let matchedRow = -1;
    let storedPinPlaintext = false;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const nameKr = String(row[nameKrIdx] || '').trim();
      const nameEn = String(row[nameEnIdx] || '').trim();
      const active = activeIdx >= 0 ? String(row[activeIdx] || '').toUpperCase() : 'Y';
      if (active && active !== 'Y' && active !== '') continue;
      if (nameKr === nameInput || nameEn === nameInput) {
        const storedPin = String(row[pinIdx] || '').trim();
        // 저장된 이름(시트의 KR)으로 해시 검증해야 일관성 유지
        const verifyName = nameKr || nameEn;
        if (storedPin && _verifyPin(storedPin, pinInput, verifyName)) {
          matched = { nameKr, nameEn };
          matchedRow = i + 1;
          storedPinPlaintext = (storedPin.indexOf(PIN_HASH_PREFIX) !== 0);
          break;
        }
      }
    }

    if (!matched) {
      // 실패 기록 + 락 카운트 증가
      const failResult = _recordAuthFail(nameInput);
      // 사용자 열거 방지를 위해 일관된 에러 메시지
      const resp = {ok: false, error: 'invalid credentials'};
      // 락 임박 경고
      if (failResult.locked) {
        resp.reason = 'now_locked';
        resp.message = '로그인 시도가 너무 많아 계정이 ' + Math.ceil(AUTH_LOCK_MS / 60000) + '분간 잠겼습니다.';
      } else if (failResult.failCount >= AUTH_MAX_FAILS - 2) {
        resp.warning = 'attempts_remaining';
        resp.attemptsLeft = AUTH_MAX_FAILS - failResult.failCount;
      }
      return resp;
    }

    // 로그인 성공 → 실패 카운트 클리어
    _clearAuthFails(nameInput);

    // ── 평문 PIN 자동 업그레이드 (성공 시에만) ──
    if (storedPinPlaintext && matchedRow > 0) {
      try {
        const verifyName = matched.nameKr || matched.nameEn;
        const hashed = _hashPin(pinInput, verifyName);
        sheet.getRange(matchedRow, pinIdx + 1).setValue(hashed);
      } catch (e) {
        // 업그레이드 실패해도 로그인은 진행
      }
    }

    // 관리자 여부 판정
    const isAdmin = ADMIN_NAMES.some(n => matched.nameKr.indexOf(n) >= 0 || matched.nameEn.indexOf(n) >= 0);
    const role = isAdmin ? 'admin' : 'driver';
    const ttl = isAdmin ? TOKEN_TTL_ADMIN_MS : TOKEN_TTL_DRIVER_MS;
    const now = new Date();
    const exp = new Date(now.getTime() + ttl);

    const token = _generateToken();
    const tokenSheet = _getAuthSheet();
    tokenSheet.appendRow([
      token,
      matched.nameKr || matched.nameEn,
      role,
      now.toISOString(),
      exp.toISOString(),
      now.toISOString(),
      userAgent
    ]);

    // 만료 토큰 정리 (확률적으로 실행 - 너무 잦은 정리 방지)
    if (Math.random() < 0.05) _cleanupExpiredTokens();

    return {
      ok: true,
      token: token,
      role: role,
      displayName: matched.nameKr || matched.nameEn,
      expiresAt: exp.toISOString()
    };
  } catch (err) {
    return {ok: false, error: 'login error: ' + err.toString()};
  }
}

function _logoutAction(payload, tokenParam) {
  try {
    const token = tokenParam || payload.token;
    if (!token) return {ok: true}; // 토큰 없어도 OK
    const sheet = _getAuthSheet();
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === token) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return {ok: true};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

function _validateToken(token) {
  // 반환: { valid: bool, role, user, reason }
  if (!token) return {valid: false, reason: 'no_token'};
  try {
    const sheet = _getAuthSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return {valid: false, reason: 'empty'};
    const now = new Date();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === token) {
        const expStr = String(data[i][4] || '');
        const exp = new Date(expStr);
        if (isNaN(exp.getTime()) || exp <= now) {
          // 만료됨 — 삭제
          try { sheet.deleteRow(i + 1); } catch(e) {}
          return {valid: false, reason: 'expired'};
        }
        // LastUsed 갱신 (성능 고려해서 하루에 한 번 정도만)
        try {
          const lastUsedStr = String(data[i][5] || '');
          const lastUsed = new Date(lastUsedStr);
          if (isNaN(lastUsed.getTime()) || (now - lastUsed) > 3600000) {
            sheet.getRange(i + 1, 6).setValue(now.toISOString());
          }
        } catch(e) {}
        return {
          valid: true,
          role: String(data[i][2] || 'driver'),
          user: String(data[i][1] || '')
        };
      }
    }
    return {valid: false, reason: 'not_found'};
  } catch (err) {
    return {valid: false, reason: 'error: ' + err.toString()};
  }
}

function _cleanupExpiredTokens() {
  try {
    const sheet = _getAuthSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
    const now = new Date();
    // 뒤에서부터 삭제 (인덱스 안 꼬이게)
    for (let i = data.length - 1; i >= 1; i--) {
      const expStr = String(data[i][4] || '');
      const exp = new Date(expStr);
      if (isNaN(exp.getTime()) || exp <= now) {
        try { sheet.deleteRow(i + 1); } catch(e) {}
      }
    }
  } catch (err) {
    Logger.log('cleanup error: ' + err.toString());
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// 보안 관리 유틸 함수 (Apps Script 에디터에서 직접 실행)
// ═══════════════════════════════════════════════════════════════════════════

/**
 * 일회성 PIN 마이그레이션: M_Drivers의 모든 평문 PIN을 해시로 변환
 * Apps Script 에디터에서 함수 선택 → 실행
 * (자동 업그레이드도 작동하므로 필수는 아님 — 첫 로그인 시 자동 변환됨)
 */
function migrateAllPinsToHash() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('M_Drivers');
    if (!sheet) { Logger.log('M_Drivers not found'); return; }
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) { Logger.log('no drivers'); return; }
    const headers = data[0].map(String);
    const krIdx = headers.indexOf('Name_KR');
    const enIdx = headers.indexOf('Name_EN');
    const pinIdx = headers.indexOf('PIN');
    if (pinIdx === -1) { Logger.log('PIN column not found'); return; }

    let migrated = 0;
    let alreadyHashed = 0;
    let skipped = 0;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const pin = String(row[pinIdx] || '').trim();
      const verifyName = String(row[krIdx] || row[enIdx] || '').trim();
      if (!pin) { skipped++; continue; }
      if (pin.indexOf(PIN_HASH_PREFIX) === 0) { alreadyHashed++; continue; }
      if (!verifyName) { skipped++; continue; }
      const hashed = _hashPin(pin, verifyName);
      sheet.getRange(i + 1, pinIdx + 1).setValue(hashed);
      migrated++;
    }
    const summary = '✅ PIN 마이그레이션 완료\n  변환: ' + migrated + '명\n  이미 해시: ' + alreadyHashed + '명\n  건너뜀(빈 PIN/이름 없음): ' + skipped + '명';
    Logger.log(summary);
    return summary;
  } catch (err) {
    Logger.log('migration error: ' + err.toString());
    return 'error: ' + err.toString();
  }
}

/**
 * 일회성 마이그레이션: 트레일러 시스템 도입을 위한 시트 헤더 갱신
 * Apps Script 에디터에서 함수 선택 → 실행
 * 변경:
 *   - Daily_Report: Trailer 다음에 Trailer_Number 추가
 *   - Pre_Departure: Signature 다음에 Trailer_Number 추가
 *   - M_Drivers: PIN 다음에 Owner 추가
 *   - M_Trailers 시트 신규 생성
 */
function migrateAddTrailerSystem() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const log = [];

    function ensureColumn(sheetName, colName, afterCol) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) { log.push(sheetName + ': sheet not found, skip'); return; }
      const lastCol = sheet.getLastColumn();
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      if (headers.indexOf(colName) >= 0) {
        log.push(sheetName + '.' + colName + ': already exists');
        return;
      }
      const afterIdx = headers.indexOf(afterCol);
      if (afterIdx < 0) {
        // afterCol이 없으면 맨 끝에 추가
        sheet.getRange(1, lastCol + 1).setValue(colName);
        log.push(sheetName + '.' + colName + ': appended at end');
        return;
      }
      // afterCol 다음 위치에 컬럼 삽입
      sheet.insertColumnAfter(afterIdx + 1);
      sheet.getRange(1, afterIdx + 2).setValue(colName);
      log.push(sheetName + '.' + colName + ': inserted after ' + afterCol);
    }

    ensureColumn('Daily_Report', 'Trailer_Number', 'Trailer');
    ensureColumn('Pre_Departure', 'Trailer_Number', 'Signature');
    ensureColumn('M_Drivers', 'Owner', 'PIN');

    // M_Vehicles에 사진 컬럼 4개 추가 (Active 다음)
    ensureColumn('M_Vehicles', 'Photo_Front', 'Active');
    ensureColumn('M_Vehicles', 'Photo_Back', 'Photo_Front');
    ensureColumn('M_Vehicles', 'Photo_Left', 'Photo_Back');
    ensureColumn('M_Vehicles', 'Photo_Right', 'Photo_Left');

    // M_Trailers 시트 생성 또는 컬럼 추가
    let tSheet = ss.getSheetByName('M_Trailers');
    if (!tSheet) {
      tSheet = ss.insertSheet('M_Trailers');
      tSheet.getRange(1, 1, 1, 7).setValues([['Trailer_Number','Owner','Capacity','Rego_Date','ESafety_Date','Notes','Active']]);
      tSheet.getRange(1, 1, 1, 7).setFontWeight('bold');
      tSheet.setFrozenRows(1);
      log.push('M_Trailers: created');
    } else {
      // 기존 시트라면 새 컬럼 추가 (있으면 스킵)
      const existing = tSheet.getRange(1, 1, 1, tSheet.getLastColumn()).getValues()[0];
      // Capacity 다음에 Rego_Date, ESafety_Date 순서로
      const insertIfMissing = (col, afterCol) => {
        if (existing.indexOf(col) >= 0) return;
        const afterIdx = existing.indexOf(afterCol);
        if (afterIdx >= 0) {
          tSheet.insertColumnAfter(afterIdx + 1);
          tSheet.getRange(1, afterIdx + 2).setValue(col);
          existing.splice(afterIdx + 1, 0, col);
          log.push('M_Trailers.' + col + ': inserted after ' + afterCol);
        } else {
          tSheet.getRange(1, tSheet.getLastColumn() + 1).setValue(col);
          existing.push(col);
          log.push('M_Trailers.' + col + ': appended');
        }
      };
      insertIfMissing('Rego_Date', 'Capacity');
      insertIfMissing('ESafety_Date', 'Rego_Date');
      // 이전 마이그레이션에서 잘못 추가된 HVIS_Date 컬럼은 그대로 둠 (데이터 손실 방지)
      // 사용자가 직접 삭제 가능
      log.push('M_Trailers: already exists');
    }

    Logger.log(log.join('\n'));
    return log.join('\n');
  } catch (err) {
    Logger.log('migrateAddTrailerSystem error: ' + err.toString());
    return 'error: ' + err.toString();
  }
}

/**
 * SUB 차량 운행 — 차주 지급 확인 시스템 마이그레이션
 *
 * Daily_Report 시트에 SUB 차량 운행에 대한 차주 지급 확인 컬럼 5개 추가:
 *   - SubPaid_Owner    : 'Y' / '' (차주가 지급했다고 관리자가 확인)
 *   - SubPaid_OwnerAt  : ISO 타임스탬프
 *   - SubPaid_OwnerBy  : 확인한 관리자/차주명
 *   - SubPaid_Driver   : 'Y' / '' (드라이버가 받았다고 확인)
 *   - SubPaid_DriverAt : ISO 타임스탬프
 *
 * 자사 차량 운행 행에서는 이 컬럼들이 빈 값으로 유지됨 (의미 없음)
 *
 * 사용법: Apps Script 에디터에서 한 번 실행
 */
function migrateAddSubPaymentColumns() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const log = [];

    function ensureColumn(sheetName, colName, afterCol) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) { log.push(sheetName + ': sheet not found, skip'); return; }
      const lastCol = sheet.getLastColumn();
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      if (headers.indexOf(colName) >= 0) {
        log.push(sheetName + '.' + colName + ': already exists');
        return;
      }
      const afterIdx = headers.indexOf(afterCol);
      if (afterIdx < 0) {
        sheet.getRange(1, lastCol + 1).setValue(colName);
        log.push(sheetName + '.' + colName + ': appended at end');
        return;
      }
      sheet.insertColumnAfter(afterIdx + 1);
      sheet.getRange(1, afterIdx + 2).setValue(colName);
      log.push(sheetName + '.' + colName + ': inserted after ' + afterCol);
    }

    // Remarks 다음에 5개 컬럼 순서대로 추가
    ensureColumn('Daily_Report', 'SubPaid_Owner',    'Remarks');
    ensureColumn('Daily_Report', 'SubPaid_OwnerAt',  'SubPaid_Owner');
    ensureColumn('Daily_Report', 'SubPaid_OwnerBy',  'SubPaid_OwnerAt');
    ensureColumn('Daily_Report', 'SubPaid_Driver',   'SubPaid_OwnerBy');
    ensureColumn('Daily_Report', 'SubPaid_DriverAt', 'SubPaid_Driver');

    Logger.log(log.join('\n'));
    return log.join('\n');
  } catch (err) {
    Logger.log('migrateAddSubPaymentColumns error: ' + err.toString());
    return 'error: ' + err.toString();
  }
}

/**
 * SUB 차량 운행 — 차주 지급 확인
 *
 * @param {number} rowIndex - Daily_Report 시트의 1-indexed row (헤더가 1행)
 * @param {string} type - 'owner' (관리자/차주 확인) 또는 'driver' (드라이버 확인)
 * @param {string} user - 확인한 사람 이름
 * @param {boolean} confirmed - true=확인, false=취소
 */
function confirmSubPayment(rowIndex, type, user, confirmed) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Daily_Report');
    if (!sheet) return {ok: false, msg: 'Daily_Report sheet not found'};

    const ri = parseInt(rowIndex);
    if (!ri || ri < 2) return {ok: false, msg: 'Invalid rowIndex'};

    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    const setCell = (colName, value) => {
      const idx = headers.indexOf(colName);
      if (idx < 0) throw new Error('Column not found: ' + colName + ' (run migrateAddSubPaymentColumns first)');
      sheet.getRange(ri, idx + 1).setValue(value);
    };

    const now = new Date().toISOString();
    const isConfirmed = confirmed !== false; // default true

    if (type === 'owner') {
      setCell('SubPaid_Owner',   isConfirmed ? 'Y' : '');
      setCell('SubPaid_OwnerAt', isConfirmed ? now : '');
      setCell('SubPaid_OwnerBy', isConfirmed ? (user || 'unknown') : '');
    } else if (type === 'driver') {
      setCell('SubPaid_Driver',   isConfirmed ? 'Y' : '');
      setCell('SubPaid_DriverAt', isConfirmed ? now : '');
    } else {
      return {ok: false, msg: 'Invalid type: must be owner or driver'};
    }

    // 현재 row 데이터 반환 (UI 갱신용)
    const updatedRow = sheet.getRange(ri, 1, 1, lastCol).getValues()[0];
    const obj = {};
    headers.forEach((h, i) => obj[h] = updatedRow[i]);

    return {ok: true, row: obj};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

/**
 * SUB 차량 운행 — 차주 지급 일괄 확인
 * 한 차주의 여러 row를 한 번에 확인 처리
 *
 * @param {Array<number>} rowIndexes - 1-indexed rows
 * @param {string} type - 'owner' or 'driver'
 * @param {string} user
 * @param {boolean} confirmed
 */
function confirmSubPaymentBulk(rowIndexes, type, user, confirmed) {
  try {
    if (!Array.isArray(rowIndexes) || rowIndexes.length === 0) {
      return {ok: false, msg: 'No rows specified'};
    }
    const results = [];
    for (const ri of rowIndexes) {
      results.push(confirmSubPayment(ri, type, user, confirmed));
    }
    const okCount = results.filter(r => r.ok).length;
    return {ok: okCount > 0, total: rowIndexes.length, success: okCount, results};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

/**
 * 관리자용: 현재 보안 설정 상태 점검
 * Apps Script 에디터에서 실행 → Logger 확인
 */
function _checkSecuritySetup() {
  const log = [];
  log.push('═══════════════════════════════════════');
  log.push('  Dong Choi 시스템 보안 설정 점검');
  log.push('═══════════════════════════════════════');
  log.push('');

  // 1) PIN_HASH_SECRET이 Script Properties에 설정됐는지
  let secretFromProps = null;
  try {
    secretFromProps = PropertiesService.getScriptProperties().getProperty('PIN_HASH_SECRET');
  } catch(e) {}

  if (secretFromProps && secretFromProps.length > 0) {
    log.push('✅ PIN_HASH_SECRET: Script Properties에서 조회 (안전)');
    log.push('   길이: ' + secretFromProps.length + '자');
    if (secretFromProps === PIN_HASH_SECRET_FALLBACK) {
      log.push('   ⚠️ 경고: 기본 secret과 동일함 — 새 secret으로 변경 권장');
    }
  } else {
    log.push('🟡 PIN_HASH_SECRET: 폴백 값 사용 중 (코드에 노출됨)');
    log.push('   조치: Apps Script 프로젝트 설정 → 스크립트 속성 추가');
    log.push('   속성: PIN_HASH_SECRET');
    log.push('   값: (예: DC_2026_xK9pQ3vN7mR_secure_2026)');
  }
  log.push('');

  // 2) Active_Tokens 시트 점검
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const tokenSheet = ss.getSheetByName('Active_Tokens');
    if (tokenSheet) {
      const data = tokenSheet.getDataRange().getValues();
      const tokenCount = Math.max(0, data.length - 1);
      log.push('✅ Active_Tokens 시트 확인: ' + tokenCount + '개 토큰');
    } else {
      log.push('⚠️ Active_Tokens 시트 없음 (첫 로그인 시 자동 생성됨)');
    }
  } catch(e) {
    log.push('❌ Active_Tokens 점검 실패: ' + e.toString());
  }
  log.push('');

  // 3) Auth_Failures 시트 점검 (Rate limiting)
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const failSheet = ss.getSheetByName('Auth_Failures');
    if (failSheet) {
      const data = failSheet.getDataRange().getValues();
      const failCount = Math.max(0, data.length - 1);
      log.push('✅ Auth_Failures 시트 확인: ' + failCount + '개 기록');
    } else {
      log.push('⚠️ Auth_Failures 시트 없음 (첫 실패 시 자동 생성됨)');
    }
  } catch(e) {
    log.push('❌ Auth_Failures 점검 실패: ' + e.toString());
  }
  log.push('');

  // 4) 드라이버 PIN 보안 점검
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const drvSheet = ss.getSheetByName('M_Drivers');
    if (drvSheet) {
      const data = drvSheet.getDataRange().getValues();
      const headers = data[0];
      const pinIdx = headers.indexOf('PIN');
      const nameIdx = headers.indexOf('Name_KR');
      const nameEnIdx = headers.indexOf('Name_EN');
      let totalPins = 0;
      let hashedPins = 0;
      let plainPins = 0;
      let weakPins = 0;
      const weakSet = new Set(['1234','0000','1111','2222','3333','4444','5555','6666','7777','8888','9999','1212','2020','2024','2025','2026','0123','4321','9876','1004']);
      for (let i = 1; i < data.length; i++) {
        const pin = String(data[i][pinIdx] || '').trim();
        const name = data[i][nameIdx] || data[i][nameEnIdx] || '(이름없음)';
        if (!pin) continue;
        totalPins++;
        if (pin.startsWith('h1$')) {
          hashedPins++;
        } else {
          plainPins++;
          log.push('   🟡 평문 PIN 사용: ' + name + ' (재로그인 시 자동 해시됨)');
          if (weakSet.has(pin)) {
            weakPins++;
            log.push('     ❌ 흔한 PIN 사용: ' + name + ' = ' + pin);
          }
        }
      }
      log.push('✅ 드라이버 PIN 점검: 총 ' + totalPins + '개');
      log.push('   • 해시된 PIN: ' + hashedPins + '개');
      log.push('   • 평문 PIN: ' + plainPins + '개');
      if (weakPins > 0) {
        log.push('   ⚠️ 흔한 PIN 사용: ' + weakPins + '명 — 변경 권장!');
      }
    }
  } catch(e) {
    log.push('❌ 드라이버 PIN 점검 실패: ' + e.toString());
  }
  log.push('');
  log.push('═══════════════════════════════════════');
  log.push('점검 완료. 위 권장사항을 검토해주세요.');
  log.push('═══════════════════════════════════════');

  Logger.log(log.join('\n'));
  return log.join('\n');
}

/**
 * 새 PIN_HASH_SECRET을 Script Properties에 설정 + 모든 PIN 재해시
 * ⚠️ 주의: 이 함수는 모든 드라이버의 평문 PIN을 새 secret으로 다시 해시함
 *         이미 해시된 PIN은 영향 없음 (구 secret으로 만들어진 해시는 그대로)
 * 사용법:
 *   1) Script Properties에 새 PIN_HASH_SECRET 설정
 *   2) (선택) 이 함수 실행해서 평문 PIN을 새 secret으로 해시
 */
function _migratePlainPinsWithNewSecret() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const drvSheet = ss.getSheetByName('M_Drivers');
    if (!drvSheet) { Logger.log('M_Drivers 시트 없음'); return; }
    const data = drvSheet.getDataRange().getValues();
    const headers = data[0];
    const pinIdx = headers.indexOf('PIN');
    const nameIdx = headers.indexOf('Name_KR');
    const nameEnIdx = headers.indexOf('Name_EN');
    if (pinIdx < 0) { Logger.log('PIN 컬럼 없음'); return; }
    let migrated = 0;
    for (let i = 1; i < data.length; i++) {
      const pin = String(data[i][pinIdx] || '').trim();
      const name = String(data[i][nameIdx] || data[i][nameEnIdx] || '').trim();
      if (!pin || pin.startsWith('h1$') || !name) continue;
      // 평문 PIN을 새 secret으로 해시 (현재 _getPinSecret()은 이미 새 secret 반환)
      const hashed = _hashPin(pin, name);
      drvSheet.getRange(i + 1, pinIdx + 1).setValue(hashed);
      migrated++;
      Logger.log('✓ ' + name + ': 평문 → 해시 완료');
    }
    Logger.log('=== 마이그레이션 완료: ' + migrated + '개 PIN 해시화 ===');
    return migrated;
  } catch(err) {
    Logger.log('error: ' + err.toString());
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// 자동 백업 시스템 (Daily Backup System)
// ═══════════════════════════════════════════════════════════════════════════
//
// 목적: 데이터 손상/실수 삭제 시 복구 가능하도록 매일 자동 백업
// 흐름:
//   1) 매일 새벽 2시 (Sydney 시간) 시간 트리거 → runDailyBackup() 실행
//   2) 같은 스프레드시트에 _BACKUP_YYYYMMDD 형태로 시트 복제
//   3) 7일 지난 백업은 자동 삭제 (BACKUP_RETENTION_DAYS)
//
// 사용법:
//   • 트리거 등록: setupBackupTrigger() 한 번만 실행
//   • 즉시 백업: runDailyBackup() 실행
//   • 트리거 제거: removeBackupTrigger() 실행
//   • 복구: 백업 시트 내용을 원본 시트에 복사
//
// 백업되는 시트 목록은 BACKUP_SHEETS 상수에서 관리
// ═══════════════════════════════════════════════════════════════════════════

const BACKUP_RETENTION_DAYS = 7;
const BACKUP_SHEET_PREFIX = '_BAK_';

// 백업 대상 시트 (운영에 핵심적인 데이터만)
const BACKUP_SHEETS = [
  'Daily_Report', 'Pre_Departure', 'End_of_Shift',
  'M_Vehicles', 'M_Drivers', 'M_Clients', 'M_Guides', 'M_Hotels', 'M_Trailers',
  'M_PriceClient', 'M_PriceDriver', 'M_PriceSub', 'M_SUB',
  'M_NightRates',
  'Wages', 'Ledger',
  'Invoices', 'Invoice_Manual_Items', 'Invoice_Deductions',
  'Notices', 'Defect_Reports',
  'Leave_Requests', 'MOT_Report',
  'Agency_Txn', 'Sub_Txn',
  'Company_Profile'
];

/**
 * 매일 자동 백업 실행 (트리거에서 호출됨)
 * 또는 수동으로 GAS 에디터에서 실행 가능
 */
function runDailyBackup() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const today = new Date();
    const dateStr = Utilities.formatDate(today, 'Australia/Sydney', 'yyyyMMdd');
    const backupSuffix = BACKUP_SHEET_PREFIX + dateStr;
    const log = [];
    log.push('═══ 자동 백업 시작: ' + dateStr + ' ═══');

    let successCount = 0;
    let skipCount = 0;
    let errorCount = 0;

    // 1) 백업할 시트들 복제
    BACKUP_SHEETS.forEach(sheetName => {
      try {
        const srcSheet = ss.getSheetByName(sheetName);
        if (!srcSheet) {
          log.push('  ⚠️ ' + sheetName + ': 원본 없음 (스킵)');
          skipCount++;
          return;
        }
        const backupName = sheetName + backupSuffix;
        // 이미 같은 날짜 백업이 있으면 스킵 (멱등성)
        const existing = ss.getSheetByName(backupName);
        if (existing) {
          log.push('  ⏭️ ' + backupName + ': 이미 존재 (스킵)');
          skipCount++;
          return;
        }
        // 시트 복제
        const copy = srcSheet.copyTo(ss);
        copy.setName(backupName);
        // 백업 시트는 숨김 처리 (원본과 헷갈림 방지)
        copy.hideSheet();
        log.push('  ✅ ' + backupName);
        successCount++;
      } catch(e) {
        log.push('  ❌ ' + sheetName + ': ' + e.toString());
        errorCount++;
      }
    });

    log.push('───');
    log.push('성공: ' + successCount + ' / 스킵: ' + skipCount + ' / 실패: ' + errorCount);

    // 2) 오래된 백업 삭제 (7일 이상)
    log.push('═══ 오래된 백업 삭제 ═══');
    const cutoffDate = new Date(today.getTime() - BACKUP_RETENTION_DAYS * 86400000);
    const allSheets = ss.getSheets();
    let deletedCount = 0;
    allSheets.forEach(sh => {
      const name = sh.getName();
      // 백업 시트 패턴: <원본이름>_BAK_YYYYMMDD
      const m = name.match(/_BAK_(\d{8})$/);
      if (!m) return;
      const dateStr = m[1];
      const y = parseInt(dateStr.substring(0, 4));
      const mo = parseInt(dateStr.substring(4, 6)) - 1;
      const d = parseInt(dateStr.substring(6, 8));
      const sheetDate = new Date(y, mo, d);
      if (sheetDate < cutoffDate) {
        try {
          ss.deleteSheet(sh);
          log.push('  🗑️ 삭제: ' + name);
          deletedCount++;
        } catch(e) {
          log.push('  ❌ 삭제 실패: ' + name + ' — ' + e.toString());
        }
      }
    });
    log.push('총 ' + deletedCount + '개 오래된 백업 삭제');
    log.push('═══ 백업 완료 ═══');

    Logger.log(log.join('\n'));

    // 백업 결과를 별도 로그 시트에도 기록
    try {
      let logSheet = ss.getSheetByName('_Backup_Log');
      if (!logSheet) {
        logSheet = ss.insertSheet('_Backup_Log');
        logSheet.getRange(1, 1, 1, 5).setValues([['Timestamp', 'Date', 'Success', 'Skipped', 'Errors']]);
        logSheet.setFrozenRows(1);
        logSheet.hideSheet();
      }
      logSheet.appendRow([new Date().toISOString(), dateStr, successCount, skipCount, errorCount]);
    } catch(e) {}

    return log.join('\n');
  } catch (err) {
    Logger.log('runDailyBackup error: ' + err.toString());
    return 'error: ' + err.toString();
  }
}

/**
 * 백업 트리거 등록 (한 번만 실행)
 * 매일 새벽 2시 (Sydney) runDailyBackup 자동 실행
 */
function setupBackupTrigger() {
  // 기존 동일 트리거 제거 (중복 방지)
  removeBackupTrigger();
  // 새 트리거 등록
  ScriptApp.newTrigger('runDailyBackup')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .inTimezone('Australia/Sydney')
    .create();
  Logger.log('✅ 자동 백업 트리거 등록: 매일 새벽 2시 (Sydney 시간)');
  return 'Backup trigger created.';
}

/**
 * 백업 트리거 제거
 */
function removeBackupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'runDailyBackup') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  Logger.log('Removed ' + removed + ' backup trigger(s).');
  return removed;
}

/**
 * 백업 시트 목록 확인
 */
function listBackups() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheets = ss.getSheets();
  const backups = {};
  sheets.forEach(sh => {
    const name = sh.getName();
    const m = name.match(/^(.+?)_BAK_(\d{8})$/);
    if (!m) return;
    const orig = m[1];
    const dateStr = m[2];
    if (!backups[dateStr]) backups[dateStr] = [];
    backups[dateStr].push(orig);
  });
  const log = ['═══ 현재 백업 목록 ═══'];
  Object.keys(backups).sort().reverse().forEach(d => {
    log.push(d + ' (' + backups[d].length + '개): ' + backups[d].join(', '));
  });
  if (Object.keys(backups).length === 0) log.push('백업 없음');
  Logger.log(log.join('\n'));
  return log.join('\n');
}

/**
 * 특정 날짜 백업으로부터 시트 복원
 * 사용법: restoreFromBackup('Daily_Report', '20260425')
 * ⚠️ 주의: 원본 시트의 현재 데이터가 백업으로 덮어씌워짐
 */
function restoreFromBackup(sheetName, dateStr) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const backupName = sheetName + BACKUP_SHEET_PREFIX + dateStr;
    const backupSheet = ss.getSheetByName(backupName);
    if (!backupSheet) {
      Logger.log('❌ 백업 시트 없음: ' + backupName);
      return 'backup not found';
    }
    const origSheet = ss.getSheetByName(sheetName);
    if (!origSheet) {
      Logger.log('❌ 원본 시트 없음: ' + sheetName);
      return 'original not found';
    }
    // 안전장치: 복원 전에 현재 시트를 _BAK_BEFORE_RESTORE_<timestamp>로 백업
    const tsLabel = Utilities.formatDate(new Date(), 'Australia/Sydney', 'yyyyMMdd_HHmmss');
    const safetyBackup = origSheet.copyTo(ss);
    safetyBackup.setName(sheetName + '_BAK_BEFORE_RESTORE_' + tsLabel);
    safetyBackup.hideSheet();
    // 원본 데이터 클리어 후 백업 데이터 복사
    origSheet.clearContents();
    const data = backupSheet.getDataRange().getValues();
    if (data.length > 0 && data[0].length > 0) {
      origSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    }
    Logger.log('✅ 복원 완료: ' + sheetName + ' (백업 날짜: ' + dateStr + ')');
    Logger.log('   안전 백업: ' + safetyBackup.getName());
    return 'restored: ' + sheetName + ' from ' + dateStr;
  } catch (err) {
    Logger.log('restoreFromBackup error: ' + err.toString());
    return 'error: ' + err.toString();
  }
}

/**
 * 관리자용: 특정 사용자의 로그인 잠금 해제
 * Apps Script 에디터에서 _adminUnlockUser 함수의 name을 바꿔서 실행
 */
function _adminUnlockUser() {
  const name = '최동철'; // ← 잠금 해제할 사용자 이름으로 변경
  _clearAuthFails(name);
  Logger.log('✅ 잠금 해제: ' + name);
  return 'unlocked: ' + name;
}

/**
 * 관리자용: 모든 활성 토큰 강제 무효화 (전체 로그아웃)
 * 보안 사고 발생 시 사용
 */
function _adminInvalidateAllTokens() {
  try {
    const sheet = _getAuthSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
    Logger.log('✅ 모든 토큰 무효화됨');
    return 'all tokens cleared';
  } catch (err) {
    Logger.log('error: ' + err.toString());
    return 'error: ' + err.toString();
  }
}

// 로그인용: 드라이버 이름 목록 (PIN 등 민감 정보 제외)
function _getLoginNames() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('M_Drivers');
    if (!sheet) return {ok: false, error: 'M_Drivers not found'};
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return {ok: true, rows: []};
    const headers = data[0].map(String);
    const nameKrIdx = headers.indexOf('Name_KR');
    const nameEnIdx = headers.indexOf('Name_EN');
    const initIdx   = headers.indexOf('Initials');
    const activeIdx = headers.indexOf('Active');
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const active = activeIdx >= 0 ? String(row[activeIdx] || '').toUpperCase() : 'Y';
      if (active && active !== 'Y' && active !== '') continue;
      const nameKr = String(row[nameKrIdx] || '').trim();
      const nameEn = String(row[nameEnIdx] || '').trim();
      if (!nameKr && !nameEn) continue;
      rows.push({
        Name_KR: nameKr,
        Name_EN: nameEn,
        Initials: initIdx >= 0 ? String(row[initIdx] || '') : ''
      });
    }
    return {ok: true, rows: rows};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

// M_Drivers 응답에서 PIN 컬럼 제거 (get_master/get_all_masters 경유 시 사용)
function _stripPinFromDrivers(result) {
  try {
    if (!result || !result.rows) return result;
    const cleanRows = result.rows.map(r => {
      if (!r || typeof r !== 'object') return r;
      const copy = {};
      Object.keys(r).forEach(k => { if (k !== 'PIN') copy[k] = r[k]; });
      return copy;
    });
    return Object.assign({}, result, {rows: cleanRows});
  } catch (err) {
    return result;
  }
}

// 요청 인증 검사 (메인 게이트)
// 반환: { allow: true } 또는 { allow: false, response: <json> }
function _authGate(action, role, tokenValid) {
  // PUBLIC: 무조건 통과
  if (PUBLIC_ACTIONS.indexOf(action) >= 0) return {allow: true};

  // 토큰 없으면 거부
  if (!tokenValid.valid) {
    return {allow: false, response: {ok: false, error: 'unauthorized', reason: tokenValid.reason || 'no_token', authRequired: true}};
  }

  // 관리자 전용 액션 검사
  if (role === 'driver') {
    if (ADMIN_ONLY_ACTIONS.indexOf(action) >= 0 || ADMIN_ONLY_GET_ACTIONS.indexOf(action) >= 0) {
      return {allow: false, response: {ok: false, error: 'forbidden', reason: 'admin_only'}};
    }
  }

  return {allow: true};
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
    } else {
      // ── 기존 시트에 누락된 컬럼 자동 추가 ──
      const expected = MASTER_HEADERS[sheetName];
      if (expected) {
        const lastCol = sheet.getLastColumn();
        const existing = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
        const missing = expected.filter(h => !existing.includes(h));
        if (missing.length > 0) {
          const startCol = lastCol + 1;
          const color = TAB_COLORS[sheetName] || '#1a56db';
          sheet.getRange(1, startCol, 1, missing.length).setValues([missing])
            .setBackground(color).setFontColor('white').setFontWeight('bold');
        }
      }
    }
    return sheet;
  } catch (err) {
    Logger.log('Error in ensureSheet: ' + err.toString());
    throw err;
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// CONSOLIDATED GET Handler (with token auth gate)
// ═══════════════════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    const action = e.parameter.action || 'ping';
    const sheet = e.parameter.sheet || '';
    const driver = e.parameter.driver || '';
    const token = e.parameter.token || '';

    // ── 인증 게이트 ──
    const tokenValid = _validateToken(token);
    const gate = _authGate(action, tokenValid.role, tokenValid);
    if (!gate.allow) return cors(gate.response);

    // 로그인된 드라이버가 다른 드라이버의 데이터를 조회하는 걸 막음
    // (관리자는 모든 드라이버 조회 가능)
    let effectiveDriver = driver;
    if (tokenValid.valid && tokenValid.role === 'driver') {
      // 드라이버 토큰이면 driver 파라미터를 본인으로 강제
      effectiveDriver = tokenValid.user;
    }

    // ── 캐시 우회 (?force_refresh=1 또는 ?nocache=1) ──
    // 클라이언트가 명시적으로 fresh 데이터 필요할 때 사용 (예: 수동 동기화 버튼)
    if (e.parameter.force_refresh === '1' || e.parameter.nocache === '1') {
      if (sheet) {
        try { _invalidateSheetCache(sheet); } catch(err) {}
      }
      // 'all_masters' 가상 키도 무효화 (마스터 조회시)
      if (action === 'get_all_masters' || (sheet && sheet.indexOf('M_') === 0)) {
        try { _invalidateSheetCache('all_masters'); } catch(err) {}
      }
    }

    switch (action) {
      case 'ping':
        return cors({ok: true, msg: 'DC Fleet API ready', ts: new Date().toISOString()});

      // ── 인증 ──
      case 'login': {
        // GET 방식 login은 URL 로그에 PIN이 남을 수 있어 권장하지 않지만 지원
        return cors(_loginAction({
          name: e.parameter.name || '',
          pin: e.parameter.pin || '',
          ua: e.parameter.ua || ''
        }));
      }
      case 'logout':
        return cors(_logoutAction({token: token}, token));
      case 'get_login_names':
        return cors(_getLoginNames());

      case 'get_reports':
        return cors(getReports(sheet, effectiveDriver));

      case 'get_master': {
        const result = getMaster(sheet);
        // M_Drivers 조회 시 PIN 컬럼 제거 (관리자든 드라이버든 무조건)
        if (sheet === 'M_Drivers') return cors(_stripPinFromDrivers(result));
        return cors(result);
      }

      case 'get_all_masters': {
        const result = getAllMasters();
        // M_Drivers 포함 시 PIN 제거
        if (result && result.data && result.data.M_Drivers) {
          const stripped = _stripPinFromDrivers({rows: result.data.M_Drivers});
          result.data.M_Drivers = stripped.rows;
        }
        return cors(result);
      }

      // ★ 관리자 앱 통합 번들 — 한 번의 openById로 모든 필요 데이터 반환
      // 기존 6개 endpoint(get_all_masters, get_sub_rates, get_ledger, get_wages,
      // get_notices, get_max_km, get_price_sub)를 단일 호출로 처리
      case 'get_admin_bundle': {
        const result = getAdminBundle();
        if (result && result.data && result.data.masters && result.data.masters.M_Drivers) {
          const stripped = _stripPinFromDrivers({rows: result.data.masters.M_Drivers});
          result.data.masters.M_Drivers = stripped.rows;
        }
        return cors(result);
      }

      case 'get_audit_log': {
        // 최근 감사 로그 조회 (관리자 전용)
        const limit = parseInt(e.parameter.limit || '200', 10);
        return cors(getAuditLog(limit));
      }

      case 'get_schedule': {
        // 운행 일정 조회 (필터: status, agency, from, to)
        const filters = {
          status: e.parameter.status || '',
          agency: e.parameter.agency || '',
          from:   e.parameter.from   || '',
          to:     e.parameter.to     || ''
        };
        return cors(getSchedule(filters));
      }

      case 'get_schedule_stats': {
        return cors(getScheduleStats());
      }

      // ── EG 리포트 미리보기 (HTML 반환) ──
      case 'preview_eg_report': {
        const reportType = e.parameter.type || 'daily'; // 'daily' | 'weekly'
        const dateParam = e.parameter.date || '';
        const fromParam = e.parameter.from || '';
        const toParam = e.parameter.to || '';
        if(reportType === 'weekly'){
          return cors(sendEGWeeklyReport({ dryRun: true, from: fromParam, to: toParam }));
        }
        return cors(sendEGDailyReport({ dryRun: true, date: dateParam }));
      }

      case 'get_driver_schedule': {
        // 드라이버에게 배정된 일정 조회 (드라이버 앱용 — 인증 불필요, 드라이버 식별만)
        const driver = e.parameter.driver || '';
        const from = e.parameter.from || '';
        const to = e.parameter.to || '';
        return cors(getDriverSchedule(driver, from, to));
      }

      case 'get_payout_overrides': {
        // 외주 지급 오버라이드 + Schedule.BillingEntity 맵 반환 (잔액 페이지에서 사용)
        return cors(getPayoutOverrides());
      }

      case 'get_sub_rates':
        return cors(getSubRatesSheet());

      case 'get_price_sub':
        return cors(getPriceSubSheet());

      case 'get_ledger':
        return cors(getLedgerSheet());

      case 'get_wages':
        return cors(getWagesSheet(effectiveDriver));

      case 'get_mot_reports':
        return cors(getReports('MOT_Report', effectiveDriver));

      case 'get_notices':
        return cors(getNoticesSheet());

      case 'get_invoices':
        return cors(getInvoices());

      case 'get_active_regos':
        return cors(getActiveRegos());

      case 'get_active_trailers':
        return cors(getActiveTrailers());

      case 'lookup_pd_trailer':
        return cors(lookupTrailerForDR({
          date: e.parameter.date,
          driver: e.parameter.driver || effectiveDriver,
          rego: e.parameter.rego
        }));

      case 'get_my_shifts':
        return cors(getMyShifts(effectiveDriver));

      case 'find_shift_for_dr':
        return cors(findShiftForDR(
          effectiveDriver,
          e.parameter.rego,
          e.parameter.date
        ));

      case 'get_max_km':
        return cors(getMaxKMPerRego());

      case 'get_agency_txn':
        return cors(getSheetRows('Agency_Txn'));

      case 'get_agency_balances': {
        const agencyParam = e.parameter.agency || '';
        return cors(getAgencyBalances(agencyParam));
      }

      case 'get_sub_txn':
        return cors(getSheetRows('SUB_Txn'));

      case 'get_defect_reports': {
        // 드라이버 토큰이면 본인 것만 강제 조회 (effectiveDriver는 token user로 강제됨)
        // 관리자 토큰이면 driver 파라미터 그대로 사용 (빈 값이면 전체)
        const defDriver = (tokenValid.valid && tokenValid.role === 'driver')
          ? effectiveDriver
          : (e.parameter.driver || '');
        return cors(getDefectReports(defDriver));
      }

      case 'get_bus_damage': {
        const dmgRego = e.parameter.rego || '';
        return cors(getBusDamage(dmgRego));
      }

      // ── Fatigue Compliance (GET) ──
      case 'get_fatigue_check':
        return cors(getFatigueComplianceCheck());

      case 'get_last_eos':
        return cors(getLastEndOfShift(effectiveDriver));

      // ── Leave Requests (GET) ──
      case 'get_my_leave_requests':
        return cors(getMyLeaveRequests(effectiveDriver));

      case 'get_all_leave_requests':
        return cors(getAllLeaveRequests(e.parameter.filter));

      case 'get_roster': {
        // 드라이버 토큰이면 본인 행만 필터링하여 반환
        const rosterRes = getRosterData(e.parameter.from, e.parameter.to);
        if (rosterRes && rosterRes.ok && tokenValid.valid && tokenValid.role === 'driver') {
          const me = effectiveDriver;
          rosterRes.roster = (rosterRes.roster || []).filter(r => String(r.Driver || '') === me);
        }
        return cors(rosterRes);
      }

      // ── Daily Report Draft (서버 백업) ──
      case 'get_daily_draft':
        return cors(getDailyDraftServer(effectiveDriver || e.parameter.driver));

      default:
        return cors({ok: false, error: 'Unknown action: ' + action});
    }
  } catch (err) {
    return cors({ok: false, error: err.toString()});
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// CONSOLIDATED POST Handler (with token auth gate)
// ═══════════════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const token  = payload.token || '';
    let _user  = payload._user || 'unknown';

    // ── 인증 게이트 ──
    const tokenValid = _validateToken(token);
    const gate = _authGate(action, tokenValid.role, tokenValid);
    if (!gate.allow) return cors(gate.response);

    // 드라이버 토큰이면 _user를 토큰 소유자로 강제 (spoofing 방지)
    if (tokenValid.valid && tokenValid.role === 'driver') {
      _user = tokenValid.user;
      // driver 필드가 payload나 data에 있으면 토큰 소유자로 강제
      if (payload.driver) payload.driver = tokenValid.user;
      if (payload.driverName) payload.driverName = tokenValid.user;
      if (payload.data && typeof payload.data === 'object' && payload.data.Driver) {
        payload.data.Driver = tokenValid.user;
      }
    }

    // ─── 멱등성 게이트 (Request_ID 중복 차단) ───
    // 같은 Request_ID로 들어온 두 번째 요청은 시트에 쓰지 않고 ok=true 반환.
    // 클라이언트 retry queue가 timeout 후 같은 요청을 다시 보내도 중복 저장되지 않음.
    // 적용 액션: write 계열만. read 계열은 멱등성 의미 없음.
    const _IDEMPOTENT_ACTIONS = {
      save_report: 1, save_predeparture: 1, save_endofshift: 1,
      save_defect_report: 1, save_mot_report: 1, save_leave_request: 1,
      save_incident_report: 1, save_sub_report: 1, save_correction_request: 1,
      update_report: 1, delete_report: 1,
      save_invoice: 1, delete_invoice: 1, update_invoice_status: 1,
      add_agency_txn: 1, update_agency_txn: 1, delete_agency_txn: 1,
      add_sub_txn: 1, update_sub_txn: 1, delete_sub_txn: 1,
      add_ledger: 1, update_ledger: 1, delete_ledger: 1,
      add_wage: 1, update_wage: 1, delete_wage: 1,
      add_master: 1, update_master: 1, delete_master: 1,
      save_schedule: 1, delete_schedule: 1, update_schedule_status: 1
    };
    const _reqId = (payload.data && payload.data.Request_ID) ? String(payload.data.Request_ID).trim() : '';
    if (_reqId && _IDEMPOTENT_ACTIONS[action]) {
      try {
        const cache = CacheService.getScriptCache();
        const key = 'rid:' + _reqId;
        const existing = cache.get(key);
        if (existing) {
          // 이미 동일 Request_ID가 처리됨 — 시트에 쓰지 않고 ok 반환
          Logger.log('[Idempotency] duplicate blocked: ' + _reqId + ' action=' + action);
          return cors({ ok: true, idempotent: true, message: 'duplicate request — already processed' });
        }
        // 24시간 동안 이 Request_ID를 기록 (단위: 초, 최대 21600 = 6시간 인 점 주의 → 21600 설정)
        // GAS CacheService는 최대 6시간 지원. 그 이상 보호하려면 Properties Service 필요.
        cache.put(key, '1', 21600);
      } catch(e) { Logger.log('[Idempotency] cache failed: ' + e); }
    }

    // ─── 시트 캐시 자동 무효화 (write 액션) ───
    // doPost 진입 시점에 영향받을 시트 캐시를 미리 삭제 → 처리 직후 read는 fresh
    // 잘못된 write로 무효화만 일어나도 안전 (TTL 60초라 곧 다시 캐싱됨)
    const _ACTION_INVALIDATES = {
      save_report: ['Daily_Report', 'Invoices'],     // DR 변경은 Invoices PaidCR에 영향
      save_predeparture: ['Pre_Departure'],
      save_endofshift: ['End_of_Shift'],
      save_defect_report: ['Defect_Reports'],
      save_mot_report: ['MOT_Report'],
      save_leave_request: ['Leave_Requests'],
      save_incident_report: ['Incident_Reports'],
      save_sub_report: ['Daily_Report'],             // SUB report도 Daily_Report에 저장됨
      update_report: ['Daily_Report', 'Pre_Departure', 'End_of_Shift', 'Invoices'],
      delete_report: ['Daily_Report', 'Pre_Departure', 'End_of_Shift', 'Invoices'],
      save_invoice: ['Invoices'],
      delete_invoice: ['Invoices'],
      update_invoice_status: ['Invoices'],
      add_agency_txn: ['Agency_Txn', 'Invoices'],    // PaidCR 변경
      update_agency_txn: ['Agency_Txn', 'Invoices'],
      delete_agency_txn: ['Agency_Txn', 'Invoices'],
      add_sub_txn: ['SUB_Txn'],
      update_sub_txn: ['SUB_Txn'],
      delete_sub_txn: ['SUB_Txn'],
      add_ledger: ['Ledger'],
      update_ledger: ['Ledger'],
      delete_ledger: ['Ledger'],
      add_wage: ['Wages'],
      update_wage: ['Wages'],
      delete_wage: ['Wages'],
      save_schedule: ['Schedule'],
      delete_schedule: ['Schedule'],
      update_schedule_status: ['Schedule', 'Invoices'],
      // 마스터 — payload.sheet에 시트명 들어있음. all_masters도 함께 무효화 (_invalidateSheetCache가 자동 처리)
      add_master: payload.sheet ? [payload.sheet] : null,
      update_master: payload.sheet ? [payload.sheet] : null,
      delete_master: payload.sheet ? [payload.sheet] : null,
      bulk_update_guide_phones: ['M_Guides']
    };
    const _invalidateList = _ACTION_INVALIDATES[action];
    if (_invalidateList && _invalidateList.length) {
      try { _invalidateSheetCache(_invalidateList); } catch(e) { Logger.log('[cache] invalidate fail: ' + e); }
    }

    switch (action) {
      // ── 인증 (POST) ──
      case 'login':
        return cors(_loginAction({
          name: payload.name || '',
          pin: payload.pin || '',
          ua: payload.ua || ''
        }));
      case 'logout':
        return cors(_logoutAction(payload, token));

      // ── Report Operations ──
      case 'save_report':
        return cors(saveReport('Daily_Report', payload.data));

      case 'release_trailer':
        return cors(releaseTrailer(payload.driver || _user, payload.trailerNum));

      case 'patch_pd_trailer':
        return cors(patchPDTrailer({
          date: payload.date,
          driver: payload.driver || _user,
          rego: payload.rego,
          trailerNum: payload.trailerNum
        }));

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

      // ── Daily Report Draft (서버 백업) ──
      case 'save_daily_draft':
        return cors(saveDailyDraftServer(
          payload.driver || _user,
          payload.draftJSON || ''
        ));

      case 'clear_daily_draft':
        return cors(clearDailyDraftServer(payload.driver || _user));

      case 'submit_mot':
        return cors(saveReport('MOT_Report', payload.data));

      case 'save_correction_request':
        return cors(saveCorrectionRequest(payload));

      // ── SUB 차량 운행 — 차주 지급 확인 ──
      case 'confirm_sub_payment': {
        // type: 'owner' (관리자) | 'driver' (드라이버)
        // 드라이버 토큰이면 type을 driver로 강제 (다른 사람 대신 확인 방지)
        let confirmType = payload.type || 'owner';
        if (tokenValid.valid && tokenValid.role === 'driver') {
          confirmType = 'driver';
        }
        const r = confirmSubPayment(
          payload.rowIndex,
          confirmType,
          _user,
          payload.confirmed !== false
        );
        if (r.ok) appendAuditLog(_user, 'confirm_sub_payment', 'Daily_Report', payload.rowIndex,
          'type:' + confirmType + ' confirmed:' + (payload.confirmed !== false));
        return cors(r);
      }

      case 'confirm_sub_payment_bulk': {
        let confirmType = payload.type || 'owner';
        if (tokenValid.valid && tokenValid.role === 'driver') {
          confirmType = 'driver';
        }
        const r = confirmSubPaymentBulk(
          payload.rowIndexes || [],
          confirmType,
          _user,
          payload.confirmed !== false
        );
        if (r.ok) appendAuditLog(_user, 'confirm_sub_payment_bulk', 'Daily_Report',
          (payload.rowIndexes||[]).join(','),
          'type:' + confirmType + ' count:' + r.success + '/' + r.total);
        return cors(r);
      }

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

      // ── EG TRAVEL 자동 리포트 (수동 트리거 + 자동 트리거가 둘 다 호출) ──
      case 'send_eg_daily_report':
        return cors(sendEGDailyReport(payload || {}));
      case 'send_eg_weekly_report':
        return cors(sendEGWeeklyReport(payload || {}));
      case 'setup_eg_report_triggers':
        return cors(setupEGReportTriggers());

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

      // ── Schedule CRUD (운행 일정) ──
      case 'save_schedule': {
        const r = saveSchedule(payload.data, _user);
        if (r.ok) appendAuditLog(_user, 'save_schedule', 'Schedule', r.tourId||'',
          `${payload.data.Agency||''} ${payload.data.TourCode||''} ${payload.data.StartDate||''}~${payload.data.EndDate||''}`);
        return cors(r);
      }
      case 'delete_schedule': {
        const r = deleteSchedule(payload.tourId);
        if (r.ok) appendAuditLog(_user, 'delete_schedule', 'Schedule', payload.tourId||'', '');
        return cors(r);
      }
      case 'update_schedule_status': {
        const r = updateScheduleStatus(payload.tourId, payload.status, payload.invoiceId);
        if (r.ok) appendAuditLog(_user, 'update_schedule_status', 'Schedule', payload.tourId||'',
          `Status→${payload.status}${payload.invoiceId?' Inv:'+payload.invoiceId:''}`);
        return cors(r);
      }

      // ── PayoutOverride: 외주 지급 자동 판단 수동 오버라이드 ──
      case 'set_payout_override': {
        const r = setPayoutOverride(payload.data, _user);
        if (r.ok) appendAuditLog(_user, 'set_payout_override', 'PayoutOverrides', '',
          `${(payload.data&&payload.data.tourCode)||''}/${(payload.data&&payload.data.subCompany)||''}=${(payload.data&&payload.data.action)||''}`);
        return cors(r);
      }

      // ── 일회성 정리: BillingEntity == SubCompany 인 자동등록 DRSUB 거래 삭제 ──
      case 'cleanup_self_owned_sub_txns': {
        const dryRun = (payload.dryRun !== false); // 기본 dry-run
        const r = cleanupSelfOwnedSubTxns(dryRun);
        if (r.ok && !r.dryRun) appendAuditLog(_user, 'cleanup_self_owned_sub_txns', 'SUB_Txn', '',
          `삭제 ${r.deleted||0}건`);
        return cors(r);
      }

      // ── 일회성 마이그레이션: Schedule 기존 행에 BillingEntity 기본값 'DC' 백필 ──
      case 'migrate_schedule_billing_entity': {
        const r = migrateScheduleBillingEntity();
        if (r.ok) appendAuditLog(_user, 'migrate_schedule_billing_entity', 'Schedule', '',
          `백필 ${r.filled||0}건, 유지 ${r.skipped||0}건`);
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
      case 'update_driver_pin': {
        // 드라이버는 자기 PIN만 변경 가능, 관리자는 누구든 가능
        if (tokenValid.role === 'driver' && payload.driverName !== tokenValid.user) {
          return cors({ok: false, error: 'forbidden', reason: 'can_only_change_own_pin'});
        }
        return cors(updateDriverPin(payload.driverName, payload.pin));
      }

      case 'update_driver_info':
        return cors(updateDriverInfo(payload.driverName, payload.data));

      // ── Defect Reports ──
      case 'save_defect_report':
        return cors(saveDefectReport(payload.data));

      case 'update_defect_status': {
        return cors(updateDefectStatus(payload.id, payload.status, payload.adminNote));
      }

      // ── Bus Damage Markers ──
      case 'save_bus_damage':
        return cors(saveBusDamage(payload.rego, payload.markers, payload.driver));

      // ── Leave Requests (POST) ──
      case 'submit_leave_request':
        return cors(submitLeaveRequest(payload.data));

      case 'review_leave_request':
        return cors(reviewLeaveRequest(payload.data));

      case 'update_roster_cell':
        return cors(updateRosterCell(payload.driver, payload.date, payload.status, _user));

      // ── HVIS Bookings (POST) ──
      case 'save_hvis_booking':
        return cors(saveHvisBooking(payload.data));

      case 'delete_hvis_booking':
        return cors(deleteHvisBooking(payload.id));

      // ── Driver Photo Upload ──
      case 'upload_driver_photo':
        return cors(uploadDriverPhoto(payload.driverName, payload.photoKey, payload.dataUrl, payload.mimeType));

      case 'get_driver_photos':
        return cors(getDriverPhotos(payload.driverName));

      // ── Maint Records (POST) ──
      case 'save_maint_record':
        return cors(saveMaintRecord(payload.data));

      case 'delete_maint_record':
        return cors(deleteSheetRowById('Maint_Records', 'ID', payload.id));

      // ── Invoice Overrides (POST) ──
      case 'save_invoice_override':
        return cors(saveInvoiceOverride(payload.rowKey, payload.value));

      case 'delete_invoice_override':
        return cors(deleteSheetRowById('Invoice_Overrides', 'RowKey', payload.rowKey));

      case 'bulk_save_invoice_overrides':
        return cors(bulkSaveInvoiceOverrides(payload.items));

      // ── Company Profile (POST) ──
      case 'save_company_profile':
        return cors(saveCompanyProfile(payload.data));

      // ── Invoice Deductions (POST) ──
      case 'save_invoice_deduction':
        return cors(saveInvoiceDeduction(payload.data));

      case 'delete_invoice_deduction':
        return cors(deleteSheetRowById('Invoice_Deductions', 'ID', payload.id));

      case 'save_invoice_deductions_bulk':
        return cors(saveInvoiceDeductionsBulk(payload.agency, payload.period, payload.items));

      // ── Invoice Manual Items (POST) ──
      case 'save_invoice_manual_item':
        return cors(saveInvoiceManualItem(payload.data));

      case 'delete_invoice_manual_item':
        return cors(deleteSheetRowById('Invoice_Manual_Items', 'ID', payload.id));

      case 'save_invoice_manual_items_bulk':
        return cors(saveInvoiceManualItemsBulk(payload.agency, payload.period, payload.items));

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
  // ★ 캐싱: driver 필터는 캐시 후 적용 (시트 전체는 한 번만 읽음)
  try {
    const cached = _cachedRead(sheetName, function() {
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

      const rows = data.slice(1).map(function(row, i) {
        const obj = {_rowIndex: i + 2};
        headers.forEach(function(h, idx) {
          obj[h] = formatCell(h, row[idx]);
        });
        return obj;
      });
      return {ok: true, rows: rows};
    });

    if (!cached.ok) return cached;
    let rows = cached.rows || [];
    if (driver) rows = rows.filter(function(r) { return r.Driver === driver; });
    return {ok: true, rows: rows};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

function getMaster(sheetName) {
  return _cachedRead(sheetName, function() { return _getMasterImpl(sheetName); });
}

function _getMasterImpl(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, sheetName); // 누락 컬럼 자동 보정

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

    // 전화번호 컬럼 인덱스 사전 탐색 (앞 0 복원용)
    const PHONE_FIELDS = ['phone','mobile','mobile_1','mobile_2','moblie_2'];
    const phoneColIdxSet = new Set();
    headers.forEach((h, i) => {
      if (PHONE_FIELDS.includes(normalizeKey(h))) phoneColIdxSet.add(i);
    });

    const rows = data.slice(1).map((row, rowIdx) => {
      const obj = {};
      headers.forEach((h, i) => {
        // 시트 헤더를 정규 키로 변환 (공백↔언더스코어 자동 처리)
        const nk = normalizeKey(h);
        let canonKey = (h && normToCanonical[nk]) || h;
        // 별칭 매핑 (예: Phone → Mobile_1)
        if (!normToCanonical[nk] && FIELD_ALIASES[nk]) {
          for (const alias of FIELD_ALIASES[nk]) {
            if (normToCanonical[alias]) { canonKey = normToCanonical[alias]; break; }
          }
        }
        let val = row[i];
        // ★ 전화번호 필드: 앞 0 자동 복원 (Google Sheets 숫자→텍스트 보정)
        if (phoneColIdxSet.has(i) && val !== '' && val !== null && val !== undefined) {
          let s = String(val).replace(/\.0+$/, '').replace(/[^0-9]/g, '');
          if (s.length === 9) s = '0' + s;   // 04xxxxxxxx → 0 복원
          val = s;
        }
        obj[canonKey] = val;
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
  return _cachedRead('all_masters', function() { return _getAllMastersImpl(); });
}

function _getAllMastersImpl() {
  try {
    const sheets = ['M_Vehicles', 'M_Drivers', 'M_Clients', 'M_Guides', 'M_Hotels', 'M_Trailers',
                    'M_PriceClient', 'M_PriceDriver', 'M_PriceSub', 'M_SUB',
                    'M_SvcOptions', 'M_HotelOptions', 'M_DistOptions', 'M_NightRates', 'M_Attractions',
                    'Sub_Rates', 'Ledger', 'MOT_Report', 'HVIS_Bookings',
                    'Maint_Records', 'Invoice_Overrides', 'Company_Profile',
                    'Invoice_Deductions', 'Invoice_Manual_Items'];
    const result = {};

    // ★ 최적화: 스프레드시트를 한 번만 열고 모든 시트를 그 인스턴스로 처리
    // 기존: 각 getMaster() 호출마다 openById 재실행 → 23번 × ~200ms 낭비
    const ss = SpreadsheetApp.openById(SHEET_ID);

    sheets.forEach(name => {
      try {
        const r = _getMasterFast(ss, name);
        result[name] = r.ok ? r.rows : [];
      } catch (e) {
        result[name] = [];
      }
    });

    return {ok: true, data: result};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

// ── getMaster 최적화 버전 (기존 ss 인스턴스 재사용) ──
function _getMasterFast(ss, sheetName) {
  try {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return {ok: true, sheet: sheetName, rows: []};

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return {ok: true, sheet: sheetName, rows: []};

    // ensureSheet 스킵 (읽기 전용이므로 헤더 보정 불필요)
    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = data[0];

    const canonicalHeaders = MASTER_HEADERS[sheetName];
    const normToCanonical = {};
    if (canonicalHeaders) {
      canonicalHeaders.forEach(ch => {
        normToCanonical[normalizeKey(ch)] = ch;
      });
    }

    const PHONE_FIELDS = ['phone','mobile','mobile_1','mobile_2','moblie_2'];
    const phoneColIdxSet = new Set();
    headers.forEach((h, i) => {
      if (PHONE_FIELDS.includes(normalizeKey(h))) phoneColIdxSet.add(i);
    });

    const rows = data.slice(1).map((row, rowIdx) => {
      const obj = {};
      headers.forEach((h, i) => {
        const nk = normalizeKey(h);
        let canonKey = (h && normToCanonical[nk]) || h;
        if (!normToCanonical[nk] && FIELD_ALIASES[nk]) {
          for (const alias of FIELD_ALIASES[nk]) {
            if (normToCanonical[alias]) { canonKey = normToCanonical[alias]; break; }
          }
        }
        let val = row[i];
        if (phoneColIdxSet.has(i) && val !== '' && val !== null && val !== undefined) {
          let s = String(val).replace(/\.0+$/, '').replace(/[^0-9]/g, '');
          if (s.length === 9) s = '0' + s;
          val = s;
        }
        obj[canonKey] = val;
      });
      obj._rowIndex = rowIdx + 2;
      return obj;
    });

    return {ok: true, sheet: sheetName, rows};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// ★ 관리자 앱 통합 번들 — 단일 openById로 6+ endpoint 한번에 처리
// 기존 흐름 (시퀀셜):
//   get_all_masters → get_sub_rates → get_ledger → get_wages → get_notices
//   → get_max_km → get_price_sub  (각각 openById 호출)
// 새 흐름:
//   openById 1회 + 모든 시트 한번에 읽기
// ═══════════════════════════════════════════════════════════════════════════
function getAdminBundle() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);

    // 1) 모든 마스터 시트 (기존 getAllMasters 동일)
    const masterSheets = ['M_Vehicles', 'M_Drivers', 'M_Clients', 'M_Guides', 'M_Hotels', 'M_Trailers',
                    'M_PriceClient', 'M_PriceDriver', 'M_PriceSub', 'M_SUB',
                    'M_SvcOptions', 'M_HotelOptions', 'M_DistOptions', 'M_NightRates', 'M_Attractions',
                    'Sub_Rates', 'Ledger', 'MOT_Report', 'HVIS_Bookings',
                    'Maint_Records', 'Invoice_Overrides', 'Company_Profile',
                    'Invoice_Deductions', 'Invoice_Manual_Items'];
    const masters = {};
    masterSheets.forEach(name => {
      try {
        const r = _getMasterFast(ss, name);
        masters[name] = r.ok ? r.rows : [];
      } catch (e) {
        masters[name] = [];
      }
    });

    // 2) Wages (별도 — driver 필터 없이 전체)
    let wages = [];
    try {
      const wagesResult = _getMasterFast(ss, 'Wages');
      wages = wagesResult.ok ? wagesResult.rows : [];
    } catch (e) { wages = []; }

    // 3) Notices
    let notices = [];
    try {
      const noticesResult = _getMasterFast(ss, 'Notices');
      notices = noticesResult.ok ? noticesResult.rows : [];
    } catch (e) { notices = []; }

    // 4) Max KM per Rego (Pre_Departure + Daily_Report + End_of_Shift 스캔)
    const kmMap = {};
    try {
      const scanForKM = (sheetName, kmFields) => {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) return;
        const lastRow = sheet.getLastRow();
        const lastCol = sheet.getLastColumn();
        if (lastRow < 2 || lastCol < 1) return;
        const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
        const headers = data[0];
        const regoIdx = headers.indexOf('Rego');
        if (regoIdx < 0) return;
        const colIdxs = kmFields.map(f => headers.indexOf(f)).filter(i => i >= 0);
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const rego = String(row[regoIdx] || '').trim();
          if (!rego) continue;
          colIdxs.forEach(ci => {
            const v = parseFloat(row[ci]);
            if (!isNaN(v) && v > 0) {
              if (!kmMap[rego] || v > kmMap[rego]) kmMap[rego] = v;
            }
          });
        }
      };
      scanForKM('Pre_Departure', ['Start_KM']);
      scanForKM('Daily_Report',  ['KM_Start', 'KM_End']);
      scanForKM('End_of_Shift',  ['End_KM']);
    } catch (e) { /* km 실패해도 진행 */ }

    return {
      ok: true,
      data: {
        masters: masters,
        wages: wages,
        notices: notices,
        kmMap: kmMap,
        // sub_rates와 ledger, price_sub은 masters에 이미 포함됨 (Sub_Rates, Ledger, M_PriceSub)
        // 클라이언트는 masters['Sub_Rates'], masters['Ledger'], masters['M_PriceSub']로 접근
      },
      ts: new Date().toISOString()
    };
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
    const rows = r.rows || [];
    // ★ Date 컬럼 정규화 — SUB_Txn/Agency_Txn은 Date/FinishDate 컬럼이 Date 객체로 저장되어 있을 수 있음
    //   클라이언트가 일관된 YYYY-MM-DD 문자열을 받도록 강제 변환 (UTC ISO 직렬화 방지)
    if ((sheetName === 'SUB_Txn' || sheetName === 'Agency_Txn') && rows.length > 0) {
      const dateFields = ['Date', 'FinishDate'];
      rows.forEach(row => {
        dateFields.forEach(f => {
          if (row[f] !== undefined && row[f] !== null && row[f] !== '') {
            const v = row[f];
            if (v instanceof Date) {
              // 시드니 로컬 날짜로 변환 (UTC 직렬화 회피)
              row[f] = Utilities.formatDate(v, 'Australia/Sydney', 'yyyy-MM-dd');
            } else if (typeof v === 'string') {
              // ISO 타임스탬프 (2026-05-11T14:00:00.000Z) → 시드니 날짜
              const m = v.match(/^(\d{4}-\d{2}-\d{2})T/);
              if (m) {
                const d = new Date(v);
                if (!isNaN(d.getTime())) {
                  row[f] = Utilities.formatDate(d, 'Australia/Sydney', 'yyyy-MM-dd');
                }
              }
            }
          }
        });
      });
    }
    return {ok: r.ok, rows: rows};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

// ── 날짜 정규화: 어떤 형식이든 'YYYY-MM-DD' 로 변환 ──
function _normalizeDateISO(val) {
  if (!val) return '';
  // Date 객체
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }
  const s = String(val).trim();
  if (!s) return '';
  // 이미 YYYY-MM-DD?
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + String(m[2]).padStart(2,'0') + '-' + String(m[3]).padStart(2,'0');
  // DD/MM/YYYY?
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + String(m[2]).padStart(2,'0') + '-' + String(m[1]).padStart(2,'0');
  // ISO timestamp?
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})T/);
  if (m) return s.slice(0, 10);
  return s;
}

/**
 * Daily_Report 저장 시 트레일러 사용료 자동 정산
 * - 차량 소유주와 트레일러 소유주가 다르면 SUB_Txn에 거래 자동 생성
 * - SUB 차량 + DC 트레일러: SUB 회사 차변(DR)에 -Rental_Fee 차감 (SUB 지급액 줄어듦)
 *   → 실제로는 운임 지급할 때 차감되어야 하므로, 별도 거래로 +Rental_Fee CR 처리
 * - DC 차량 + SUB 트레일러: 트레일러 소유주(SUB)에게 +Rental_Fee 지급 (DR)
 * - 자동 중복 방지: 같은 (Date + Driver + Trailer + Source) 거래가 이미 있으면 생성 안 함
 */
function _autoCreateTrailerRentalTxn(data) {
  if (!data) return;
  const trailerNum = String(data.Trailer_Number || data.Trailer || '').trim();
  if (!trailerNum) return;
  // Trailer 값이 0이거나 'No' 같은 것은 사용 안 함을 의미
  const trailerUsed = (data.Trailer_Number) || (data.Trailer && Number(data.Trailer) > 0);
  if (!trailerUsed) return;

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const trSheet = ss.getSheetByName('M_Trailers');
  const vSheet = ss.getSheetByName('M_Vehicles');
  if (!trSheet || !vSheet) return;

  // M_Trailers에서 트레일러 소유주 + Rental_Fee 조회
  const trData = trSheet.getDataRange().getValues();
  if (trData.length < 2) return;
  const trH = trData[0];
  const trNumIdx = trH.indexOf('Trailer_Number');
  const trOwnerIdx = trH.indexOf('Owner');
  const trFeeIdx = trH.indexOf('Rental_Fee');
  if (trNumIdx < 0 || trOwnerIdx < 0) return;

  let trOwner = '';
  let trFee = 30;
  for (let i = 1; i < trData.length; i++) {
    if (String(trData[i][trNumIdx] || '').trim() === trailerNum) {
      trOwner = String(trData[i][trOwnerIdx] || '').trim();
      if (trFeeIdx >= 0 && trData[i][trFeeIdx]) {
        trFee = Number(trData[i][trFeeIdx]) || 30;
      }
      break;
    }
  }
  if (!trOwner) return;

  // M_Vehicles에서 차량 소유주 조회
  const rego = String(data.Rego || '').trim();
  if (!rego) return;
  const vData = vSheet.getDataRange().getValues();
  if (vData.length < 2) return;
  const vH = vData[0];
  const vRegoIdx = vH.indexOf('Rego');
  const vOwnerIdx = vH.indexOf('Owner');
  if (vRegoIdx < 0 || vOwnerIdx < 0) return;

  let vehOwner = '';
  for (let i = 1; i < vData.length; i++) {
    if (String(vData[i][vRegoIdx] || '').trim() === rego) {
      vehOwner = String(vData[i][vOwnerIdx] || '').trim();
      break;
    }
  }
  if (!vehOwner) return;

  // 같은 소유주이면 정산 불필요
  if (trOwner === vehOwner) return;

  // DC 회사 정의 (영문/공백 변형 고려)
  const DC_NAMES = ['DONG CHOI PTY LTD', 'DONG CHOI', '동초이'];
  const isVehDC = DC_NAMES.indexOf(vehOwner) >= 0;
  const isTrDC = DC_NAMES.indexOf(trOwner) >= 0;

  // 중복 방지: 같은 날짜 + 같은 트레일러 + 같은 driver의 거래가 이미 있으면 스킵
  const txnSheet = ss.getSheetByName('SUB_Txn') || ss.getSheetByName('Sub_Txn');
  if (!txnSheet) return;
  const sourceId = 'DR-trailer-' + (data.Date || '') + '-' + trailerNum + '-' + (data.Driver || '');
  if (txnSheet.getLastRow() > 1) {
    const tData = txnSheet.getDataRange().getValues();
    const tH = tData[0];
    const remarkIdx = tH.indexOf('Remark');
    if (remarkIdx >= 0) {
      for (let i = 1; i < tData.length; i++) {
        if (String(tData[i][remarkIdx] || '').indexOf(sourceId) >= 0) {
          Logger.log('[trailer rental] already exists: ' + sourceId);
          return;
        }
      }
    }
  }

  // 거래 생성
  let subCo, dr, descPrefix;
  if (isVehDC && !isTrDC) {
    // DC 차량 + SUB 트레일러: SUB에게 사용료 지급 (DR)
    subCo = trOwner;
    dr = trFee;
    descPrefix = '트레일러 ' + trailerNum + ' 사용료';
  } else if (!isVehDC && isTrDC) {
    // SUB 차량 + DC 트레일러: SUB 운임에서 차감 (CR — 우리가 받을 돈)
    // SUB가 우리에게 트레일러 빌렸으니 우리가 SUB에게 받을 금액 = +CR
    subCo = vehOwner;
    dr = 0;
    descPrefix = '트레일러 ' + trailerNum + ' 사용료 (자사 트레일러 빌림)';
  } else {
    // 양쪽 모두 SUB (이론적으로 가능, 다른 SUB끼리)
    // 트레일러 소유주가 받음
    subCo = trOwner;
    dr = trFee;
    descPrefix = '트레일러 ' + trailerNum + ' 사용료';
  }

  const dateISO = _normalizeDateISO(data.Date) || data.Date;
  const txnData = {
    SubCompany: subCo,
    Category: 'trailer',
    Date: dateISO,
    InvoiceNo: '',
    Description: descPrefix + ' · DR(' + (data.Driver || '') + ' / ' + rego + ')',
    DR: dr,
    CR: dr === 0 ? trFee : 0,  // SUB 차량 + DC 트레일러일 때 CR=trFee (받을 돈)
    Remark: 'DR 자동 · ' + sourceId
  };

  const r = addMasterRow('SUB_Txn', txnData);
  if (r.ok) {
    appendAuditLog('system', 'auto_trailer_txn', 'SUB_Txn', r.row || '',
      'Sub:' + subCo + ' DR:' + dr + ' CR:' + (txnData.CR));
    Logger.log('[trailer rental] auto-created: ' + JSON.stringify(txnData));
  } else {
    Logger.log('[trailer rental] failed: ' + JSON.stringify(r));
  }
}

/**
 * Daily_Report 수정/삭제 시 자동 생성된 트레일러 정산 거래 삭제
 * Source ID로 매칭: 'DR-trailer-{date}-{trailer}-{driver}'
 * 같은 source ID를 가진 모든 SUB_Txn 행 삭제
 * (수정 시: 삭제 후 _autoCreateTrailerRentalTxn 다시 호출)
 */
function _deleteTrailerRentalTxn(oldData) {
  if (!oldData) return 0;
  const trailerNum = String(oldData.Trailer_Number || '').trim();
  if (!trailerNum) return 0;
  // 식별자 — saveReport에서 만든 것과 동일 형식
  const sourceId = 'DR-trailer-' + (oldData.Date || '') + '-' + trailerNum + '-' + (oldData.Driver || '');

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const txnSheet = ss.getSheetByName('SUB_Txn') || ss.getSheetByName('Sub_Txn');
  if (!txnSheet || txnSheet.getLastRow() < 2) return 0;

  const tData = txnSheet.getDataRange().getValues();
  const tH = tData[0];
  const remarkIdx = tH.indexOf('Remark');
  if (remarkIdx < 0) return 0;

  // 뒤에서부터 삭제 (인덱스 흐트러짐 방지)
  let deleted = 0;
  for (let i = tData.length - 1; i >= 1; i--) {
    if (String(tData[i][remarkIdx] || '').indexOf(sourceId) >= 0) {
      txnSheet.deleteRow(i + 1); // 1-indexed
      deleted++;
    }
  }
  if (deleted > 0) {
    Logger.log('[trailer rental] deleted ' + deleted + ' txns for: ' + sourceId);
    appendAuditLog('system', 'auto_trailer_txn_delete', 'SUB_Txn', '',
      'Deleted ' + deleted + ' txns: ' + sourceId);
  }
  return deleted;
}

/**
 * Daily Report 저장 시 인보이스 드래프트(Manual Items)에 항목 자동 추가
 *
 * 식별 키: TourCode + Date + Driver + Rego (같은 운행 1건)
 * 같은 키의 항목이 이미 있으면 → 업데이트 (DR 데이터 우선)
 * 없으면 → 신규 추가
 *
 * Period: TourCode가 있으면 'TC-{TourCode}', 없으면 'AG-{Agency}-{YYYY-MM}' (월별 그룹)
 */
function _autoAddInvoiceDraftItem(data) {
  if (!data) return;
  const agency = String(data.Agency || '').trim();
  const tourCode = String(data.Tour_Code || '').trim();
  const date = _normalizeDateISO(data.Date) || data.Date;
  const driver = String(data.Driver || '').trim();
  const rego = String(data.Rego || '').trim();

  if (!agency || !date || !driver) return; // 필수 정보 없음
  // 자체운행/Private은 청구 안 함
  if (String(data.Night_Owner || '').toLowerCase() === 'private') return;

  // ★★ BillingEntity 분기 — DC가 인보이스 발행할 운행만 등록
  //    BillingEntity = DC (또는 비어있음 = 기본 자사) → 정상 등록
  //    BillingEntity = 다른 회사 (EG TRAVEL 등) → 그 회사가 자체 발행 → 등록 안 함
  if (!isBillingEntityDC_(data.Billing_Entity || data.BillingEntity || '')) {
    return; // 비-DC 발행 운행 → Manual Items 등록 안 함
  }

  // Period 결정 — TourCode 있으면 TC 단위, 없으면 월별
  const period = tourCode ? ('TC-' + tourCode) : ('AG-' + agency + '-' + date.slice(0,7));

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ensureSheet(ss, 'Invoice_Manual_Items');
  const headers = MASTER_HEADERS['Invoice_Manual_Items'];
  // 시트 헤더가 비어있으면 생성
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#1a56db').setFontColor('white').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // Source ID — DR 동기화용 (수정/삭제 시 매칭)
  // 형식: 'DR-draft-{date}-{tourCode}-{driver}-{rego}'
  const sourceId = 'DR-draft-' + date + '-' + (tourCode || 'NOTC') + '-' + driver + '-' + rego;

  // 기존 항목 검색 (Source ID가 Note에 포함되어 있는지)
  const lastCol = sheet.getLastColumn();
  const actualHeaders = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : headers;
  const noteIdx = actualHeaders.indexOf('Note');
  const periodIdx = actualHeaders.indexOf('Period');
  const idIdx = actualHeaders.indexOf('ID');

  let existingRow = -1;
  if (sheet.getLastRow() > 1 && noteIdx >= 0) {
    const allData = sheet.getDataRange().getValues();
    for (let i = 1; i < allData.length; i++) {
      const noteVal = String(allData[i][noteIdx] || '');
      if (noteVal.indexOf(sourceId) >= 0) {
        existingRow = i + 1; // 1-indexed
        break;
      }
    }
  }

  // 항목 데이터 구성
  const baseAmount = Number(data.SVC_Charge) || 0;
  const hotel = Number(data.Hotel_Surcharge) || 0;
  const dist = Number(data.Dist_Surcharge) || 0;
  const ot = Number(data.OT) || 0;
  const trailer = Number(data.Trailer) || 0;
  const totalTA = Number(data.Total_TA) || (baseAmount + hotel + dist + ot + trailer);

  const itemId = existingRow > 0 ? '' : ('IT-' + Date.now() + '-' + Math.random().toString(36).slice(2,8));
  // Note에 source ID 포함 (수정/삭제 매칭용) + 자동 생성 표시
  const noteText = '[자동·DR] ' + sourceId + (data.Remarks ? ' · ' + String(data.Remarks).slice(0,80) : '');

  const rowData = {
    ID: itemId,
    Agency: agency,
    Period: period,
    Date: date,
    Rego: rego,
    Tour: data.Attraction || '',
    Seats: data.Seats || '',
    TourCode: tourCode,
    Note: noteText,
    Amount: baseAmount,
    OT: ot,
    Hotel: hotel,
    Dist: dist,
    Trailer: trailer,
    Toll: Number(data.Toll) || 0,
    Start: data.Time_Start || '',
    End: data.Time_End || '',
    Driver: driver,
    Guide: data.Guide || '',
    Pickup: data.Pickup || '',
    Dropoff: data.Dropoff || ''
  };

  if (existingRow > 0) {
    // 업데이트 — 기존 ID는 보존
    if (idIdx >= 0) {
      const existingId = sheet.getRange(existingRow, idIdx + 1).getValue();
      if (existingId) rowData.ID = existingId;
    }
    const row = actualHeaders.map(h => rowData[h] !== undefined ? rowData[h] : '');
    sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
    Logger.log('[invoice draft] updated: ' + sourceId);
  } else {
    // 신규 추가
    const row = actualHeaders.map(h => rowData[h] !== undefined ? rowData[h] : '');
    sheet.appendRow(row);
    Logger.log('[invoice draft] added: ' + sourceId);
    appendAuditLog('system', 'auto_invoice_draft', 'Invoice_Manual_Items', sheet.getLastRow(),
      'Period:' + period + ' Date:' + date + ' Amount:' + totalTA);
  }
}

/**
 * Daily Report 수정/삭제 시 자동 생성된 인보이스 드래프트 항목 삭제
 * Source ID로 매칭: 'DR-draft-{date}-{tourCode}-{driver}-{rego}'
 */
function _deleteInvoiceDraftItem(oldData) {
  if (!oldData) return 0;
  const date = _normalizeDateISO(oldData.Date) || oldData.Date;
  const tourCode = String(oldData.Tour_Code || '').trim();
  const driver = String(oldData.Driver || '').trim();
  const rego = String(oldData.Rego || '').trim();

  if (!date || !driver) return 0;

  const sourceId = 'DR-draft-' + date + '-' + (tourCode || 'NOTC') + '-' + driver + '-' + rego;

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Invoice_Manual_Items');
  if (!sheet || sheet.getLastRow() < 2) return 0;

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const noteIdx = headers.indexOf('Note');
  if (noteIdx < 0) return 0;

  const allData = sheet.getDataRange().getValues();
  let deleted = 0;
  // 뒤에서부터 삭제
  for (let i = allData.length - 1; i >= 1; i--) {
    if (String(allData[i][noteIdx] || '').indexOf(sourceId) >= 0) {
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }
  if (deleted > 0) {
    Logger.log('[invoice draft] deleted ' + deleted + ' items for: ' + sourceId);
    appendAuditLog('system', 'auto_invoice_draft_delete', 'Invoice_Manual_Items', '',
      'Deleted ' + deleted + ' items: ' + sourceId);
  }
  return deleted;
}

function _todayISO_Sydney() {
  const now = new Date();
  // 호주 동부 표준시 보정 (서머타임 무시 — Pre_Departure는 ±1일 허용 범위에서 비교됨)
  const sydOffset = 10 * 60;
  const utc = now.getTime() + now.getTimezoneOffset() * 60000;
  const syd = new Date(utc + sydOffset * 60000);
  const yy = syd.getFullYear();
  const mm = String(syd.getMonth() + 1).padStart(2, '0');
  const dd = String(syd.getDate()).padStart(2, '0');
  return yy + '-' + mm + '-' + dd;
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

    const todayISO = _todayISO_Sydney();

    const preRows = preData.slice(1).map(row => {
      const obj = {};
      preHeaders.forEach((h, i) => obj[h] = row[i]);
      obj._iso = _normalizeDateISO(obj.Date);
      return obj;
    }).filter(r => r._iso === todayISO);

    // Collect EoS data for today
    const eosSet = new Set();
    if (eosSheet && eosSheet.getLastRow() > 1) {
      const eosData = eosSheet.getDataRange().getValues();
      const eosH = eosData[0];
      eosData.slice(1).forEach(row => {
        const obj = {};
        eosH.forEach((h, i) => obj[h] = row[i]);
        const iso = _normalizeDateISO(obj.Date);
        if (iso === todayISO) {
          eosSet.add(String(obj.Rego).trim() + '|' + iso);
        }
      });
    }

    // Find active regos (Pre_Departure without End_of_Shift)
    const active = [];
    const seen = new Set();
    preRows.forEach(r => {
      const regoKey = String(r.Rego).trim() + '|' + r._iso;
      if (!eosSet.has(regoKey) && !seen.has(regoKey)) {
        seen.add(regoKey);
        const driverName = String(r.Driver || '').trim();
        active.push({
          driver: driverName || 'Unknown',
          rego: String(r.Rego).trim(),
          date: r._iso,
          startTime: String(r.Start_Time || '').trim()
        });
      }
    });

    return {ok: true, regos: active};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

// ═══════════════════════════════════════════════════════════════════
// 트레일러 잠금 시스템
// 트레일러 잠금 = Pre_Departure에 Trailer_Number 기록 + End_of_Shift 없음
// "트레일러 반납" 시 Pre_Departure 행의 Trailer_Number를 비움
// ═══════════════════════════════════════════════════════════════════
function getActiveTrailers() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const preSheet = ss.getSheetByName('Pre_Departure');
    const eosSheet = ss.getSheetByName('End_of_Shift');
    if (!preSheet) return {ok: true, trailers: []};

    const preData = preSheet.getDataRange().getValues();
    if (preData.length < 2) return {ok: true, trailers: []};
    const preH = preData[0];

    const todayISO = _todayISO_Sydney();

    // Pre_Departure 오늘 행 + Trailer_Number 있는 행만
    const preRows = preData.slice(1).map((row, idx) => {
      const obj = {};
      preH.forEach((h, i) => obj[h] = row[i]);
      obj._iso = _normalizeDateISO(obj.Date);
      obj._rowIndex = idx + 2; // 시트 행 번호 (1-based + 헤더)
      return obj;
    }).filter(r => r._iso === todayISO && String(r.Trailer_Number || '').trim());

    // 오늘 EOS된 차량 찾기 (Rego 기준 — 차량 마감 = 트레일러도 마감)
    const eosSet = new Set();
    if (eosSheet && eosSheet.getLastRow() > 1) {
      const eosData = eosSheet.getDataRange().getValues();
      const eosH = eosData[0];
      eosData.slice(1).forEach(row => {
        const obj = {};
        eosH.forEach((h, i) => obj[h] = row[i]);
        const iso = _normalizeDateISO(obj.Date);
        if (iso === todayISO) {
          eosSet.add(String(obj.Rego).trim() + '|' + iso);
        }
      });
    }

    const active = [];
    const seen = new Set();
    preRows.forEach(r => {
      const trailer = String(r.Trailer_Number || '').trim();
      if (!trailer) return;
      // 차량이 EOS 됐으면 트레일러도 자동 반납
      const regoKey = String(r.Rego).trim() + '|' + r._iso;
      if (eosSet.has(regoKey)) return;
      if (seen.has(trailer)) return;
      seen.add(trailer);
      active.push({
        trailer: trailer,
        driver: String(r.Driver || '').trim() || 'Unknown',
        rego: String(r.Rego).trim(),
        date: r._iso,
        startTime: String(r.Start_Time || '').trim(),
        rowIndex: r._rowIndex
      });
    });

    return {ok: true, trailers: active};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

// DR 저장 직전 검증용: 같은 (날짜, 드라이버, 차량)의 PD에서 트레일러 정보 조회
// 반환: {ok, pdTrailer: 'TR-001' or '', hasPDTrailer: bool}
function lookupTrailerForDR(opts) {
  try {
    opts = opts || {};
    const date = String(opts.date || '').trim();
    const driver = String(opts.driver || '').trim();
    const rego = String(opts.rego || '').trim();
    if (!date || !driver || !rego) {
      return {ok: false, error: 'missing date/driver/rego'};
    }
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const preSheet = ss.getSheetByName('Pre_Departure');
    if (!preSheet) return {ok: true, pdTrailer: '', hasPDTrailer: false};
    const data = preSheet.getDataRange().getValues();
    if (data.length < 2) return {ok: true, pdTrailer: '', hasPDTrailer: false};
    const headers = data[0];
    const idx = {};
    headers.forEach((h, i) => { idx[String(h)] = i; });
    const targetISO = _normalizeDateISO(date);

    // 가장 최근의 PD (같은 날짜+드라이버+차량) 찾기
    let foundTrailer = '';
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const rowISO = _normalizeDateISO(row[idx.Date]);
      if (rowISO !== targetISO) continue;
      if (String(row[idx.Driver] || '').trim() !== driver) continue;
      if (String(row[idx.Rego] || '').trim() !== rego) continue;
      foundTrailer = String(row[idx.Trailer_Number] || '').trim();
      break;
    }
    return {
      ok: true,
      pdTrailer: foundTrailer,
      hasPDTrailer: !!foundTrailer
    };
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

// PD에 트레일러 번호 사후 추가 (DR 작성 중 누락이 발견된 경우)
function patchPDTrailer(opts) {
  try {
    opts = opts || {};
    const date = String(opts.date || '').trim();
    const driver = String(opts.driver || '').trim();
    const rego = String(opts.rego || '').trim();
    const trailerNum = String(opts.trailerNum || '').trim();
    if (!date || !driver || !rego || !trailerNum) {
      return {ok: false, error: 'missing required field'};
    }
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const preSheet = ss.getSheetByName('Pre_Departure');
    if (!preSheet) return {ok: false, error: 'Pre_Departure sheet not found'};
    const data = preSheet.getDataRange().getValues();
    if (data.length < 2) return {ok: false, error: 'no PD rows'};
    const headers = data[0];
    const idx = {};
    headers.forEach((h, i) => { idx[String(h)] = i; });
    const targetISO = _normalizeDateISO(date);

    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const rowISO = _normalizeDateISO(row[idx.Date]);
      if (rowISO !== targetISO) continue;
      if (String(row[idx.Driver] || '').trim() !== driver) continue;
      if (String(row[idx.Rego] || '').trim() !== rego) continue;
      // 해당 PD 행 발견 → Trailer_Number 셀 업데이트
      preSheet.getRange(i + 1, idx.Trailer_Number + 1).setValue(trailerNum);
      return {ok: true, updated: true, rowIndex: i + 1};
    }
    return {ok: false, error: 'matching PD not found'};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

// 트레일러 반납: Pre_Departure 행의 Trailer_Number 셀 비우기
function releaseTrailer(driver, trailerNum) {
  try {
    if (!driver || !trailerNum) return {ok: false, error: 'driver and trailerNum required'};
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const preSheet = ss.getSheetByName('Pre_Departure');
    if (!preSheet) return {ok: false, error: 'Pre_Departure sheet not found'};

    const data = preSheet.getDataRange().getValues();
    const headers = data[0];
    const idxDriver = headers.indexOf('Driver');
    const idxDate = headers.indexOf('Date');
    const idxTN = headers.indexOf('Trailer_Number');
    if (idxTN < 0) return {ok: false, error: 'Trailer_Number column missing — add it to Pre_Departure sheet'};

    const todayISO = _todayISO_Sydney();
    const trailer = String(trailerNum).trim();
    const driverName = String(driver).trim();

    // 가장 최근의 매칭 행 찾기 (역방향 검색)
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const rowDriver = String(row[idxDriver] || '').trim();
      const rowDate = _normalizeDateISO(row[idxDate]);
      const rowTrailer = String(row[idxTN] || '').trim();
      if (rowDriver === driverName && rowDate === todayISO && rowTrailer === trailer) {
        // 셀 비우기
        preSheet.getRange(i + 1, idxTN + 1).setValue('');
        return {ok: true, msg: 'Trailer ' + trailer + ' released', rowIndex: i + 1};
      }
    }
    return {ok: false, error: 'No matching active trailer found for ' + driverName + ' / ' + trailer};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

// ── 특정 드라이버의 미완료 shift 조회 (날짜 무관) ──
function getMyShifts(driverName) {
  try {
    if (!driverName) return {ok: false, msg: 'driver param required'};
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const preSheet = ss.getSheetByName('Pre_Departure');
    const eosSheet = ss.getSheetByName('End_of_Shift');
    if (!preSheet) return {ok: true, shifts: []};

    const preData = preSheet.getDataRange().getValues();
    if (preData.length < 2) return {ok: true, shifts: []};
    const preH = preData[0];

    // 해당 드라이버의 Pre_Departure 기록 추출
    const myPres = preData.slice(1).map(row => {
      const obj = {};
      preH.forEach((h, i) => obj[h] = row[i]);
      return obj;
    }).filter(r => String(r.Driver||'').trim() === driverName.trim());

    // 날짜를 dd/MM/yyyy 형식으로 통일하는 헬퍼
    const fmtD = v => (v instanceof Date) ? formatDateForSheet(v) : String(v||'').trim();
    const fmtT = v => {
      if (v instanceof Date) return Utilities.formatDate(v, 'Australia/Sydney', 'HH:mm');
      return String(v||'').trim();
    };

    // End_of_Shift 완료된 (Driver + Date + Rego) 조합 수집
    const eosSet = new Set();
    if (eosSheet && eosSheet.getLastRow() > 1) {
      const eosData = eosSheet.getDataRange().getValues();
      const eosH = eosData[0];
      eosData.slice(1).forEach(row => {
        const obj = {};
        eosH.forEach((h, i) => obj[h] = row[i]);
        if (String(obj.Driver||'').trim() === driverName.trim()) {
          eosSet.add(String(obj.Rego).trim() + '|' + fmtD(obj.Date));
        }
      });
    }

    // 미완료 shift 필터링
    const shifts = [];
    const seen = new Set();
    myPres.forEach(r => {
      const dateStr = fmtD(r.Date);
      const key = String(r.Rego).trim() + '|' + dateStr;
      if (!eosSet.has(key) && !seen.has(key)) {
        seen.add(key);
        shifts.push({
          rego: String(r.Rego).trim(),
          date: dateStr,
          seats: String(r.Seats || '').trim(),
          startKm: Number(r.Start_KM) || 0,
          startTime: fmtT(r.Start_Time),
          fuel: String(r.Fuel || '').trim()
        });
      }
    });

    return {ok: true, shifts};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

/**
 * findShiftForDR — Daily Report 누락 일정용 시프트 검색
 *
 * 목적: 특정 드라이버가 특정 차량+날짜로 Pre_Departure를 작성한 적이 있는지 확인.
 *       EOS 완료 여부와 무관하게 반환 (이미 닫힌 시프트에도 Daily Report만 추가할 수 있도록).
 *
 * 매칭 우선순위:
 *   1. 같은 드라이버 + 같은 차량 + 같은 날짜의 Pre_Departure (정확 매칭)
 *   2. 같은 드라이버 + 같은 차량 (날짜 무관) — 가장 가까운 날짜의 Pre 반환
 *
 * 반환: { ok, shift: {rego, date, seats, startKm, startTime, fuel, closed} }
 *       closed: true면 이미 EOS까지 완료된 시프트 (Daily Report만 추가 가능)
 */
function findShiftForDR(driverName, rego, date) {
  try {
    if (!driverName) return {ok: false, msg: 'driver param required'};
    if (!rego) return {ok: false, msg: 'rego param required'};
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const preSheet = ss.getSheetByName('Pre_Departure');
    const eosSheet = ss.getSheetByName('End_of_Shift');
    if (!preSheet) return {ok: true, shift: null};

    const preData = preSheet.getDataRange().getValues();
    if (preData.length < 2) return {ok: true, shift: null};
    const preH = preData[0];

    const fmtD = v => (v instanceof Date) ? formatDateForSheet(v) : String(v||'').trim();
    const fmtT = v => {
      if (v instanceof Date) return Utilities.formatDate(v, 'Australia/Sydney', 'HH:mm');
      return String(v||'').trim();
    };

    const targetDriver = String(driverName).trim();
    const targetRego = String(rego).trim().toUpperCase();
    // date 입력은 dd/MM/yyyy 또는 YYYY-MM-DD 모두 가능 — 정규화
    let targetDate = String(date||'').trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(targetDate)) {
      // YYYY-MM-DD → dd/MM/yyyy
      const parts = targetDate.slice(0, 10).split('-');
      targetDate = parts[2] + '/' + parts[1] + '/' + parts[0];
    }

    // 해당 드라이버의 Pre_Departure 기록 추출
    const myPres = preData.slice(1).map(row => {
      const obj = {};
      preH.forEach((h, i) => obj[h] = row[i]);
      return obj;
    }).filter(r =>
      String(r.Driver||'').trim() === targetDriver &&
      String(r.Rego||'').trim().toUpperCase() === targetRego
    );

    if (!myPres.length) return {ok: true, shift: null};

    // 정확 매칭 우선
    let match = null;
    if (targetDate) {
      match = myPres.find(r => fmtD(r.Date) === targetDate);
    }
    // Fallback: 가장 가까운 날짜의 Pre (날짜 미입력 또는 매칭 실패)
    if (!match) {
      // 날짜순 정렬 (최신 우선)
      myPres.sort((a, b) => {
        const da = fmtD(a.Date), db = fmtD(b.Date);
        // dd/MM/yyyy를 YYYYMMDD로 변환해 비교
        const _conv = s => {
          const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
          return m ? (m[3] + m[2] + m[1]) : '';
        };
        return _conv(db).localeCompare(_conv(da));
      });
      match = myPres[0];
    }

    if (!match) return {ok: true, shift: null};

    const matchDateStr = fmtD(match.Date);

    // EOS 완료 여부 확인
    let closed = false;
    if (eosSheet && eosSheet.getLastRow() > 1) {
      const eosData = eosSheet.getDataRange().getValues();
      const eosH = eosData[0];
      for (let i = 1; i < eosData.length; i++) {
        const row = eosData[i];
        const obj = {};
        eosH.forEach((h, idx) => obj[h] = row[idx]);
        if (String(obj.Driver||'').trim() === targetDriver &&
            String(obj.Rego||'').trim().toUpperCase() === targetRego &&
            fmtD(obj.Date) === matchDateStr) {
          closed = true;
          break;
        }
      }
    }

    return {
      ok: true,
      shift: {
        rego: String(match.Rego).trim(),
        date: matchDateStr,
        seats: String(match.Seats || '').trim(),
        startKm: Number(match.Start_KM) || 0,
        startTime: fmtT(match.Start_Time),
        fuel: String(match.Fuel || '').trim(),
        closed: closed,
        exactDateMatch: matchDateStr === targetDate
      }
    };
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

    // ★ Pre_Departure: 같은 날짜에 같은 차량을 다른 드라이버가 잠갔는지 서버단 검증 (race condition 방지)
    if (sheetName === 'Pre_Departure') {
      const myDriver = String(data.Driver || '').trim();
      const myRego = String(data.Rego || '').trim();
      const myDate = _normalizeDateISO(data.Date);
      const myTrailer = String(data.Trailer_Number || '').trim();
      if (myRego && myDate) {
        const active = getActiveRegos();
        if (active.ok && active.regos) {
          const conflict = active.regos.find(r =>
            r.rego === myRego && r.date === myDate && r.driver !== myDriver
          );
          if (conflict) {
            return {
              ok: false,
              code: 'VEHICLE_LOCKED',
              error: '차량 ' + myRego + '은(는) 이미 ' + conflict.driver + ' 드라이버가 운행 중입니다.',
              conflict: conflict
            };
          }
        }
      }
      // ★ 트레일러 충돌 검사
      if (myTrailer) {
        const activeTr = getActiveTrailers();
        if (activeTr.ok && activeTr.trailers) {
          const trConflict = activeTr.trailers.find(t =>
            t.trailer === myTrailer && t.driver !== myDriver
          );
          if (trConflict) {
            return {
              ok: false,
              code: 'TRAILER_LOCKED',
              error: '트레일러 ' + myTrailer + '은(는) 이미 ' + trConflict.driver + ' 드라이버가 사용 중입니다.',
              conflict: trConflict
            };
          }
        }
      }
    }

    // ★★ Daily_Report: TourCode가 Schedule에 매칭되면 Billing_Entity를 강제로 일정 값으로 덮어쓰기
    //   드라이버 앱 클라이언트 측 lock을 우회한 경우(개발자 도구 등)나
    //   prefill 이후 사용자가 임의 변경한 경우 모두 방어
    //   매칭 안 되면 (개인일정 등) 드라이버가 입력한 값 사용
    if (sheetName === 'Daily_Report') {
      try {
        const tcRaw = String(data.Tour_Code || data.TourCode || '').trim();
        if (tcRaw) {
          const schSheet = ss.getSheetByName('Schedule');
          if (schSheet) {
            const sLastRow = schSheet.getLastRow();
            if (sLastRow > 1) {
              const sHeaders = schSheet.getRange(1, 1, 1, schSheet.getLastColumn()).getValues()[0];
              const tcIdx = sHeaders.indexOf('TourCode');
              const tidIdx = sHeaders.indexOf('TourID');
              const beIdx = sHeaders.indexOf('BillingEntity');
              if ((tcIdx >= 0 || tidIdx >= 0) && beIdx >= 0) {
                const sData = schSheet.getRange(2, 1, sLastRow - 1, sHeaders.length).getValues();
                const tcU = tcRaw.toUpperCase();
                let scheduleBE = null;
                for (let i = 0; i < sData.length; i++) {
                  const r1 = tcIdx >= 0 ? String(sData[i][tcIdx]||'').trim().toUpperCase() : '';
                  const r2 = tidIdx >= 0 ? String(sData[i][tidIdx]||'').trim().toUpperCase() : '';
                  if (r1 === tcU || r2 === tcU) {
                    scheduleBE = String(sData[i][beIdx] || '').trim() || 'DC';
                    break;
                  }
                }
                if (scheduleBE !== null) {
                  const submittedBE = String(data.Billing_Entity || data.BillingEntity || '').trim();
                  if (submittedBE && submittedBE.toUpperCase() !== scheduleBE.toUpperCase()) {
                    Logger.log('[saveReport] BE override for TourCode ' + tcRaw +
                              ': submitted="' + submittedBE + '" → schedule="' + scheduleBE + '" (driver=' + (data.Driver||'') + ')');
                  }
                  data.Billing_Entity = scheduleBE;
                }
              }
            }
          }
        }
      } catch(beErr) {
        Logger.log('[saveReport] BE enforcement error (continuing with submitted value): ' + beErr);
      }
    }

    // ★ 실제 시트 헤더를 읽어서 매핑 (컬럼 순서 불일치 방지)
    const lastCol = sheet.getLastColumn();
    const actualHeaders = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : headers;
    const row = actualHeaders.map(h => data[h] !== undefined ? data[h] : '');
    sheet.appendRow(row);

    // ★ Daily_Report 저장 시 트레일러 사용료 자동 정산 (Sub_Txn 생성)
    //   조건: Trailer_Number 있고, 차량/트레일러 소유주 다름
    if (sheetName === 'Daily_Report') {
      try {
        _autoCreateTrailerRentalTxn(data);
      } catch(e) {
        Logger.log('[trailer rental] auto-txn error: ' + e);
      }
      // ★ Daily_Report 저장 시 인보이스 드래프트 항목 자동 누적
      //   투어코드별 드래프트(Manual Items)에 항목 추가 (이미 같은 항목 있으면 업데이트)
      try {
        _autoAddInvoiceDraftItem(data);
      } catch(e) {
        Logger.log('[invoice draft] auto-add error: ' + e);
      }
    }

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

    // ★ Daily_Report 수정 시: 기존 트레일러 거래를 먼저 가져와서 (수정 후 변경 감지용)
    let oldData = null;
    if (sheetName === 'Daily_Report') {
      try {
        const lastCol0 = sheet.getLastColumn();
        const oldHeaders = sheet.getRange(1, 1, 1, lastCol0).getValues()[0];
        const oldRow = sheet.getRange(ri, 1, 1, lastCol0).getValues()[0];
        oldData = {};
        oldHeaders.forEach((h, i) => oldData[h] = oldRow[i]);
      } catch(e) { Logger.log('[trailer rental] read old: ' + e); }
    }

    // ★ 실제 시트 헤더를 읽어서 매핑 (컬럼 순서 불일치 방지)
    const lastCol = sheet.getLastColumn();
    const actualHeaders = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : headers;
    const row = actualHeaders.map(h => data[h] !== undefined ? data[h] : '');
    sheet.getRange(ri, 1, 1, row.length).setValues([row]);

    // ★ Daily_Report 수정 시 트레일러 거래 동기화
    //   기존 거래 삭제 → 새 데이터로 재생성
    if (sheetName === 'Daily_Report') {
      try {
        if (oldData) _deleteTrailerRentalTxn(oldData);
        _autoCreateTrailerRentalTxn(data);
      } catch(e) { Logger.log('[trailer rental] sync on update: ' + e); }
      // ★ 인보이스 드래프트 항목 동기화 — 옛 항목 삭제 → 새 항목 추가
      try {
        if (oldData) _deleteInvoiceDraftItem(oldData);
        _autoAddInvoiceDraftItem(data);
      } catch(e) { Logger.log('[invoice draft] sync on update: ' + e); }
    }

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

    // ★ Daily_Report 삭제 시: 삭제 전 데이터를 먼저 가져와서 트레일러 거래도 같이 삭제
    let oldData = null;
    if (sheetName === 'Daily_Report') {
      try {
        const lastCol = sheet.getLastColumn();
        const oldHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
        const oldRow = sheet.getRange(ri, 1, 1, lastCol).getValues()[0];
        oldData = {};
        oldHeaders.forEach((h, i) => oldData[h] = oldRow[i]);
      } catch(e) { Logger.log('[trailer rental] read before delete: ' + e); }
    }

    sheet.deleteRow(ri);

    // ★ Daily_Report 삭제 후 트레일러 자동 거래도 삭제
    if (sheetName === 'Daily_Report' && oldData) {
      try { _deleteTrailerRentalTxn(oldData); } catch(e) { Logger.log('[trailer rental] sync on delete: ' + e); }
      // ★ 인보이스 드래프트 항목도 삭제
      try { _deleteInvoiceDraftItem(oldData); } catch(e) { Logger.log('[invoice draft] sync on delete: ' + e); }
    }

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

    // ── M_Drivers의 PIN은 해시화 (이미 해시면 그대로) ──
    if (sheetName === 'M_Drivers') {
      const pinColIdx = headers.indexOf('PIN');
      const krIdx = headers.indexOf('Name_KR');
      const enIdx = headers.indexOf('Name_EN');
      if (pinColIdx >= 0 && row[pinColIdx]) {
        const rawPin = String(row[pinColIdx] || '').trim();
        if (rawPin && rawPin.indexOf(PIN_HASH_PREFIX) !== 0) {
          const verifyName = String(row[krIdx] || row[enIdx] || '').trim();
          row[pinColIdx] = _hashPin(rawPin, verifyName);
        }
      }
    }

    // ★ Date/FinishDate 컬럼은 셀이 "Automatic" 포맷이면 YYYY-MM-DD 문자열을 Date 객체로 자동 변환해버림.
    //   이걸 막기 위해 (1) 값을 명시적으로 문자열로 변환, (2) 추가 후 해당 셀을 plain text 포맷으로 강제.
    //   영향 시트: SUB_Txn, Agency_Txn (Date 컬럼 사용하는 거래 시트)
    const _DATE_COLS_TO_PROTECT = ['Date', 'FinishDate'];
    const _dateColIdxs = [];
    if (sheetName === 'SUB_Txn' || sheetName === 'Agency_Txn') {
      _DATE_COLS_TO_PROTECT.forEach(colName => {
        const idx = headers.indexOf(colName);
        if (idx >= 0) {
          _dateColIdxs.push(idx);
          // 값이 YYYY-MM-DD 형식 문자열이면 그대로 두되, Date 객체로 들어왔으면 문자열로 변환
          const v = row[idx];
          if (v instanceof Date) {
            // Date 객체를 시드니 로컬 YYYY-MM-DD로
            row[idx] = Utilities.formatDate(v, 'Australia/Sydney', 'yyyy-MM-dd');
          } else if (v && typeof v === 'string') {
            // ISO 타임스탬프 형식(2026-05-11T14:00:00.000Z)이면 시드니 날짜로 정규화
            const m = v.match(/^(\d{4}-\d{2}-\d{2})T/);
            if (m) {
              const d = new Date(v);
              if (!isNaN(d.getTime())) {
                row[idx] = Utilities.formatDate(d, 'Australia/Sydney', 'yyyy-MM-dd');
              }
            }
            // 이미 YYYY-MM-DD면 그대로 두기 (변경 없음)
          }
        }
      });
    }

    sheet.appendRow(row);
    const newRowNum = sheet.getLastRow();

    // ★ 새로 추가된 행의 Date 컬럼을 plain text 포맷으로 강제 (다음번에 읽을 때 Date 객체로 변환 안 됨)
    if (_dateColIdxs.length > 0) {
      _dateColIdxs.forEach(idx => {
        try {
          sheet.getRange(newRowNum, idx + 1).setNumberFormat('@');
        } catch(fmtErr) {
          // 포맷 설정 실패는 치명적이지 않음 — 로그만
          Logger.log('[addMasterRow] setNumberFormat failed: ' + fmtErr);
        }
      });
    }

    return {ok: true, row: newRowNum};
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
      const sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const agencyCol = sheetHeaders.indexOf('Agency');
      const courseCol = sheetHeaders.indexOf('Course');
      if (agencyCol >= 0 && courseCol >= 0) {
        const data = sheet.getRange(2, 1, lastRow - 1, sheetHeaders.length).getValues();
        data.forEach(r => {
          if (r[agencyCol] && r[courseCol]) existing.add(r[agencyCol] + '||' + r[courseCol]);
        });
      }
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

// ── 필드 별칭 맵: 시트 헤더 ↔ 코드 키 불일치 자동 해소 ──
const FIELD_ALIASES = {
  'phone': ['mobile_1', 'mobile'],
  'mobile_1': ['phone', 'mobile'],
  'mobile': ['phone', 'mobile_1'],
  'license_#': ['license_no'],
  'license_no': ['license_#'],
  'authority_#': ['authority_no'],
  'authority_no': ['authority_#'],
  'next_of_kin': ['next of kin'],
  'engine_number': ['engine number'],
  'manufacture_date': ['manufacture date']
};

// data 객체를 정규화 키로 조회하는 맵 생성 (별칭 포함)
function buildNormMap(data) {
  const m = {};
  Object.keys(data).forEach(k => {
    const nk = normalizeKey(k);
    m[nk] = data[k];
    // 별칭도 등록 (이미 있는 키는 덮어쓰지 않음)
    const aliases = FIELD_ALIASES[nk];
    if (aliases) {
      aliases.forEach(a => { if (m[a] === undefined) m[a] = data[k]; });
    }
  });
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
    var PHONE_COL_NAMES = ['phone','mobile','mobile_1','mobile_2','moblie_2'];

    // ── M_Drivers 업데이트 시 기존 PIN 보존을 위한 사전 조회 ──
    // _stripPinFromDrivers로 클라이언트에서 PIN이 빠진 상태로 오기 때문에,
    // payload에 PIN이 없거나 빈 값이면 기존 PIN을 유지해야 함.
    let existingPinValue = null;
    if (sheetName === 'M_Drivers') {
      const pinColIdx = headers.indexOf('PIN');
      if (pinColIdx >= 0) {
        try {
          existingPinValue = sheet.getRange(ri, pinColIdx + 1).getValue();
        } catch(e){ existingPinValue = null; }
      }
    }

    const row = headers.map((h, i) => {
      let val;
      if (data[h] !== undefined) val = data[h];
      else {
        const nk = normalizeKey(h);
        val = normMap[nk] !== undefined ? normMap[nk] : '';
      }
      // ★ 전화번호 필드: 앞 0 복원 + 텍스트 서식
      if (PHONE_COL_NAMES.includes(normalizeKey(h)) && val !== '' && val !== null && val !== undefined) {
        let s = String(val).replace(/\.0+$/, '').replace(/[^0-9]/g, '');
        if (s.length === 9) s = '0' + s;
        val = s;
      }
      // ★ M_Drivers의 PIN 컬럼: 빈 값이면 기존 값 보존, 평문이면 해시화
      if (sheetName === 'M_Drivers' && h === 'PIN') {
        const incoming = String(val || '').trim();
        if (!incoming || incoming === '••••' || incoming === '****') {
          // payload에 PIN이 없거나 마스킹 → 기존 값 유지
          val = existingPinValue !== null ? existingPinValue : '';
        } else if (incoming.indexOf(PIN_HASH_PREFIX) !== 0) {
          // 평문 PIN이면 해시화 (4자리 이상 숫자 검증)
          if (/^\d{4,}$/.test(incoming)) {
            const krIdx = headers.indexOf('Name_KR');
            const enIdx = headers.indexOf('Name_EN');
            const verifyName = String(row[krIdx] !== undefined ? data[headers[krIdx]] || normMap[normalizeKey(headers[krIdx])] : '') ||
                               String(data[headers[enIdx]] || '') || '';
            // verifyName이 빈 경우 시트에서 가져옴
            let nameForHash = verifyName;
            if (!nameForHash) {
              try {
                const krVal = krIdx >= 0 ? sheet.getRange(ri, krIdx + 1).getValue() : '';
                const enVal = enIdx >= 0 ? sheet.getRange(ri, enIdx + 1).getValue() : '';
                nameForHash = String(krVal || enVal || '').trim();
              } catch(e){}
            }
            val = _hashPin(incoming, nameForHash);
          } else {
            // 형식 불량 → 기존 값 유지 (안전 우선)
            val = existingPinValue !== null ? existingPinValue : '';
          }
        }
        // 이미 해시면 그대로 사용
      }
      return val;
    });
    sheet.getRange(ri, 1, 1, row.length).setValues([row]);
    // ★ 전화번호 셀에 텍스트 서식 적용 (앞 0 보존)
    headers.forEach((h, i) => {
      if (PHONE_COL_NAMES.includes(normalizeKey(h))) {
        sheet.getRange(ri, i + 1).setNumberFormat('@');
      }
    });

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

    // ── M_Drivers 일괄 교체 시 기존 PIN 백업 (이름 → PIN 맵) ──
    // 클라이언트가 _stripPinFromDrivers로 PIN 없이 보내기 때문에, 빈 값이 와도
    // 원래 PIN을 보존해야 한다.
    let pinBackup = null;
    if (sheetName === 'M_Drivers') {
      pinBackup = {};
      try {
        const lastR = sheet.getLastRow();
        const lastC = sheet.getLastColumn();
        if (lastR > 1 && lastC > 0) {
          const existingHeaders = sheet.getRange(1, 1, 1, lastC).getValues()[0].map(String);
          const krIdx = existingHeaders.indexOf('Name_KR');
          const enIdx = existingHeaders.indexOf('Name_EN');
          const pinIdx = existingHeaders.indexOf('PIN');
          if (pinIdx >= 0 && (krIdx >= 0 || enIdx >= 0)) {
            const existing = sheet.getRange(2, 1, lastR - 1, lastC).getValues();
            existing.forEach(r => {
              const kr = krIdx >= 0 ? String(r[krIdx] || '').trim() : '';
              const en = enIdx >= 0 ? String(r[enIdx] || '').trim() : '';
              const pin = String(r[pinIdx] || '').trim();
              if (pin) {
                if (kr) pinBackup[kr] = pin;
                if (en) pinBackup[en] = pin;
              }
            });
          }
        }
      } catch(e){ pinBackup = {}; }
    }

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

    if (rows && rows.length > 0) {
      const data = rows.map(obj => headers.map(h => {
        let val = obj[h] !== undefined ? obj[h] : '';
        // ── M_Drivers PIN 복원/해시 처리 ──
        if (sheetName === 'M_Drivers' && h === 'PIN') {
          const incoming = String(val || '').trim();
          if (!incoming || incoming === '••••' || incoming === '****') {
            // 비어있으면 백업에서 복원
            const kr = String(obj['Name_KR'] || '').trim();
            const en = String(obj['Name_EN'] || '').trim();
            val = (pinBackup && (pinBackup[kr] || pinBackup[en])) || '';
          } else if (incoming.indexOf(PIN_HASH_PREFIX) !== 0) {
            // 평문이면 해시화 (4자리 이상 숫자만)
            if (/^\d{4,}$/.test(incoming)) {
              const verifyName = String(obj['Name_KR'] || obj['Name_EN'] || '').trim();
              val = _hashPin(incoming, verifyName);
            } else {
              // 형식 불량 → 백업 복원
              const kr = String(obj['Name_KR'] || '').trim();
              const en = String(obj['Name_EN'] || '').trim();
              val = (pinBackup && (pinBackup[kr] || pinBackup[en])) || '';
            }
          }
          // 이미 해시면 그대로
        }
        return val;
      }));
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

    const driver = String(data.Driver || '').trim();
    const weekStart = String(data.WeekStart || '').trim();
    const date = String(data.Date || '').trim();
    const amount = parseFloat(data.Amount) || 0;

    // ★ 중복 클릭 방어: 같은 (Driver, WeekStart, Date, Amount)가 최근 10초 내에 추가됐으면 기존 row 반환
    //   증상: 사용자 더블클릭 또는 비동기 race로 GAS에 동일 row가 2건 등록되는 버그
    const headers = MASTER_HEADERS['Wages'];
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const data2 = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
      const idIdx = headers.indexOf('RowID');
      const drvIdx = headers.indexOf('Driver');
      const wsIdx = headers.indexOf('WeekStart');
      const dtIdx = headers.indexOf('Date');
      const amtIdx = headers.indexOf('Amount');
      const now = Date.now();
      for (let i = data2.length - 1; i >= 0; i--) {  // 최근 row부터 역순 검사
        const r = data2[i];
        const rid = parseInt(r[idIdx]) || 0;
        if (rid > 0 && (now - rid) > 10000) break;  // 10초 넘은 row까지만 본 후 break (이전 row는 더 오래됨)
        const rDrv = String(r[drvIdx]||'').trim();
        const rWs = String(r[wsIdx]||'').trim();
        const rDt = String(r[dtIdx]||'').trim();
        const rAmt = parseFloat(r[amtIdx]) || 0;
        if (rDrv === driver && rWs === weekStart && rDt === date && Math.abs(rAmt - amount) < 0.01) {
          Logger.log('[addWage] duplicate detected, skip insert: ' + driver + ' ' + date + ' $' + amount);
          return {ok: true, row: i + 2, rowId: String(rid), duplicate: true};
        }
      }
    }

    const rowId = Date.now().toString();
    const newRow = [
      rowId,
      driver,
      weekStart,
      date,
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

    const headers = MASTER_HEADERS['Wages'];
    const amount = parseFloat(data.Amount) || 0;
    const rowData = headers.map(h => {
      if (h === 'RowID') return data.RowID || Date.now().toString();
      if (h === 'Driver') return data.Driver || '';
      if (h === 'WeekStart') return data.WeekStart || '';
      if (h === 'Date') return data.Date || '';
      if (h === 'Amount') return amount;
      if (h === 'PayMethod') return data.PayMethod || 'Cash';
      if (h === 'Notes') return data.Notes || '';
      return data[h] !== undefined ? data[h] : '';
    });

    sheet.getRange(ri, 1, 1, headers.length).setValues([rowData]);

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
      const headers = MASTER_HEADERS['Wages'];
      const newData = rows.map(r => headers.map(h => {
        if (h === 'RowID') return r.RowID || Date.now().toString();
        if (h === 'Driver') return r.Driver || '';
        if (h === 'WeekStart') return r.WeekStart || '';
        if (h === 'Date') return r.Date || '';
        if (h === 'Amount') return parseFloat(r.Amount) || 0;
        if (h === 'PayMethod') return r.PayMethod || 'Cash';
        if (h === 'Notes') return r.Notes || '';
        return r[h] !== undefined ? r[h] : '';
      }));
      sheet.getRange(2, 1, newData.length, headers.length).setValues(newData);
    }

    return {ok: true, count: rows ? rows.length : 0};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

/**
 * 일회성 정리 — Wages 시트의 중복 row 제거
 *
 * 중복 판단 기준: 같은 (Driver, WeekStart, Date, Amount) 조합
 *   - PayMethod / Notes는 다를 수 있어도 중복으로 봄 (사용자가 같은 지급을 두 번 입력했을 가능성)
 *   - 첫 row(가장 오래된 RowID)는 유지, 나머지 삭제
 *
 * 사용법 (Apps Script 편집기에서 1회 실행):
 *   cleanupDuplicateWages()  → 미리보기 (삭제 안 함)
 *   cleanupDuplicateWages(true)  → 실제 삭제
 */
function cleanupDuplicateWages(execute) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Wages');
  if (!sheet) return {ok: false, msg: 'Wages sheet not found'};

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return {ok: true, msg: 'No data', duplicates: 0};

  const headers = MASTER_HEADERS['Wages'];
  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  const idIdx = headers.indexOf('RowID');
  const drvIdx = headers.indexOf('Driver');
  const wsIdx = headers.indexOf('WeekStart');
  const dtIdx = headers.indexOf('Date');
  const amtIdx = headers.indexOf('Amount');

  // 키별로 첫 번째 row만 유지, 나머지는 삭제 대상
  const seen = new Map();  // key → first row index (1-indexed sheet row)
  const toDelete = [];     // [{rowIndex, driver, date, amount, rowId}]

  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const sheetRow = i + 2;  // 1-indexed, header is row 1
    const drv = String(r[drvIdx]||'').trim();
    const ws = String(r[wsIdx]||'').trim();
    const dt = String(r[dtIdx]||'').trim();
    const amt = parseFloat(r[amtIdx]) || 0;
    if (!drv || !dt) continue;  // 빈 row skip
    const key = `${drv}|${ws}|${dt}|${amt.toFixed(2)}`;
    if (seen.has(key)) {
      toDelete.push({
        sheetRow: sheetRow,
        driver: drv,
        date: dt,
        amount: amt,
        rowId: String(r[idIdx]||''),
        firstAtRow: seen.get(key)
      });
    } else {
      seen.set(key, sheetRow);
    }
  }

  Logger.log('[cleanupDuplicateWages] Found ' + toDelete.length + ' duplicates out of ' + data.length + ' rows');
  toDelete.forEach(d => Logger.log('  - row ' + d.sheetRow + ': ' + d.driver + ' ' + d.date + ' $' + d.amount + ' (first at row ' + d.firstAtRow + ')'));

  if (!execute) {
    return {ok: true, msg: 'PREVIEW MODE (no deletion). Call cleanupDuplicateWages(true) to execute.', duplicates: toDelete.length, details: toDelete};
  }

  // 실제 삭제 — 큰 row index부터 (작은 것부터 지우면 index가 밀림)
  toDelete.sort((a, b) => b.sheetRow - a.sheetRow);
  toDelete.forEach(d => sheet.deleteRow(d.sheetRow));

  Logger.log('[cleanupDuplicateWages] Deleted ' + toDelete.length + ' duplicate rows');
  return {ok: true, msg: 'Deleted ' + toDelete.length + ' duplicates', duplicates: toDelete.length};
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

    // 입력된 PIN 검증
    const pinStr = String(pin || '').trim();
    if (!pinStr || pinStr.length < 4 || !/^\d+$/.test(pinStr)) {
      return {ok: false, msg: 'PIN은 4자리 이상의 숫자여야 합니다'};
    }

    for (let r = 1; r < data.length; r++) {
      if (data[r][nameENIdx] === driverName || data[r][nameKRIdx] === driverName) {
        // 시트의 KR 이름으로 해시 (로그인 시와 일관)
        const verifyName = String(data[r][nameKRIdx] || data[r][nameENIdx] || '').trim();
        const hashed = _hashPin(pinStr, verifyName);
        sheet.getRange(r + 1, pinIdx + 1).setValue(hashed);
        // PIN 변경 시 해당 사용자의 실패 카운트도 클리어
        try { _clearAuthFails(driverName); } catch(e){}
        return {ok: true};
      }
    }

    return {ok: false, msg: 'Driver not found: ' + driverName};
  } catch (err) {
    return {ok: false, error: err.toString()};
  }
}

// ── 날짜 입력 정규화 (서버 측 방어선) → 'dd/mm/yyyy' 또는 '' ─────────
function _normalizeDateForSheet(raw) {
  if (raw === null || raw === undefined) return '';
  var s = String(raw).trim();
  if (!s) return '';
  // 이미 dd/mm/yyyy
  var m1 = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m1) {
    var d1 = +m1[1], mo1 = +m1[2], y1 = +m1[3];
    if (_validDMY_(d1, mo1, y1)) return s;
    return '';
  }
  // ISO yyyy-mm-dd / yyyy/mm/dd
  var m2 = s.match(/^(\d{4})[-\/\.](\d{1,2})[-\/\.](\d{1,2})/);
  if (m2) {
    var y2 = +m2[1], mo2 = +m2[2], d2 = +m2[3];
    if (_validDMY_(d2, mo2, y2)) return _padDMY_(d2, mo2, y2);
    return '';
  }
  // dd-mm-yyyy / dd.mm.yyyy / dd mm yyyy
  var m3 = s.match(/^(\d{1,2})[-\/\.\s](\d{1,2})[-\/\.\s](\d{4})$/);
  if (m3) {
    var d3 = +m3[1], mo3 = +m3[2], y3 = +m3[3];
    if (_validDMY_(d3, mo3, y3)) return _padDMY_(d3, mo3, y3);
    return '';
  }
  // 숫자만 8자리 (ddmmyyyy)
  var digits = s.replace(/[^0-9]/g, '');
  if (digits.length === 8) {
    var d4 = +digits.slice(0,2), mo4 = +digits.slice(2,4), y4 = +digits.slice(4,8);
    if (_validDMY_(d4, mo4, y4)) return _padDMY_(d4, mo4, y4);
  }
  // 텍스트 월: "13 Jan 2027", "4-Jun-2026"
  var months = {jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12};
  var m5 = s.toLowerCase().match(/^(\d{1,2})[-\s\/](jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[-\s\/](\d{4})$/);
  if (m5) {
    var d5 = +m5[1], mo5 = months[m5[2]], y5 = +m5[3];
    if (_validDMY_(d5, mo5, y5)) return _padDMY_(d5, mo5, y5);
  }
  // Date 객체 (시트가 raw Date를 보낸 경우)
  if (Object.prototype.toString.call(raw) === '[object Date]' && !isNaN(raw)) {
    return _padDMY_(raw.getDate(), raw.getMonth()+1, raw.getFullYear());
  }
  return '';
}
function _validDMY_(d, m, y) {
  if (!d || !m || !y) return false;
  if (y < 1900 || y > 2100) return false;
  if (m < 1 || m > 12) return false;
  if (d < 1 || d > 31) return false;
  var dt = new Date(y, m-1, d);
  return dt.getFullYear() === y && dt.getMonth() === m-1 && dt.getDate() === d;
}
function _padDMY_(d, m, y) {
  return ('0'+d).slice(-2) + '/' + ('0'+m).slice(-2) + '/' + y;
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
      wwcNo: 'WWC_No', wwcExp: 'WWC_Expiry',
      nokName: 'NEXT_OF_KIN', nokPhone: 'Mobile_2',
      address: 'Address', suburb: 'Suburb',
      bank: 'Bank', bsb: 'BSB', account: 'Account'
    };

    // ★ 진단: 매핑된 컬럼이 실제 시트에 있는지 미리 검증 (저장 누락 디버깅용)
    const missingColumns = [];
    Object.values(fieldMap).forEach(col => {
      if (headers.indexOf(col) === -1) missingColumns.push(col);
    });
    if (missingColumns.length > 0) {
      Logger.log('[updateDriverInfo] 누락된 시트 컬럼: ' + missingColumns.join(', '));
    }

    for (let r = 1; r < sheetData.length; r++) {
      if (sheetData[r][nameENIdx] === driverName || sheetData[r][nameKRIdx] === driverName) {
        const PHONE_SAVE_FIELDS = ['Mobile_1', 'Mobile_2', 'Phone', 'Mobile'];
        const DATE_SAVE_FIELDS = ['License_Expiry', 'Authority_Expiry', 'WWC_Expiry'];
        const savedFields = [];
        const skippedFields = [];
        Object.entries(data).forEach(([key, val]) => {
          const col = fieldMap[key];
          if (col) {
            const colIdx = headers.indexOf(col);
            if (colIdx !== -1) {
              const cell = sheet.getRange(r + 1, colIdx + 1);
              if (PHONE_SAVE_FIELDS.includes(col)) {
                let s = String(val||'').replace(/[^0-9]/g, '');
                if (s.length === 9) s = '0' + s;
                cell.setNumberFormat('@').setValue(s);
              } else if (DATE_SAVE_FIELDS.includes(col)) {
                const norm = _normalizeDateForSheet(val);
                cell.setNumberFormat('@').setValue(norm);
              } else {
                cell.setValue(val);
              }
              savedFields.push(key + '→' + col);
            } else {
              skippedFields.push(key + '→' + col + ' (시트에 컬럼 없음)');
            }
          } else if (key !== 'savedAt' && !key.startsWith('photoUrl_')) {
            skippedFields.push(key + ' (매핑 없음)');
          }
        });
        return {
          ok: true,
          saved: savedFields,
          skipped: skippedFields,
          missingColumns: missingColumns
        };
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
      const headers = MASTER_HEADERS['Notices'];
      const newData = rows.map(r => headers.map(h => {
        if (h === 'ID') return r.id || r.ID || String(Date.now());
        if (h === 'Title') return r.title || r.Title || '';
        if (h === 'Content') return r.content || r.Content || '';
        if (h === 'Type') return r.type || r.Type || 'info';
        if (h === 'Date') return r.date || r.Date || '';
        if (h === 'Active') return (r.active === false || r.Active === 'false') ? 'false' : 'true';
        return r[h] !== undefined ? r[h] : '';
      }));
      sheet.getRange(2, 1, newData.length, headers.length).setValues(newData);
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

/**
 * 감사 로그 조회 (관리자 전용)
 * 최신순으로 limit건 반환
 */
function getAuditLog(limit) {
  try {
    limit = limit || 200;
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const logSheet = ss.getSheetByName('Audit_Log');
    if (!logSheet) return { ok: true, rows: [], total: 0 };
    const data = logSheet.getDataRange().getValues();
    if (data.length < 2) return { ok: true, rows: [], total: 0 };
    const headers = data[0];
    const rows = [];
    // 최신순 (마지막부터)
    for (let i = data.length - 1; i >= 1 && rows.length < limit; i--) {
      const row = {};
      headers.forEach((h, j) => { row[h] = data[i][j]; });
      rows.push(row);
    }
    return { ok: true, rows: rows, total: data.length - 1 };
  } catch (err) {
    return { ok: false, error: err.toString() };
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

    // 기존 행 찾기 (InvNumber 기준 — 헤더명으로 컬럼 위치 조회)
    const allData = sheet.getDataRange().getValues();
    const sheetHeaders = allData[0];
    const invNumCol = sheetHeaders.indexOf('InvNumber');
    if (invNumCol < 0) return { ok: false, error: 'InvNumber column not found in Invoices sheet' };
    let existingRow = -1;
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][invNumCol]).trim() === invNum) { existingRow = i + 1; break; }
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
  // getInvoices는 Invoices + Agency_Txn 합성. 캐시 키 'Invoices'로 통일.
  // Agency_Txn 변경 시도 'Invoices' 캐시 함께 무효화 (saveInvoice/addAgencyTxn 등에서 처리)
  return _cachedRead('Invoices', function() { return _getInvoicesImpl(); });
}

function _getInvoicesImpl() {
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

    // ★ Agency_Txn에서 인보이스별 CR 합계 계산 → PaidCR 필드 추가
    try {
      const txnSheet = ss.getSheetByName('Agency_Txn');
      if (txnSheet && txnSheet.getLastRow() > 1) {
        const txnData = txnSheet.getDataRange().getValues();
        const txnHeaders = txnData[0];
        const invIdCol = txnHeaders.indexOf('InvoiceID');
        const crCol = txnHeaders.indexOf('CR');
        if (invIdCol >= 0 && crCol >= 0) {
          const crMap = {};
          for (let i = 1; i < txnData.length; i++) {
            const invId = String(txnData[i][invIdCol] || '').trim();
            const cr = Number(txnData[i][crCol]) || 0;
            if (invId && cr > 0) {
              crMap[invId] = (crMap[invId] || 0) + cr;
            }
          }
          rows.forEach(inv => {
            const invNum = String(inv.InvNumber || '').trim();
            inv.PaidCR = Math.round((crMap[invNum] || 0) * 100) / 100;
          });
        }
      }
    } catch (e) {
      // PaidCR 계산 실패해도 인보이스 데이터는 정상 반환
      Logger.log('PaidCR calculation error: ' + e.toString());
    }

    return { ok: true, rows };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

/**
 * 거래처별 선수금/크레딧 잔액 조회
 * Agency_Txn의 Type 필드 기반:
 *   prepaid_in / prepaid_use → 선수금 잔액
 *   credit_in / credit_use → 크레딧 잔액
 * agency 파라미터가 있으면 해당 거래처만, 없으면 전체
 */
function getAgencyBalances(agency) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Agency_Txn');
    if (!sheet || sheet.getLastRow() <= 1) return { ok: true, balances: {} };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const agCol = headers.indexOf('Agency');
    const typeCol = headers.indexOf('Type');
    const crCol = headers.indexOf('CR');
    const drCol = headers.indexOf('DR');

    if (agCol < 0 || typeCol < 0) return { ok: true, balances: {} };

    // { agency: { prepaid: {in, used, balance}, credit: {in, used, balance} } }
    const map = {};
    for (let i = 1; i < data.length; i++) {
      const ag = String(data[i][agCol] || '').trim();
      if (!ag) continue;
      if (agency && ag !== agency) continue;
      const type = String(data[i][typeCol] || '').trim();
      const cr = Number(data[i][crCol]) || 0;
      const dr = Number(data[i][drCol]) || 0;

      if (!map[ag]) map[ag] = { prepaid: { in: 0, used: 0 }, credit: { in: 0, used: 0 } };

      if (type === 'prepaid_in')  map[ag].prepaid.in += cr;
      if (type === 'prepaid_use') map[ag].prepaid.used += cr;
      if (type === 'credit_in')   map[ag].credit.in += cr;
      if (type === 'credit_use')  map[ag].credit.used += cr;
    }

    // 잔액 계산
    Object.values(map).forEach(v => {
      v.prepaid.balance = Math.round((v.prepaid.in - v.prepaid.used) * 100) / 100;
      v.credit.balance  = Math.round((v.credit.in - v.credit.used) * 100) / 100;
    });

    return { ok: true, balances: map };
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
    const sheetHeaders = data[0];
    const invNumCol = sheetHeaders.indexOf('InvNumber');
    if (invNumCol < 0) return { ok: false, error: 'InvNumber column not found in Invoices sheet' };
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][invNumCol]).trim() === invNumber) { targetRow = i + 1; break; }
    }
    if (targetRow < 0) return { ok: false, error: 'Invoice not found: ' + invNumber };

    const now = Utilities.formatDate(new Date(), 'Australia/Sydney', 'dd/MM/yyyy HH:mm:ss');

    // Status 열 업데이트 (시트 헤더 기준)
    const statusCol = sheetHeaders.indexOf('Status') + 1;
    if (statusCol > 0) sheet.getRange(targetRow, statusCol).setValue(status);

    // 날짜 필드 업데이트
    if (field) {
      const fieldCol = sheetHeaders.indexOf(field) + 1;
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
    const sheetHeaders = data[0];
    const invNumCol = sheetHeaders.indexOf('InvNumber');
    if (invNumCol < 0) return { ok: false, error: 'InvNumber column not found' };

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][invNumCol]).trim() === invNumber) {
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
// INVOICE EMAIL (GmailApp — PDF 첨부)
// ═══════════════════════════════════════════════════════════════════════════

/**
 * 인보이스 이메일 발송 (PDF 첨부)
 * payload: { to, subject, body, cc, pdfBase64, pdfName, senderName, replyTo, _user }
 *   pdfBase64: 클라이언트에서 생성한 PDF의 base64 문자열
 */
function sendInvoiceEmail(payload) {
  try {
    // 다중 이메일 정규화: 콤마/세미콜론/공백/줄바꿈으로 구분된 여러 주소 → "a@x.com, b@y.com" 형식
    function _normEmails(s){
      if(!s) return '';
      return String(s).split(/[,;\s\n\r]+/)
        .map(e => e.trim())
        .filter(e => e && e.indexOf('@') !== -1)
        .join(', ');
    }
    const to        = _normEmails(payload.to);
    const subject   = (payload.subject || '').trim();
    const body      = (payload.body || '').trim();
    const cc        = _normEmails(payload.cc);
    const name      = payload.senderName || 'Dong Choi Pty Ltd';
    const replyTo   = (payload.replyTo || '').trim();
    const pdfBase64 = payload.pdfBase64 || '';
    const pdfName   = payload.pdfName || 'Invoice.pdf';
    const docHtml   = payload.docHtml || '';

    if (!to)      return { ok: false, error: '수신자 이메일이 없습니다 (to is empty)' };
    if (!subject) return { ok: false, error: '제목이 없습니다 (subject is empty)' };

    const options = { name: name };
    if (cc) options.cc = cc;
    if (replyTo) options.replyTo = replyTo;

    // ★ PDF 첨부: docHtml 우선 (서버사이드 변환), base64는 폴백
    //   pdfAttached 플래그를 추적해서 클라이언트 토스트가 정확히 표시되도록 한다
    let pdfAttached = false;
    let pdfError = '';
    if (docHtml) {
      try {
        var htmlBlob = Utilities.newBlob(docHtml, 'text/html', 'invoice.html');
        var pdfBlob  = htmlBlob.getAs('application/pdf').setName(pdfName);
        options.attachments = [pdfBlob];
        pdfAttached = true;
      } catch (pdfErr) {
        pdfError = 'HTML→PDF 변환 실패: ' + pdfErr;
      }
    } else if (pdfBase64) {
      try {
        var pdfBytes = Utilities.base64Decode(pdfBase64);
        var pdfBlob2 = Utilities.newBlob(pdfBytes, 'application/pdf', pdfName);
        options.attachments = [pdfBlob2];
        pdfAttached = true;
      } catch (pdfErr2) {
        pdfError = 'base64 디코딩 실패: ' + pdfErr2;
      }
    } else {
      pdfError = 'PDF 데이터 없음 (docHtml/pdfBase64 둘 다 비어있음)';
    }

    // GmailApp 우선 시도, 실패 시 MailApp 폴백
    try {
      GmailApp.sendEmail(to, subject, body, options);
    } catch (gmailErr) {
      var mailOpts = {
        to: to,
        subject: subject,
        body: body,
        name: name,
        attachments: options.attachments || []
      };
      if (cc) mailOpts.cc = cc;
      if (replyTo) mailOpts.replyTo = replyTo;
      MailApp.sendEmail(mailOpts);
    }

    // 감사 로그
    appendAuditLog(payload._user, 'send_invoice_email', '—', '—',
      `인보이스 이메일 발송 ${pdfAttached?'(PDF 첨부 ✅)':'(PDF 첨부 실패: '+pdfError+')'} → ${to} | ${subject}`);

    return { ok: true, to: to, pdfAttached: pdfAttached, pdfError: pdfError };
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

// ═══════════════════════════════════════════════════════════════════════════
// Defect Reports — Google Sheets 동기화
// ═══════════════════════════════════════════════════════════════════════════

function getDefectReports(driverName) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'Defect_Reports');
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { ok: true, reports: [] };
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const rows = data.map((row, idx) => {
      const obj = { _rowIndex: idx + 2 };
      headers.forEach((h, ci) => { obj[h] = row[ci]; });
      return obj;
    });
    // 드라이버 필터 (빈 문자열이면 전체)
    const filtered = driverName
      ? rows.filter(r => String(r.Driver || '') === driverName)
      : rows;
    return { ok: true, reports: filtered };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

function saveDefectReport(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'Defect_Reports');
    const headers = MASTER_HEADERS['Defect_Reports'];
    const row = headers.map(h => {
      const nk = normalizeKey(h);
      // data의 키를 lowercase로 매칭
      for (const k of Object.keys(data)) {
        if (normalizeKey(k) === nk) return data[k] || '';
      }
      return '';
    });
    sheet.appendRow(row);
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

function updateDefectStatus(id, status, adminNote) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'Defect_Reports');
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { ok: false, error: 'No data' };
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    const idCol = headers.indexOf('ID');
    const statusCol = headers.indexOf('Status');
    const noteCol = headers.indexOf('AdminNote');
    if (idCol < 0) return { ok: false, error: 'ID column not found' };
    const data = sheet.getRange(2, idCol + 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        if (statusCol >= 0) sheet.getRange(i + 2, statusCol + 1).setValue(status || '');
        if (noteCol >= 0 && adminNote !== undefined) sheet.getRange(i + 2, noteCol + 1).setValue(adminNote || '');
        return { ok: true };
      }
    }
    return { ok: false, error: 'ID not found: ' + id };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Bus Damage Markers — Google Sheets 동기화
// ═══════════════════════════════════════════════════════════════════════════

function getBusDamage(rego) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'Bus_Damage');
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { ok: true, markers: [], rego: rego };
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const regoCol = headers.indexOf('Rego');
    const markersCol = headers.indexOf('Markers');
    if (regoCol < 0) return { ok: true, markers: [], rego: rego };
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][regoCol]).trim() === String(rego).trim()) {
        let markers = [];
        try { markers = JSON.parse(data[i][markersCol] || '[]'); } catch(e) {}
        return { ok: true, markers: markers, rego: rego };
      }
    }
    return { ok: true, markers: [], rego: rego };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

function saveBusDamage(rego, markers, driver) {
  try {
    if (!rego) return { ok: false, error: 'Rego required' };
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'Bus_Damage');
    const lastRow = sheet.getLastRow();
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    const headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
    const regoCol = headers.indexOf('Rego');
    const markersCol = headers.indexOf('Markers');
    const updatedAtCol = headers.indexOf('UpdatedAt');
    const updatedByCol = headers.indexOf('UpdatedBy');
    const now = Utilities.formatDate(new Date(), 'Australia/Sydney', 'dd/MM/yyyy HH:mm:ss');
    const markersJson = JSON.stringify(markers || []);

    // 기존 행 찾기
    if (lastRow > 1 && regoCol >= 0) {
      const data = sheet.getRange(2, regoCol + 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim() === String(rego).trim()) {
          // 기존 행 업데이트
          if (markersCol >= 0) sheet.getRange(i + 2, markersCol + 1).setValue(markersJson);
          if (updatedAtCol >= 0) sheet.getRange(i + 2, updatedAtCol + 1).setValue(now);
          if (updatedByCol >= 0) sheet.getRange(i + 2, updatedByCol + 1).setValue(driver || '');
          return { ok: true, updated: true };
        }
      }
    }
    // 새 행 추가
    const expected = MASTER_HEADERS['Bus_Damage'];
    const row = expected.map(h => {
      if (h === 'Rego') return rego;
      if (h === 'Markers') return markersJson;
      if (h === 'UpdatedAt') return now;
      if (h === 'UpdatedBy') return driver || '';
      return '';
    });
    sheet.appendRow(row);
    return { ok: true, created: true };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// LEAVE REQUEST SYSTEM
// ═══════════════════════════════════════════════════════════════════════════

function ensureLeaveSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName('Leave_Requests');
  if (!sh) {
    sh = ss.insertSheet('Leave_Requests');
    const headers = [
      'Request_ID','Driver','Date_From','Date_To','Days','Reason',
      'Status','Requested_At','Reviewed_At','Reviewed_By','Admin_Note'
    ];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }
  return sh;
}

function submitLeaveRequest(data) {
  const sh = ensureLeaveSheet_();
  const syd = Utilities.formatDate(new Date(), 'Australia/Sydney', 'dd/MM/yyyy HH:mm');
  const from = parseDateFlexible_(data.Date_From);
  const to = parseDateFlexible_(data.Date_To);
  const days = Math.round((to - from) / (1000 * 60 * 60 * 24)) + 1;

  const existing = sh.getDataRange().getValues();
  const headers = existing[0];
  const driverIdx = headers.indexOf('Driver');
  const fromIdx = headers.indexOf('Date_From');
  const statusIdx = headers.indexOf('Status');
  for (let i = 1; i < existing.length; i++) {
    if (existing[i][driverIdx] === data.Driver &&
        existing[i][statusIdx] === 'Pending' &&
        existing[i][fromIdx] === data.Date_From) {
      return { ok: false, error: 'Duplicate pending request for this date (이미 같은 날짜에 대기 중인 요청이 있습니다)' };
    }
  }

  const requestId = 'LR_' + Date.now();
  sh.appendRow([
    requestId, data.Driver, data.Date_From, data.Date_To, days,
    data.Reason || '', 'Pending', syd, '', '', ''
  ]);
  return { ok: true, requestId: requestId, message: 'Leave request submitted (휴무 요청이 제출되었습니다)' };
}

function formatLeaveCell_(val, header) {
  if (!(val instanceof Date)) return val;
  const tz = 'Australia/Sydney';
  if (header === 'Date_From' || header === 'Date_To') {
    return Utilities.formatDate(val, tz, 'dd/MM/yyyy');
  }
  return Utilities.formatDate(val, tz, 'dd/MM/yyyy HH:mm');
}

function getMyLeaveRequests(driverName) {
  const sh = ensureLeaveSheet_();
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return { ok: true, requests: [] };
  const headers = data[0];
  const results = [];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = formatLeaveCell_(data[i][idx], h); });
    if (obj.Driver === driverName) results.push(obj);
  }
  results.reverse();
  return { ok: true, requests: results };
}

function getAllLeaveRequests(filter) {
  const sh = ensureLeaveSheet_();
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return { ok: true, requests: [] };
  const headers = data[0];
  const results = [];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = formatLeaveCell_(data[i][idx], h); });
    obj._row = i + 1;
    if (filter === 'all' || !filter || obj.Status === filter) results.push(obj);
  }
  results.sort((a, b) => (a.Status === 'Pending' ? -1 : 1) - (b.Status === 'Pending' ? -1 : 1));
  return { ok: true, requests: results };
}

function reviewLeaveRequest(data) {
  const sh = ensureLeaveSheet_();
  const allData = sh.getDataRange().getValues();
  const headers = allData[0];
  const idIdx = headers.indexOf('Request_ID');
  const statusIdx = headers.indexOf('Status');
  const reviewedAtIdx = headers.indexOf('Reviewed_At');
  const reviewedByIdx = headers.indexOf('Reviewed_By');
  const adminNoteIdx = headers.indexOf('Admin_Note');
  const syd = Utilities.formatDate(new Date(), 'Australia/Sydney', 'dd/MM/yyyy HH:mm');

  let targetRow = -1;
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][idIdx] === data.Request_ID) { targetRow = i + 1; break; }
  }
  if (targetRow === -1) return { ok: false, error: 'Request not found (요청을 찾을 수 없습니다)' };

  sh.getRange(targetRow, statusIdx + 1).setValue(data.Status);
  sh.getRange(targetRow, reviewedAtIdx + 1).setValue(syd);
  sh.getRange(targetRow, reviewedByIdx + 1).setValue(data.Reviewed_By || 'Admin');
  sh.getRange(targetRow, adminNoteIdx + 1).setValue(data.Admin_Note || '');

  const rowRange = sh.getRange(targetRow, 1, 1, headers.length);
  if (data.Status === 'Approved') {
    rowRange.setBackground('#C6EFCE');
    syncLeaveToRoster_(allData[targetRow - 1], headers);
  } else if (data.Status === 'Rejected') {
    rowRange.setBackground('#FFC7CE');
  }
  return { ok: true, message: data.Status === 'Approved' ? 'Approved (승인 완료)' : 'Rejected (거절 완료)' };
}

function syncLeaveToRoster_(rowData, headers) {
  try {
    const driver = rowData[headers.indexOf('Driver')];
    const dateFrom = rowData[headers.indexOf('Date_From')];
    const dateTo = rowData[headers.indexOf('Date_To')];
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let rosterSheet = ss.getSheetByName('Driver_Roster');
    if (!rosterSheet) {
      rosterSheet = ss.insertSheet('Driver_Roster');
      rosterSheet.getRange(1, 1, 1, 5).setValues([['Driver','Date','Status','Updated_At','Source']]);
      rosterSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF');
      rosterSheet.setFrozenRows(1);
    }
    const syd = Utilities.formatDate(new Date(), 'Australia/Sydney', 'dd/MM/yyyy HH:mm');
    const from = parseDateFlexible_(dateFrom);
    const to = parseDateFlexible_(dateTo);
    const current = new Date(from);
    while (current <= to) {
      const dateStr = Utilities.formatDate(current, 'Australia/Sydney', 'dd/MM/yyyy');
      const existing = rosterSheet.getDataRange().getValues();
      const rosterH = existing[0];
      const rDriverCol = rosterH.indexOf('Driver');
      const rDateCol = rosterH.indexOf('Date');
      const rStatusCol = rosterH.indexOf('Status');
      const rUpdatedCol = rosterH.indexOf('Updated_At');
      const rSourceCol = rosterH.indexOf('Source');
      if (rDriverCol < 0 || rDateCol < 0) { Logger.log('Driver/Date column not found in Driver_Roster'); return; }
      let found = false;
      for (let i = 1; i < existing.length; i++) {
        if (existing[i][rDriverCol] === driver && existing[i][rDateCol] === dateStr) {
          if (rStatusCol >= 0) rosterSheet.getRange(i + 1, rStatusCol + 1).setValue('LEAVE');
          if (rUpdatedCol >= 0) rosterSheet.getRange(i + 1, rUpdatedCol + 1).setValue(syd);
          if (rSourceCol >= 0) rosterSheet.getRange(i + 1, rSourceCol + 1).setValue('Auto - Leave Approved');
          found = true; break;
        }
      }
      if (!found) rosterSheet.appendRow([driver, dateStr, 'LEAVE', syd, 'Auto - Leave Approved']);
      current.setDate(current.getDate() + 1);
    }
  } catch (e) { Logger.log('Roster sync error: ' + e.toString()); }
}

function parseDateFlexible_(dateStr) {
  if (!dateStr) return new Date();
  const str = String(dateStr);
  if (str.includes('/')) {
    const p = str.split('/');
    return new Date(p[2], p[1] - 1, p[0]);
  }
  return new Date(str);
}

// ═══════════════════════════════════════════════════════════════════════════
// DRIVER ROSTER — 주간 가용현황 (Available / Leave / Worked / Off)
// ═══════════════════════════════════════════════════════════════════════════

/**
 * getRosterData(fromISO, toISO)
 * 기간 내 Driver_Roster + Pre_Departure 기록을 합쳐서 반환
 * Pre_Departure에 기록이 있으면 Worked 상태로 자동 반영
 */
function getRosterData(fromISO, toISO) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const tz = 'Australia/Sydney';

    // ── 1) Driver_Roster 시트에서 수동 상태 로드 ──
    const rosterSheet = ss.getSheetByName('Driver_Roster');
    const rosterMap = {}; // { 'DriverName|yyyy-MM-dd': status }
    if (rosterSheet && rosterSheet.getLastRow() > 1) {
      const rData = rosterSheet.getDataRange().getValues();
      const rH = rData[0];
      const drvCol = rH.indexOf('Driver');
      const dateCol = rH.indexOf('Date');
      const statusCol = rH.indexOf('Status');
      if (drvCol >= 0 && dateCol >= 0 && statusCol >= 0) {
        rData.slice(1).forEach(row => {
          const drv = String(row[drvCol] || '').trim();
          const dateVal = row[dateCol];
          let iso;
          if (dateVal instanceof Date) {
            iso = Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd');
          } else {
            const s = String(dateVal || '');
            if (s.includes('/')) {
              const p = s.split('/');
              iso = p[2] + '-' + p[1].padStart(2,'0') + '-' + p[0].padStart(2,'0');
            } else {
              iso = s;
            }
          }
          if (drv && iso) rosterMap[drv + '|' + iso] = String(row[statusCol] || '').trim();
        });
      }
    }

    // ── 2) Pre_Departure에서 Worked 날짜 수집 ──
    const preSheet = ss.getSheetByName('Pre_Departure');
    const workedMap = {}; // { 'DriverName|yyyy-MM-dd': true }
    if (preSheet && preSheet.getLastRow() > 1) {
      const preData = preSheet.getDataRange().getValues();
      const preH = preData[0];
      const pDrvCol = preH.indexOf('Driver');
      const pDateCol = preH.indexOf('Date');
      if (pDrvCol >= 0 && pDateCol >= 0) {
        preData.slice(1).forEach(row => {
          const drv = String(row[pDrvCol] || '').trim();
          const dateVal = row[pDateCol];
          let iso;
          if (dateVal instanceof Date) {
            iso = Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd');
          } else {
            const s = String(dateVal || '');
            if (s.includes('/')) {
              const p = s.split('/');
              iso = p[2] + '-' + p[1].padStart(2,'0') + '-' + p[0].padStart(2,'0');
            } else {
              iso = s;
            }
          }
          if (drv && iso) workedMap[drv + '|' + iso] = true;
        });
      }
    }

    // ── 3) 병합: Worked 우선, 그 다음 Roster 수동 상태 ──
    // 결과를 배열로 반환
    const result = [];
    const allKeys = new Set([...Object.keys(rosterMap), ...Object.keys(workedMap)]);
    allKeys.forEach(key => {
      const [drv, iso] = key.split('|');
      // 날짜 범위 필터
      if (fromISO && iso < fromISO) return;
      if (toISO && iso > toISO) return;
      const manualStatus = rosterMap[key] || '';
      const worked = workedMap[key] || false;
      // Worked는 Pre_Departure 기록이 있을 때 자동 설정
      // 단, 수동으로 다른 상태(LEAVE, OFF)를 설정한 경우 수동 상태 우선
      let finalStatus;
      if (worked && (!manualStatus || manualStatus === 'Available' || manualStatus === 'WORKED')) {
        finalStatus = 'WORKED';
      } else if (manualStatus) {
        finalStatus = manualStatus;
      } else {
        finalStatus = 'Available';
      }
      result.push({ Driver: drv, Date: iso, Status: finalStatus });
    });

    return { ok: true, roster: result };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

/**
 * updateRosterCell(driver, dateISO, status, user)
 * 관리자가 그리드에서 셀 클릭 시 상태 변경
 */
function updateRosterCell(driver, dateISO, status, user) {
  try {
    if (!driver || !dateISO) return { ok: false, error: 'Missing driver or date' };
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('Driver_Roster');
    if (!sheet) {
      sheet = ss.insertSheet('Driver_Roster');
      const headers = MASTER_HEADERS['Driver_Roster'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }

    const tz = 'Australia/Sydney';
    const now = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');

    // 날짜를 dd/MM/yyyy 형식으로 변환
    const dp = dateISO.split('-');
    const dateDisplay = dp[2] + '/' + dp[1] + '/' + dp[0];

    const sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const drvCol = sheetHeaders.indexOf('Driver');
    const dateCol = sheetHeaders.indexOf('Date');
    const statusCol = sheetHeaders.indexOf('Status');
    const updCol = sheetHeaders.indexOf('Updated_At');
    const srcCol = sheetHeaders.indexOf('Source');

    if (drvCol < 0 || dateCol < 0) return { ok: false, error: 'Required columns not found' };

    // 기존 행 찾기
    const lastRow = sheet.getLastRow();
    let found = false;
    if (lastRow > 1) {
      const data = sheet.getRange(2, 1, lastRow - 1, sheetHeaders.length).getValues();
      for (let i = 0; i < data.length; i++) {
        const rowDrv = String(data[i][drvCol] || '').trim();
        const rowDate = String(data[i][dateCol] || '').trim();
        // dd/MM/yyyy 또는 ISO 형식 모두 대응
        if (rowDrv === driver && (rowDate === dateDisplay || rowDate === dateISO)) {
          if (statusCol >= 0) sheet.getRange(i + 2, statusCol + 1).setValue(status);
          if (updCol >= 0) sheet.getRange(i + 2, updCol + 1).setValue(now);
          if (srcCol >= 0) sheet.getRange(i + 2, srcCol + 1).setValue('Admin - ' + (user || 'unknown'));
          found = true;
          break;
        }
      }
    }

    if (!found) {
      const headers = MASTER_HEADERS['Driver_Roster'];
      const row = headers.map(h => {
        if (h === 'Driver') return driver;
        if (h === 'Date') return dateDisplay;
        if (h === 'Status') return status;
        if (h === 'Updated_At') return now;
        if (h === 'Source') return 'Admin - ' + (user || 'unknown');
        return '';
      });
      sheet.appendRow(row);
    }

    return { ok: true, driver, date: dateISO, status };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// FATIGUE COMPLIANCE — NHVR (National Heavy Vehicle Regulator) Table 2
// ═══════════════════════════════════════════════════════════════════════════

/**
 * getFatigueComplianceCheck()
 * Returns fatigue alerts for ALL drivers:
 *   - consecutive_work: drivers working 6+ consecutive days without 24hr rest
 *   - seven_day_rest: drivers missing 24hr continuous Night Rest in last 7 days
 *   - twentyeight_day_rest: drivers missing 4× 24hr Night Rest in last 28 days
 *   - rest_gap_violation: drivers whose last EoS → next Pre time gap < 7 hours
 */
function getFatigueComplianceCheck() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const tz = 'Australia/Sydney';
    const now = new Date();
    const sydNow = new Date(Utilities.formatDate(now, tz, "yyyy-MM-dd'T'HH:mm:ss"));
    const alerts = [];

    // ── Collect all driver names from Drivers master ──
    const drvSheet = ss.getSheetByName('Drivers');
    const driverNames = [];
    if (drvSheet && drvSheet.getLastRow() > 1) {
      const drvData = drvSheet.getDataRange().getValues();
      const drvH = drvData[0];
      const nameIdx = drvH.indexOf('Name_EN') >= 0 ? drvH.indexOf('Name_EN') : 0;
      const nameKrIdx = drvH.indexOf('Name_KR');
      drvData.slice(1).forEach(r => {
        const n = String(r[nameIdx] || '').trim();
        if (n) driverNames.push({ en: n, kr: nameKrIdx >= 0 ? String(r[nameKrIdx] || '').trim() : '' });
      });
    }

    // ── Collect work dates per driver from Pre_Departure ──
    const preSheet = ss.getSheetByName('Pre_Departure');
    const driverWorkDates = {}; // { driverName: Set of 'yyyy-MM-dd' }
    if (preSheet && preSheet.getLastRow() > 1) {
      const preData = preSheet.getDataRange().getValues();
      const preH = preData[0];
      preData.slice(1).forEach(row => {
        const obj = {};
        preH.forEach((h, i) => obj[h] = row[i]);
        const drv = String(obj.Driver || '').trim();
        if (!drv) return;
        if (!driverWorkDates[drv]) driverWorkDates[drv] = new Set();
        const d = obj.Date instanceof Date
          ? Utilities.formatDate(obj.Date, tz, 'yyyy-MM-dd')
          : parseDateToISO_(obj.Date);
        if (d) driverWorkDates[drv].add(d);
      });
    }

    // ── Collect leave dates per driver from Driver_Roster ──
    const rosterSheet = ss.getSheetByName('Driver_Roster');
    const driverLeaveDates = {}; // { driverName: Set of 'yyyy-MM-dd' }
    if (rosterSheet && rosterSheet.getLastRow() > 1) {
      const rData = rosterSheet.getDataRange().getValues();
      const rH = rData[0];
      const rDriverIdx = rH.indexOf('Driver');
      const rDateIdx = rH.indexOf('Date');
      const rStatusIdx = rH.indexOf('Status');
      if (rDriverIdx < 0 || rDateIdx < 0 || rStatusIdx < 0) {
        Logger.log('Driver_Roster missing required columns (Driver/Date/Status)');
      } else {
      rData.slice(1).forEach(row => {
        const drv = String(row[rDriverIdx] || '').trim();
        const status = String(row[rStatusIdx] || '').trim();
        if (drv && status === 'LEAVE') {
          if (!driverLeaveDates[drv]) driverLeaveDates[drv] = new Set();
          const dateVal = row[rDateIdx];
          const d = dateVal instanceof Date
            ? Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd')
            : parseDateToISO_(dateVal);
          if (d) driverLeaveDates[drv].add(d);
        }
      });
      } // else (columns found)
    }

    // ── Collect last End_of_Shift time per driver ──
    const eosSheet = ss.getSheetByName('End_of_Shift');
    const driverLastEos = {}; // { driverName: { date, endTime, submitted } }
    if (eosSheet && eosSheet.getLastRow() > 1) {
      const eosData = eosSheet.getDataRange().getValues();
      const eosH = eosData[0];
      eosData.slice(1).forEach(row => {
        const obj = {};
        eosH.forEach((h, i) => obj[h] = row[i]);
        const drv = String(obj.Driver || '').trim();
        if (!drv) return;
        const dateStr = obj.Date instanceof Date
          ? Utilities.formatDate(obj.Date, tz, 'yyyy-MM-dd')
          : parseDateToISO_(obj.Date);
        const endTime = obj.End_Time instanceof Date
          ? Utilities.formatDate(obj.End_Time, tz, 'HH:mm')
          : String(obj.End_Time || '').trim();
        const submitted = String(obj.Submitted || '').trim();
        // Keep latest
        if (!driverLastEos[drv] || (dateStr && dateStr > (driverLastEos[drv].date || ''))) {
          driverLastEos[drv] = { date: dateStr, endTime: endTime, submitted: submitted };
        } else if (dateStr === (driverLastEos[drv].date || '') && endTime > (driverLastEos[drv].endTime || '')) {
          driverLastEos[drv] = { date: dateStr, endTime: endTime, submitted: submitted };
        }
      });
    }

    // ── Collect first Pre_Departure time per driver per date ──
    const driverFirstPre = {}; // { driverName: { date: startTime } }
    if (preSheet && preSheet.getLastRow() > 1) {
      const preData = preSheet.getDataRange().getValues();
      const preH = preData[0];
      preData.slice(1).forEach(row => {
        const obj = {};
        preH.forEach((h, i) => obj[h] = row[i]);
        const drv = String(obj.Driver || '').trim();
        if (!drv) return;
        const dateStr = obj.Date instanceof Date
          ? Utilities.formatDate(obj.Date, tz, 'yyyy-MM-dd')
          : parseDateToISO_(obj.Date);
        const startTime = obj.Start_Time instanceof Date
          ? Utilities.formatDate(obj.Start_Time, tz, 'HH:mm')
          : String(obj.Start_Time || '').trim();
        if (!dateStr || !startTime) return;
        if (!driverFirstPre[drv]) driverFirstPre[drv] = {};
        if (!driverFirstPre[drv][dateStr] || startTime < driverFirstPre[drv][dateStr]) {
          driverFirstPre[drv][dateStr] = startTime;
        }
      });
    }

    // ── Check each driver ──
    const todayISO = Utilities.formatDate(now, tz, 'yyyy-MM-dd');

    driverNames.forEach(drv => {
      const name = drv.en;
      const displayName = drv.kr ? drv.kr + ' (' + drv.en + ')' : drv.en;
      const workDates = driverWorkDates[name] || new Set();
      const leaveDates = driverLeaveDates[name] || new Set();

      // ─── 1. Consecutive work days (6+ without a rest day) ───
      let consecutive = 0;
      for (let i = 0; i < 14; i++) {
        const d = new Date(sydNow);
        d.setDate(d.getDate() - i);
        const ds = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
        if (workDates.has(ds) && !leaveDates.has(ds)) {
          consecutive++;
        } else {
          break;
        }
      }
      if (consecutive >= 6) {
        alerts.push({
          type: 'consecutive_work',
          driver: displayName,
          days: consecutive,
          severity: consecutive >= 7 ? 'critical' : 'warning'
        });
      }

      // ─── 2. 7-day rest check (need 24hr continuous Night Rest in last 7 days) ───
      let hasNightRest7 = false;
      for (let i = 0; i < 7; i++) {
        const d = new Date(sydNow);
        d.setDate(d.getDate() - i);
        const ds = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
        // A day with no work AND no next-day early start = rest day
        if (!workDates.has(ds) || leaveDates.has(ds)) {
          hasNightRest7 = true;
          break;
        }
      }
      if (!hasNightRest7 && workDates.size > 0) {
        alerts.push({
          type: 'seven_day_rest',
          driver: displayName,
          severity: 'critical'
        });
      }

      // ─── 3. 28-day rest check (need 4× 24hr Night Rest days in last 28 days) ───
      let restDays28 = 0;
      for (let i = 0; i < 28; i++) {
        const d = new Date(sydNow);
        d.setDate(d.getDate() - i);
        const ds = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
        if (!workDates.has(ds) || leaveDates.has(ds)) {
          restDays28++;
        }
      }
      if (restDays28 < 4 && workDates.size > 0) {
        alerts.push({
          type: 'twentyeight_day_rest',
          driver: displayName,
          restDays: restDays28,
          severity: restDays28 < 2 ? 'critical' : 'warning'
        });
      }

      // ─── 4. 7-hour rest gap (last EoS End_Time → today's first Pre Start_Time) ───
      const lastEos = driverLastEos[name];
      if (lastEos && lastEos.date && lastEos.endTime && lastEos.endTime.includes(':')) {
        // Find next day's first pre-departure
        const eosDate = new Date(lastEos.date + 'T' + lastEos.endTime + ':00');
        const nextDay = new Date(lastEos.date);
        nextDay.setDate(nextDay.getDate() + 1);
        const nextDayISO = Utilities.formatDate(nextDay, tz, 'yyyy-MM-dd');

        if (driverFirstPre[name] && driverFirstPre[name][nextDayISO]) {
          const preTime = driverFirstPre[name][nextDayISO];
          const preDate = new Date(nextDayISO + 'T' + preTime + ':00');
          const gapHours = (preDate.getTime() - eosDate.getTime()) / (1000 * 60 * 60);
          if (gapHours < 7 && gapHours >= 0) {
            alerts.push({
              type: 'rest_gap_violation',
              driver: displayName,
              eosDate: lastEos.date,
              eosTime: lastEos.endTime,
              preDate: nextDayISO,
              preTime: preTime,
              gapHours: Math.round(gapHours * 10) / 10,
              severity: gapHours < 5 ? 'critical' : 'warning'
            });
          }
        }
      }
    });

    return { ok: true, alerts: alerts };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

/**
 * getLastEndOfShift(driverName)
 * Returns the most recent End_of_Shift record for a specific driver
 * Used by driver app to check 7-hour rest gap before allowing Pre-Departure
 */
function getLastEndOfShift(driverName) {
  try {
    if (!driverName) return { ok: false, error: 'driver param required' };
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const tz = 'Australia/Sydney';
    const eosSheet = ss.getSheetByName('End_of_Shift');
    if (!eosSheet || eosSheet.getLastRow() < 2) return { ok: true, lastEos: null };

    const eosData = eosSheet.getDataRange().getValues();
    const eosH = eosData[0];
    let latest = null;
    let latestKey = '';

    eosData.slice(1).forEach(row => {
      const obj = {};
      eosH.forEach((h, i) => obj[h] = row[i]);
      if (String(obj.Driver || '').trim() !== driverName.trim()) return;

      const dateStr = obj.Date instanceof Date
        ? Utilities.formatDate(obj.Date, tz, 'yyyy-MM-dd')
        : parseDateToISO_(obj.Date);
      const endTime = obj.End_Time instanceof Date
        ? Utilities.formatDate(obj.End_Time, tz, 'HH:mm')
        : String(obj.End_Time || '').trim();
      const key = (dateStr || '') + 'T' + (endTime || '');
      if (key > latestKey) {
        latestKey = key;
        latest = {
          date: dateStr,
          endTime: endTime,
          dateDMY: dateStr ? formatDateDMY_(dateStr) : ''
        };
      }
    });

    return { ok: true, lastEos: latest };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

/** Helper: convert various date formats to yyyy-MM-dd */
function parseDateToISO_(val) {
  if (!val) return '';
  const str = String(val).trim().replace(/\s+.*/, '');
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(str)) {
    const p = str.split('/');
    return p[2] + '-' + p[1] + '-' + p[0];
  }
  try {
    const d = new Date(val);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, 'Australia/Sydney', 'yyyy-MM-dd');
    }
  } catch(e) {}
  return '';
}

/** Helper: yyyy-MM-dd → dd/MM/yyyy */
function formatDateDMY_(isoStr) {
  if (!isoStr || !isoStr.includes('-')) return isoStr;
  const p = isoStr.split('-');
  return p[2] + '/' + p[1] + '/' + p[0];
}

// ═══════════════════════════════════════════════════════════════════════════
// HVIS Bookings — Google Sheets 동기화
// ═══════════════════════════════════════════════════════════════════════════

function saveHvisBooking(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'HVIS_Bookings');
    const headers = MASTER_HEADERS['HVIS_Bookings'];
    const row = headers.map(h => {
      const nk = normalizeKey(h);
      for (const k of Object.keys(data)) {
        if (normalizeKey(k) === nk) return data[k] || '';
      }
      return '';
    });
    sheet.appendRow(row);
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

function deleteHvisBooking(id) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'HVIS_Bookings');
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { ok: false, error: 'No data' };
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    const idCol = headers.indexOf('ID');
    if (idCol < 0) return { ok: false, error: 'ID column not found' };
    const data = sheet.getRange(2, idCol + 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 2);
        return { ok: true };
      }
    }
    return { ok: false, error: 'ID not found: ' + id };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Driver Photo Upload — Google Drive
// ═══════════════════════════════════════════════════════════════════════════

/**
 * 드라이버 이름별 폴더에 사진 업로드
 * DongChoi_DriverDocs / {driverName} / {photoKey}.jpg
 */
function uploadDriverPhoto(driverName, photoKey, dataUrl, mimeType) {
  try {
    if (!driverName || !photoKey || !dataUrl) {
      return { ok: false, error: 'Missing required fields' };
    }
    const base64 = dataUrl.replace(/^data:[^;]+;base64,/, '');
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType || 'image/jpeg', photoKey + '.jpg');

    const rootFolder = getOrCreateFolder_(null, DRIVE_ROOT_FOLDER);
    const driverFolder = getOrCreateFolder_(rootFolder, driverName);

    // 기존 같은 photoKey 파일 삭제 (최신 1장만 유지)
    const existing = driverFolder.getFilesByName(photoKey + '.jpg');
    while (existing.hasNext()) { existing.next().setTrashed(true); }

    const file = driverFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return { ok: true, fileId: file.getId(), url: 'https://drive.google.com/uc?id=' + file.getId(), photoKey: photoKey };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

/** 드라이버의 모든 사진 URL 조회 */
function getDriverPhotos(driverName) {
  try {
    if (!driverName) return { ok: false, error: 'Missing driverName' };
    const rootFolders = DriveApp.getFoldersByName(DRIVE_ROOT_FOLDER);
    if (!rootFolders.hasNext()) return { ok: true, photos: {} };
    const driverFolders = rootFolders.next().getFoldersByName(driverName);
    if (!driverFolders.hasNext()) return { ok: true, photos: {} };
    const driverFolder = driverFolders.next();

    const photos = {};
    ['licFront', 'licBack', 'authFront', 'authBack'].forEach(function(key) {
      const files = driverFolder.getFilesByName(key + '.jpg');
      if (files.hasNext()) {
        const f = files.next();
        photos[key] = {
          fileId: f.getId(),
          url: 'https://drive.google.com/uc?id=' + f.getId(),
          updated: Utilities.formatDate(f.getLastUpdated(), 'Australia/Sydney', 'dd/MM/yyyy HH:mm')
        };
      }
    });
    return { ok: true, photos: photos };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

/** 폴더 찾기 또는 생성 헬퍼 */
function getOrCreateFolder_(parent, name) {
  var folders = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent ? parent.createFolder(name) : DriveApp.createFolder(name);
}

// ═══════════════════════════════════════════════════════════════════════════
// Maint Records (정비 기록)
// ═══════════════════════════════════════════════════════════════════════════
function saveMaintRecord(data) {
  try {
    if (!data || !data.ID) return { ok: false, error: 'Missing ID' };
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'Maint_Records');
    const headers = MASTER_HEADERS['Maint_Records'];

    // 시트 헤더에서 ID 컬럼 위치 동적 조회
    const lastRow = sheet.getLastRow();
    let found = false;
    if (lastRow > 1) {
      const sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const idCol = sheetHeaders.indexOf('ID');
      if (idCol < 0) return { ok: false, error: 'ID column not found in Maint_Records' };
      const ids = sheet.getRange(2, idCol + 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(data.ID)) {
          const row = headers.map(h => data[h] !== undefined ? data[h] : '');
          sheet.getRange(i + 2, 1, 1, headers.length).setValues([row]);
          found = true;
          break;
        }
      }
    }
    if (!found) {
      const row = headers.map(h => data[h] !== undefined ? data[h] : '');
      sheet.appendRow(row);
    }

    // ── M_Vehicles 동기화: 정비 기록 저장 시 Next_Service_KM / Last_Service_KM 자동 갱신 ──
    // (정비 카드 / 대시보드 알림이 차량 마스터의 이 필드를 읽기 때문에 반드시 동기화 필요)
    try {
      const rego = data.Rego;
      const nextKM = Number(data.NextServiceKM) || 0;
      const lastKM = Number(data.KM) || 0;  // 정비 시점 KM = Last_Service_KM
      if (rego && (nextKM > 0 || lastKM > 0)) {
        const vSheet = ss.getSheetByName('M_Vehicles');
        if (vSheet) {
          const vLastRow = vSheet.getLastRow();
          const vLastCol = vSheet.getLastColumn();
          if (vLastRow >= 2) {
            const vHeaders = vSheet.getRange(1, 1, 1, vLastCol).getValues()[0];
            const regoCol = vHeaders.indexOf('Rego');
            if (regoCol >= 0) {
              // Next_Service_KM 컬럼이 없으면 자동 생성 (Last_Service_KM 다음 위치)
              let nextSvcCol = vHeaders.indexOf('Next_Service_KM');
              if (nextSvcCol < 0 && nextKM > 0) {
                const lastSvcIdx = vHeaders.indexOf('Last_Service_KM');
                if (lastSvcIdx >= 0) {
                  vSheet.insertColumnAfter(lastSvcIdx + 1);
                  vSheet.getRange(1, lastSvcIdx + 2).setValue('Next_Service_KM');
                  nextSvcCol = lastSvcIdx + 1;
                } else {
                  vSheet.getRange(1, vLastCol + 1).setValue('Next_Service_KM');
                  nextSvcCol = vLastCol;
                }
              }
              const lastSvcCol = vHeaders.indexOf('Last_Service_KM');
              // Rego 매칭 행 검색
              const vRegos = vSheet.getRange(2, regoCol + 1, vLastRow - 1, 1).getValues();
              for (let i = 0; i < vRegos.length; i++) {
                if (String(vRegos[i][0]) === String(rego)) {
                  if (nextKM > 0 && nextSvcCol >= 0) {
                    vSheet.getRange(i + 2, nextSvcCol + 1).setValue(nextKM);
                  }
                  if (lastKM > 0 && lastSvcCol >= 0) {
                    vSheet.getRange(i + 2, lastSvcCol + 1).setValue(lastKM);
                  }
                  break;
                }
              }
            }
          }
        }
      }
    } catch (e2) {
      // M_Vehicles 동기화 실패는 본 저장에 영향 없도록 흡수 (로그만 남김)
      Logger.log('saveMaintRecord: M_Vehicles sync skipped: ' + e2);
    }

    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Generic: Delete row by ID column
// ═══════════════════════════════════════════════════════════════════════════
function deleteSheetRowById(sheetName, idCol, idValue) {
  try {
    if (!idValue) return { ok: false, error: 'Missing ID value' };
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, sheetName);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: false, error: 'No data' };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colIdx = headers.indexOf(idCol);
    if (colIdx < 0) return { ok: false, error: 'Column not found: ' + idCol };

    const vals = sheet.getRange(2, colIdx + 1, lastRow - 1, 1).getValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      if (String(vals[i][0]) === String(idValue)) {
        sheet.deleteRow(i + 2);
        return { ok: true };
      }
    }
    return { ok: false, error: 'ID not found: ' + idValue };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Invoice Overrides (서차지 오버라이드)
// ═══════════════════════════════════════════════════════════════════════════
function saveInvoiceOverride(rowKey, value) {
  try {
    if (!rowKey) return { ok: false, error: 'Missing rowKey' };
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'Invoice_Overrides');

    const lastRow = sheet.getLastRow();
    const sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rkCol = sheetHeaders.indexOf('RowKey');
    const valCol = sheetHeaders.indexOf('Value');
    if (rkCol < 0) return { ok: false, error: 'RowKey column not found' };

    let found = false;
    if (lastRow > 1) {
      const keys = sheet.getRange(2, rkCol + 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < keys.length; i++) {
        if (String(keys[i][0]) === String(rowKey)) {
          if (value === null || value === undefined || value === '') {
            sheet.deleteRow(i + 2);
          } else {
            sheet.getRange(i + 2, (valCol >= 0 ? valCol : 1) + 1).setValue(value);
          }
          found = true;
          break;
        }
      }
    }
    if (!found && value !== null && value !== undefined && value !== '') {
      sheet.appendRow([rowKey, value]);
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

function bulkSaveInvoiceOverrides(items) {
  try {
    if (!items || !items.length) return { ok: true };
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'Invoice_Overrides');

    const sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rkCol = sheetHeaders.indexOf('RowKey');
    const valCol = sheetHeaders.indexOf('Value');
    if (rkCol < 0) return { ok: false, error: 'RowKey column not found' };

    // 기존 데이터 로드
    const lastRow = sheet.getLastRow();
    const existing = {};
    if (lastRow > 1) {
      const data = sheet.getRange(2, rkCol + 1, lastRow - 1, 1).getValues();
      data.forEach((row, i) => { existing[String(row[0])] = i + 2; });
    }

    const valColNum = (valCol >= 0 ? valCol : 1) + 1;
    items.forEach(item => {
      const rk = String(item.rowKey);
      if (existing[rk]) {
        if (item.value === null || item.value === '') {
          sheet.getRange(existing[rk], valColNum).setValue('__DELETE__');
        } else {
          sheet.getRange(existing[rk], valColNum).setValue(item.value);
        }
      } else if (item.value !== null && item.value !== '') {
        sheet.appendRow([rk, item.value]);
      }
    });

    // __DELETE__ 마킹된 행 제거 (역순)
    const lr2 = sheet.getLastRow();
    if (lr2 > 1) {
      const vals = sheet.getRange(2, valColNum, lr2 - 1, 1).getValues();
      for (let i = vals.length - 1; i >= 0; i--) {
        if (vals[i][0] === '__DELETE__') sheet.deleteRow(i + 2);
      }
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Company Profile (회사 정보)
// ═══════════════════════════════════════════════════════════════════════════
function saveCompanyProfile(data) {
  try {
    if (!data) return { ok: false, error: 'Missing data' };
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'Company_Profile');

    const sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const keyCol = sheetHeaders.indexOf('Key');
    const valueCol = sheetHeaders.indexOf('Value');
    if (keyCol < 0) return { ok: false, error: 'Key column not found in Company_Profile' };

    // 기존 키-값 쌍 로드
    const lastRow = sheet.getLastRow();
    const existing = {};
    if (lastRow > 1) {
      const rows = sheet.getRange(2, keyCol + 1, lastRow - 1, 1).getValues();
      rows.forEach((row, i) => { existing[String(row[0])] = i + 2; });
    }

    // 각 키-값 업데이트 또는 추가
    const valColNum = (valueCol >= 0 ? valueCol : 1) + 1;
    Object.keys(data).forEach(key => {
      const val = data[key] || '';
      if (existing[key]) {
        sheet.getRange(existing[key], valColNum).setValue(val);
      } else {
        sheet.appendRow([key, val]);
      }
    });

    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Invoice Deductions (인보이스 공제)
// ═══════════════════════════════════════════════════════════════════════════
function saveInvoiceDeduction(data) {
  try {
    if (!data || !data.ID) return { ok: false, error: 'Missing ID' };
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'Invoice_Deductions');
    const headers = MASTER_HEADERS['Invoice_Deductions'];
    const row = headers.map(h => data[h] !== undefined ? data[h] : '');
    sheet.appendRow(row);
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

function saveInvoiceDeductionsBulk(agency, period, items) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'Invoice_Deductions');
    const headers = MASTER_HEADERS['Invoice_Deductions'];

    // 해당 agency+period 기존 행 삭제 (역순)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const agIdx = hdr.indexOf('Agency');
      const prIdx = hdr.indexOf('Period');
      const data = sheet.getRange(2, 1, lastRow - 1, hdr.length).getValues();
      for (let i = data.length - 1; i >= 0; i--) {
        if (String(data[i][agIdx]) === String(agency) && String(data[i][prIdx]) === String(period)) {
          sheet.deleteRow(i + 2);
        }
      }
    }

    // 새 항목 추가
    if (items && items.length) {
      items.forEach(item => {
        item.Agency = agency;
        item.Period = period;
        if (!item.ID) item.ID = Date.now().toString() + Math.random().toString(36).slice(2, 6);
        const row = headers.map(h => item[h] !== undefined ? item[h] : '');
        sheet.appendRow(row);
      });
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Invoice Manual Items (인보이스 수동 항목)
// ═══════════════════════════════════════════════════════════════════════════
function saveInvoiceManualItem(data) {
  try {
    if (!data || !data.ID) return { ok: false, error: 'Missing ID' };
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'Invoice_Manual_Items');
    const headers = MASTER_HEADERS['Invoice_Manual_Items'];

    const lastRow = sheet.getLastRow();
    let found = false;
    if (lastRow > 1) {
      const sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const idCol = sheetHeaders.indexOf('ID');
      if (idCol < 0) return { ok: false, error: 'ID column not found in Invoice_Manual_Items' };
      const ids = sheet.getRange(2, idCol + 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(data.ID)) {
          const row = headers.map(h => data[h] !== undefined ? data[h] : '');
          sheet.getRange(i + 2, 1, 1, headers.length).setValues([row]);
          found = true;
          break;
        }
      }
    }
    if (!found) {
      const row = headers.map(h => data[h] !== undefined ? data[h] : '');
      sheet.appendRow(row);
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

function saveInvoiceManualItemsBulk(agency, period, items) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'Invoice_Manual_Items');
    const headers = MASTER_HEADERS['Invoice_Manual_Items'];

    // 해당 agency+period 기존 행 삭제 (역순)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const agIdx = hdr.indexOf('Agency');
      const prIdx = hdr.indexOf('Period');
      const data = sheet.getRange(2, 1, lastRow - 1, hdr.length).getValues();
      for (let i = data.length - 1; i >= 0; i--) {
        if (String(data[i][agIdx]) === String(agency) && String(data[i][prIdx]) === String(period)) {
          sheet.deleteRow(i + 2);
        }
      }
    }

    // 새 항목 추가
    if (items && items.length) {
      items.forEach(item => {
        item.Agency = agency;
        item.Period = period;
        if (!item.ID) item.ID = Date.now().toString() + Math.random().toString(36).slice(2, 6);
        const row = headers.map(h => item[h] !== undefined ? item[h] : '');
        sheet.appendRow(row);
      });
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Schedule (운행 일정) — 중기 자동화 핵심 시트
// ═══════════════════════════════════════════════════════════════════════════
//
// 상태 흐름:
//   scheduled → in_progress → completed → invoiced → paid
//                                       ↘ cancelled
//
// 자동 상태 전환 (매일 새벽 1시 트리거):
//   StartDate <= 오늘 <= EndDate    → in_progress
//   EndDate < 오늘 + 'scheduled'/'in_progress' → completed
//   인보이스 발행/결제 시 → invoiced/paid (admin.html에서 호출)
// ═══════════════════════════════════════════════════════════════════════════

const SCHEDULE_STATUSES = ['scheduled','in_progress','completed','invoiced','paid','cancelled'];

/**
 * 운행 일정 조회 (필터링 가능)
 * filters: { status, agency, from(YYYY-MM-DD), to(YYYY-MM-DD) }
 */
function getSchedule(filters) {
  try {
    filters = filters || {};
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'Schedule');
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { ok: true, rows: [] };
    const headers = data[0];
    const DATE_FIELDS = ['StartDate','EndDate','CreatedAt','UpdatedAt'];
    // ★ DD/MM/YYYY → YYYY-MM-DD 변환 (필터 비교용)
    const _toISO = (s) => {
      const str = String(s||'').trim();
      if (!str) return '';
      // 이미 YYYY-MM-DD 형식
      if (/^\d{4}-\d{2}-\d{2}/.test(str)) return str.slice(0,10);
      // DD/MM/YYYY 형식
      const m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (m) return m[3] + '-' + String(m[2]).padStart(2,'0') + '-' + String(m[1]).padStart(2,'0');
      return str;
    };
    let rows = [];
    for (let i = 1; i < data.length; i++) {
      const obj = {};
      headers.forEach((h, ci) => {
        let v = data[i][ci];
        if (DATE_FIELDS.indexOf(h) !== -1 && v instanceof Date && !isNaN(v)) {
          v = formatDateForSheet(v);
        }
        obj[h] = v;
      });
      obj._rowIndex = i + 1;
      rows.push(obj);
    }
    if (filters.status) rows = rows.filter(r => String(r.Status||'').trim() === filters.status);
    if (filters.agency) rows = rows.filter(r => String(r.Agency||'').trim().toLowerCase() === filters.agency.toLowerCase());
    // ★ 날짜 필터 — ISO 형식으로 정규화 후 비교
    if (filters.from)   rows = rows.filter(r => _toISO(r.EndDate)   >= filters.from);
    if (filters.to)     rows = rows.filter(r => _toISO(r.StartDate) <= filters.to);
    rows.sort((a, b) => _toISO(b.StartDate).localeCompare(_toISO(a.StartDate)));
    return { ok: true, rows: rows };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

/**
 * 드라이버에게 배정된 일정 조회
 * driver: 드라이버 한국어 이름 (예: "최동철")
 * from/to: 'YYYY-MM-DD' (해당 범위에 일부라도 걸치는 일정 반환)
 * 반환: 일별 슬롯 평탄화 [{ tourId, tourCode, agency, date, slotKey, slot, hotel, guide, guidePhone, pax, seats, flightIn, flightOut, status }]
 */
function getDriverSchedule(driver, from, to) {
  try {
    if (!driver) return { ok: true, slots: [] };
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'Schedule');
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { ok: true, slots: [] };
    const headers = data[0];
    const DATE_FIELDS = ['StartDate','EndDate'];
    const idx = {};
    headers.forEach((h, ci) => idx[h] = ci);

    // ★ DD/MM/YYYY → YYYY-MM-DD 변환 (필터 비교용)
    const _toISO = (s) => {
      const str = String(s||'').trim();
      if (!str) return '';
      if (/^\d{4}-\d{2}-\d{2}/.test(str)) return str.slice(0,10);
      const m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (m) return m[3] + '-' + String(m[2]).padStart(2,'0') + '-' + String(m[1]).padStart(2,'0');
      return str;
    };

    const result = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = String(row[idx.Status]||'').trim();
      if (status === 'cancelled') continue;
      // 날짜 범위 체크 — ISO 형식으로 변환 후 비교
      const sdRaw = row[idx.StartDate];
      const edRaw = row[idx.EndDate];
      const sd = sdRaw instanceof Date ? Utilities.formatDate(sdRaw, 'Australia/Sydney', 'yyyy-MM-dd') : _toISO(sdRaw);
      const ed = edRaw instanceof Date ? Utilities.formatDate(edRaw, 'Australia/Sydney', 'yyyy-MM-dd') : _toISO(edRaw);
      if (from && ed && ed < from) continue;
      if (to && sd && sd > to) continue;

      // TourPlan 파싱
      let days = [];
      try { days = JSON.parse(row[idx.TourPlan] || '[]'); } catch(e) { continue; }
      if (!Array.isArray(days)) continue;

      const tourId = row[idx.TourID];
      const tourCode = row[idx.TourCode] || '';
      const agency = row[idx.Agency] || '';
      const guide = row[idx.Guide] || '';
      const guidePhone = row[idx.GuidePhone] || '';
      const pax = row[idx.Pax] || '';
      const seats = row[idx.Seats] || '';
      const flightIn = row[idx.FlightIn] || '';
      const flightOut = row[idx.FlightOut] || '';
      const hotel = row[idx.Hotel] || '';
      // ★ BillingEntity — 빈 값이면 'DC' (자사 발행 기본)
      const billingEntity = String(row[idx.BillingEntity] || '').trim() || 'DC';

      days.forEach(d => {
        if (!d || !d.date) return;
        const dateStr = String(d.date).slice(0,10);
        if (from && dateStr < from) return;
        if (to && dateStr > to) return;
        // 그 날 트레일러 사용 여부
        const trailer = !!d.trailer;
        ['morning','fullday','evening'].forEach(slotKey => {
          const slot = d.slots && d.slots[slotKey];
          if (!slot) return;
          // ★ 드라이버 매칭 — prefix(🏠/🏢/⚠️/🚫 등) 제거 후 비교
          //   어드민 dropdown 라벨이 잘못 저장된 경우 대비
          const _stripPrefix = (s) => String(s||'')
            .replace(/^[\u2B50\u26A0\uFE0F\u26AA\s]*/, '')      // ⭐⚠️⚪
            .replace(/^[\u{1F3E0}\u{1F3E2}\u{1F3E8}]\s*/u, '')  // 🏠🏢🏨
            .replace(/^[\u{1F535}\u{1F6AB}]\s*/u, '')           // 🔵🚫
            .replace(/^[\u{1F690}\u{1F68C}\u{1F699}\u{1F69B}\u{1F69C}]\s*/u, '') // 🚐🚌🚙🚛🚜
            .trim();
          // ★ 슬롯 모드(자사/외주)에 따라 driver 필드 위치가 다름
          //   자사: slot.driver = 드라이버 이름
          //   외주: slot.subDriver = 외주 드라이버 이름
          //   둘 다 매칭 시도 → 외주 드라이버도 자기 일정 볼 수 있게
          const slotDriver = _stripPrefix(slot.driver);
          const slotSubDriver = _stripPrefix(slot.subDriver);
          const targetDriver = _stripPrefix(driver);
          const isMatch = (slotDriver === targetDriver) || (slotSubDriver === targetDriver);
          if (!isMatch) return;
          // ★ 외주 매칭이면 slot에 모드 표시 (드라이버 앱이 사용)
          const isSubMode = (slot.mode === 'sub') || (slotSubDriver === targetDriver && slotDriver !== targetDriver);
          result.push({
            tourId: tourId,
            tourCode: tourCode,
            agency: agency,
            billingEntity: billingEntity,
            BillingEntity: billingEntity,
            date: dateStr,
            slotKey: slotKey,
            slot: slot,
            isSubMode: isSubMode,  // ★ 외주 모드 슬롯 식별
            hotel: hotel,
            trailer: trailer,
            guide: guide,
            guidePhone: guidePhone,
            pax: pax,
            seats: seats,
            flightIn: flightIn,
            flightOut: flightOut,
            status: status
          });
        });
      });
    }
    // 날짜순 정렬 → 같은 날 슬롯 순
    const slotOrder = { morning: 0, fullday: 1, evening: 2 };
    result.sort((a,b) => {
      if (a.date !== b.date) return a.date < b.date ? -1 : 1;
      return slotOrder[a.slotKey] - slotOrder[b.slotKey];
    });
    return { ok: true, slots: result };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

/**
 * 운행 일정 통계 (대시보드용)
 */
function getScheduleStats() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Schedule');
    if (!sheet) return { ok: true, stats: { total: 0, byStatus: {} } };
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { ok: true, stats: { total: 0, byStatus: {} } };
    const headers = data[0];
    const statusIdx = headers.indexOf('Status');
    const stats = { total: data.length - 1, byStatus: {} };
    SCHEDULE_STATUSES.forEach(s => stats.byStatus[s] = 0);
    for (let i = 1; i < data.length; i++) {
      const s = String(data[i][statusIdx]||'').trim();
      if (stats.byStatus[s] !== undefined) stats.byStatus[s]++;
    }
    return { ok: true, stats: stats };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

/**
 * 운행 일정 추가/수정
 * data.TourID 가 있으면 수정, 없으면 추가
 */
function saveSchedule(data, user) {
  try {
    if (!data) return { ok: false, error: 'data is empty' };
    if (!data.Agency)    return { ok: false, error: '여행사를 선택하세요' };
    if (!data.StartDate) return { ok: false, error: '시작일을 입력하세요' };
    if (!data.EndDate)   return { ok: false, error: '종료일을 입력하세요' };

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'Schedule');
    const headers = MASTER_HEADERS['Schedule'];
    const allData = sheet.getDataRange().getValues();
    const sheetHeaders = allData[0];
    const tourIdCol = sheetHeaders.indexOf('TourID');
    if (tourIdCol < 0) return { ok: false, error: 'TourID column not found' };

    const now = new Date();
    const sydNow = Utilities.formatDate(now, 'Australia/Sydney', 'yyyy-MM-dd HH:mm:ss');

    let existingRow = -1;
    let existingCreated = '';
    if (data.TourID) {
      for (let i = 1; i < allData.length; i++) {
        if (String(allData[i][tourIdCol]).trim() === String(data.TourID).trim()) {
          existingRow = i + 1;
          existingCreated = allData[i][sheetHeaders.indexOf('CreatedAt')] || '';
          break;
        }
      }
    }

    if (!data.TourID) {
      const yymm = Utilities.formatDate(now, 'Australia/Sydney', 'yyyyMM');
      let maxSeq = 0;
      const prefix = `T${yymm}-`;
      for (let i = 1; i < allData.length; i++) {
        const id = String(allData[i][tourIdCol] || '');
        if (id.indexOf(prefix) === 0) {
          const seq = parseInt(id.substring(prefix.length), 10) || 0;
          if (seq > maxSeq) maxSeq = seq;
        }
      }
      data.TourID = `${prefix}${String(maxSeq + 1).padStart(3, '0')}`;
    }

    if (!data.Status) data.Status = 'scheduled';
    if (!data.CreatedAt) data.CreatedAt = existingCreated || sydNow;
    data.UpdatedAt = sydNow;

    if (data.Status === 'scheduled' || data.Status === 'in_progress') {
      const today = Utilities.formatDate(now, 'Australia/Sydney', 'yyyy-MM-dd');
      const sd = String(data.StartDate||'');
      const ed = String(data.EndDate||'');
      if (today >= sd && today <= ed) data.Status = 'in_progress';
      else if (today > ed) data.Status = 'completed';
    }

    // ★ 시트의 실제 헤더 순서로 row 만들기
    //   ensureSheet이 누락 컬럼(BillingEntity 등)을 시트 끝에 추가하므로
    //   MASTER_HEADERS 순서가 아닌 시트 헤더 순서가 진실의 출처
    const actualHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowArr = actualHeaders.map(h => data[h] !== undefined ? data[h] : '');

    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, actualHeaders.length).setValues([rowArr]);
    } else {
      sheet.appendRow(rowArr);
    }

    return { ok: true, tourId: data.TourID, updated: existingRow > 0 };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

/**
 * 운행 일정 삭제
 */
function deleteSchedule(tourId) {
  try {
    if (!tourId) return { ok: false, error: 'tourId is empty' };
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Schedule');
    if (!sheet) return { ok: false, error: 'Schedule sheet not found' };
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('TourID');
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][idCol]).trim() === String(tourId).trim()) {
        sheet.deleteRow(i + 1);
        return { ok: true, deleted: tourId };
      }
    }
    return { ok: false, error: 'TourID not found' };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

/**
 * 운행 일정 상태 업데이트 (인보이스 발행/결제 시 자동 호출)
 */
function updateScheduleStatus(tourId, status, invoiceId) {
  try {
    if (!tourId) return { ok: false, error: 'tourId is empty' };
    if (SCHEDULE_STATUSES.indexOf(status) < 0) return { ok: false, error: 'Invalid status: ' + status };
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Schedule');
    if (!sheet) return { ok: false, error: 'Schedule sheet not found' };
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('TourID');
    const stCol = headers.indexOf('Status');
    const invCol = headers.indexOf('InvoiceID');
    const upCol = headers.indexOf('UpdatedAt');
    const now = new Date();
    const sydNow = Utilities.formatDate(now, 'Australia/Sydney', 'yyyy-MM-dd HH:mm:ss');

    // ★ 2026-05-23 가드: 같은 InvoiceID가 다른 TourID에 이미 있는지 사전 검사
    //   하나의 InvoiceID는 하나의 TourID에만 연결되어야 함 (data integrity)
    if (invoiceId && invCol >= 0) {
      const tgtId = String(tourId).trim();
      const tgtInv = String(invoiceId).trim();
      for (let i = 1; i < data.length; i++) {
        const rid = String(data[i][idCol]).trim();
        const riv = String(data[i][invCol]||'').trim();
        if (rid !== tgtId && riv === tgtInv) {
          // 다른 TourID가 이미 같은 InvoiceID 사용 중 → 충돌
          Logger.log('[updateScheduleStatus] CONFLICT: InvoiceID ' + tgtInv +
                     ' already used by TourID ' + rid + ' (request was for ' + tgtId + ')');
          return {
            ok: false,
            error: 'InvoiceID conflict',
            conflictMessage: 'InvoiceID ' + tgtInv + '이 이미 다른 일정(' + rid + ')에 연결되어 있습니다',
            conflictTourId: rid,
            conflictInvoiceId: tgtInv
          };
        }
      }
    }

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]).trim() === String(tourId).trim()) {
        sheet.getRange(i + 1, stCol + 1).setValue(status);
        if (invoiceId) sheet.getRange(i + 1, invCol + 1).setValue(invoiceId);
        if (upCol >= 0) sheet.getRange(i + 1, upCol + 1).setValue(sydNow);
        return { ok: true, tourId: tourId, status: status };
      }
    }
    return { ok: false, error: 'TourID not found' };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

/**
 * 자동 상태 업데이트 (매일 새벽 1시 트리거)
 */
function runScheduleStatusUpdate() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Schedule');
    if (!sheet) {
      Logger.log('Schedule sheet not found, skipping');
      return { ok: true, updated: 0 };
    }
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { ok: true, updated: 0 };
    const headers = data[0];
    const sdCol = headers.indexOf('StartDate');
    const edCol = headers.indexOf('EndDate');
    const stCol = headers.indexOf('Status');
    const upCol = headers.indexOf('UpdatedAt');
    const now = new Date();
    const today = Utilities.formatDate(now, 'Australia/Sydney', 'yyyy-MM-dd');
    const sydNow = Utilities.formatDate(now, 'Australia/Sydney', 'yyyy-MM-dd HH:mm:ss');

    // ★ 시트 셀이 Date 객체일 수도, 문자열일 수도 → 통일된 yyyy-MM-dd 추출
    function _toISODate(v) {
      if (!v && v !== 0) return '';
      if (v instanceof Date && !isNaN(v.getTime())) {
        return Utilities.formatDate(v, 'Australia/Sydney', 'yyyy-MM-dd');
      }
      const s = String(v).trim();
      // 이미 yyyy-MM-dd로 시작하면 그대로
      if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
      // dd/MM/yyyy 형식 변환
      const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
      if (m) return `${m[3]}-${m[2]}-${m[1]}`;
      return '';
    }

    let updated = 0;
    for (let i = 1; i < data.length; i++) {
      const sd = _toISODate(data[i][sdCol]);
      const ed = _toISODate(data[i][edCol]);
      const st = String(data[i][stCol]||'').trim();
      let newSt = '';
      if ((st === 'scheduled' || st === 'in_progress') && today > ed && ed) {
        newSt = 'completed';
      } else if (st === 'scheduled' && sd && ed && today >= sd && today <= ed) {
        newSt = 'in_progress';
      }
      if (newSt && newSt !== st) {
        sheet.getRange(i + 1, stCol + 1).setValue(newSt);
        if (upCol >= 0) sheet.getRange(i + 1, upCol + 1).setValue(sydNow);
        updated++;
      }
    }
    Logger.log(`runScheduleStatusUpdate: ${updated} 건 상태 변경`);
    return { ok: true, updated: updated };
  } catch (err) {
    Logger.log('runScheduleStatusUpdate error: ' + err.toString());
    return { ok: false, error: err.toString() };
  }
}

/**
 * Schedule 자동 상태 전환 트리거 등록 (한 번만)
 */
/**
 * 일회성 마이그레이션 — SUB_Txn 시트에 TourCode 컬럼 추가 + 기존 행 자동 채움
 *
 * 사용법: GAS 편집기에서 이 함수를 한 번 실행하면…
 *  1) SUB_Txn 시트의 헤더에 'TourCode' 컬럼이 InvoiceNo와 Description 사이에 삽입됨
 *     (이미 있으면 건너뜀)
 *  2) 기존 행의 Description이 'DRSUB:YYYY-MM-DD_REGO_TOURCODE' 형식이면 TourCode 자동 추출
 *  3) Description이 'PAID_TC:{tourcode}' 형식이면 그것도 TourCode 자동 채움
 *
 * 반복 실행해도 안전 (멱등). 실행 결과는 Logger에 출력됨.
 */
function migrateSubTxnAddTourCode() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('SUB_Txn');
  if (!sheet) {
    Logger.log('❌ SUB_Txn 시트가 없습니다');
    return 'SUB_Txn sheet not found';
  }

  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // 1) TourCode 컬럼 추가 (InvoiceNo 다음 위치에)
  let tcIdx = headers.indexOf('TourCode');
  if (tcIdx < 0) {
    const invIdx = headers.indexOf('InvoiceNo');
    const insertAfter = invIdx >= 0 ? invIdx + 1 : lastCol; // InvoiceNo 뒤, 없으면 맨 끝
    // insertColumnAfter는 1-based
    sheet.insertColumnAfter(insertAfter);
    sheet.getRange(1, insertAfter + 1).setValue('TourCode');
    tcIdx = insertAfter; // 0-based 인덱스로 저장
    Logger.log('✅ TourCode 컬럼 추가됨 (위치: ' + (insertAfter + 1) + ')');
  } else {
    Logger.log('ℹ️ TourCode 컬럼이 이미 존재함 (위치: ' + (tcIdx + 1) + ')');
  }

  // 2) 기존 행을 다시 읽어서 Description으로부터 TourCode 추출
  if (lastRow < 2) {
    Logger.log('ℹ️ 데이터 행 없음 - 헤더만 추가하고 종료');
    return 'header added, no data rows';
  }

  const newLastCol = sheet.getLastColumn();
  const newHeaders = sheet.getRange(1, 1, 1, newLastCol).getValues()[0];
  const tcIdxFinal = newHeaders.indexOf('TourCode');
  const descIdx = newHeaders.indexOf('Description');
  if (tcIdxFinal < 0) {
    Logger.log('❌ TourCode 컬럼 추가 실패');
    return 'TourCode column missing after insert';
  }
  if (descIdx < 0) {
    Logger.log('⚠️ Description 컬럼 없음 - 자동 채움 건너뜀');
    return 'no Description column';
  }

  const data = sheet.getRange(2, 1, lastRow - 1, newLastCol).getValues();
  let filled = 0;
  let skipped = 0;
  const drsubRE = /^DRSUB:\d{4}-\d{2}-\d{2}_[^_]*_(.+)$/;
  const paidTcRE = /^PAID_TC:(.+)$/;

  for (let i = 0; i < data.length; i++) {
    const existingTC = String(data[i][tcIdxFinal] || '').trim();
    if (existingTC) { skipped++; continue; } // 이미 채워진 행 건너뜀

    const desc = String(data[i][descIdx] || '');
    let tc = '';
    let m = desc.match(drsubRE);
    if (m && m[1]) tc = m[1].trim();
    if (!tc) {
      m = desc.match(paidTcRE);
      if (m && m[1]) tc = m[1].trim();
    }

    if (tc) {
      data[i][tcIdxFinal] = tc;
      filled++;
    }
  }

  if (filled > 0) {
    // 변경된 열만 일괄 업데이트
    const tcCol = data.map(r => [r[tcIdxFinal]]);
    sheet.getRange(2, tcIdxFinal + 1, data.length, 1).setValues(tcCol);
  }

  Logger.log('✅ 마이그레이션 완료: 채움 ' + filled + '건, 기존값 유지 ' + skipped + '건, 총 ' + data.length + '행');
  return 'Migration complete: filled=' + filled + ', skipped=' + skipped + ', total=' + data.length;
}

// ═══════════════════════════════════════════════════════════════════════════
// PAYOUT OVERRIDES — 외주 지급 자동 판단(BillingEntity) + 수동 오버라이드
// ═══════════════════════════════════════════════════════════════════════════

/**
 * 외주 지급 오버라이드 + Schedule.BillingEntity 맵 조회
 * 응답 형식:
 *   { ok: true,
 *     billingEntities: { tourCode: 'DC' | 'EG TRAVEL PTY LTD' | ... },
 *     overrides: { tourCode: { subCompanyUpper: 'INCLUDE' | 'EXCLUDE' } }
 *   }
 * Frontend는 이 두 정보로 자동/수동 제외를 판단함
 */
function getPayoutOverrides() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);

    // 1) Schedule에서 TourCode → BillingEntity 맵 추출 (1차 소스)
    const scheduleSheet = ss.getSheetByName('Schedule');
    const billingEntities = {};
    if (scheduleSheet) {
      const lastRow = scheduleSheet.getLastRow();
      const lastCol = scheduleSheet.getLastColumn();
      if (lastRow >= 2 && lastCol >= 1) {
        const headers = scheduleSheet.getRange(1, 1, 1, lastCol).getValues()[0];
        const tcIdx = headers.indexOf('TourCode');
        const beIdx = headers.indexOf('BillingEntity');
        if (tcIdx >= 0 && beIdx >= 0) {
          const data = scheduleSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
          data.forEach(row => {
            const tc = String(row[tcIdx] || '').trim();
            if (!tc) return;
            let be = String(row[beIdx] || '').trim();
            if (!be) be = 'DC';
            billingEntities[tc] = be.toUpperCase();
          });
        }
      }
    }

    // 1-b) Daily_Report에서 추가 추출 (2차 소스)
    //   Schedule에 미등록된 TourCode 또는 BillingEntity가 비어있는 경우를 보완.
    //   Daily_Report의 Billing_Entity가 명확하면 그 값을 사용.
    //   ★ Schedule에 'DC'로 명시된 경우는 덮어쓰지 않음 (의도적 설정 보호).
    //      단 Schedule에 키가 아예 없거나, 빈 값/DC인데 DR이 모두 같은 비-DC BE를 가지면 DR을 따른다.
    //      → 안전하게: Schedule에 키가 없는 경우만 DR로 보충
    try {
      const drSheet = ss.getSheetByName('Daily_Report');
      if (drSheet) {
        const drLastRow = drSheet.getLastRow();
        const drLastCol = drSheet.getLastColumn();
        if (drLastRow >= 2 && drLastCol >= 1) {
          const drHeaders = drSheet.getRange(1, 1, 1, drLastCol).getValues()[0];
          const dTC = drHeaders.indexOf('Tour_Code');
          const dBE = drHeaders.indexOf('Billing_Entity');
          if (dTC >= 0 && dBE >= 0) {
            const drData = drSheet.getRange(2, 1, drLastRow - 1, drLastCol).getValues();
            // 각 TourCode가 가지는 BE 집합 수집
            const drBEMap = {};   // tc -> Set of BE (uppercased)
            drData.forEach(row => {
              const tc = String(row[dTC] || '').trim();
              if (!tc) return;
              const be = String(row[dBE] || '').trim().toUpperCase();
              if (!be) return;
              if (!drBEMap[tc]) drBEMap[tc] = new Set();
              drBEMap[tc].add(be);
            });
            // Schedule에 키가 없는 TourCode만 보충 (Schedule 명시값 보호)
            Object.keys(drBEMap).forEach(tc => {
              if (billingEntities[tc]) return; // Schedule에 이미 있음 — 보호
              const beSet = drBEMap[tc];
              // DR에 단일 BE만 있을 때 그 값으로 채움
              if (beSet.size === 1) {
                const single = Array.from(beSet)[0];
                billingEntities[tc] = single;
              }
              // 여러 BE가 섞여 있으면 채우지 않음 (수동 확인 필요)
            });
          }
        }
      }
    } catch(drErr) {
      Logger.log('[getPayoutOverrides] DR supplement failed: ' + drErr);
    }

    // 2) PayoutOverrides 시트에서 수동 오버라이드 로드 (없으면 자동 생성)
    let overridesSheet = ss.getSheetByName('PayoutOverrides');
    if (!overridesSheet) {
      overridesSheet = ss.insertSheet('PayoutOverrides');
      overridesSheet.appendRow(MASTER_HEADERS.PayoutOverrides);
      overridesSheet.getRange(1, 1, 1, MASTER_HEADERS.PayoutOverrides.length)
        .setFontWeight('bold').setBackground('#f3f4f6');
      overridesSheet.setFrozenRows(1);
    }
    const overrides = {};
    const oLastRow = overridesSheet.getLastRow();
    const oLastCol = overridesSheet.getLastColumn();
    if (oLastRow >= 2 && oLastCol >= 1) {
      const oHeaders = overridesSheet.getRange(1, 1, 1, oLastCol).getValues()[0];
      const tcI = oHeaders.indexOf('TourCode');
      const scI = oHeaders.indexOf('SubCompany');
      const acI = oHeaders.indexOf('Action');
      if (tcI >= 0 && scI >= 0 && acI >= 0) {
        const oData = overridesSheet.getRange(2, 1, oLastRow - 1, oLastCol).getValues();
        oData.forEach(row => {
          const tc = String(row[tcI] || '').trim();
          const sc = String(row[scI] || '').trim().toUpperCase().replace(/\s+/g, ' ');
          const ac = String(row[acI] || '').trim().toUpperCase();
          if (!tc || !sc || (ac !== 'INCLUDE' && ac !== 'EXCLUDE')) return;
          if (!overrides[tc]) overrides[tc] = {};
          overrides[tc][sc] = ac;
        });
      }
    }

    return { ok: true, billingEntities: billingEntities, overrides: overrides };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

/**
 * 외주 지급 오버라이드 저장/삭제
 * data: { tourCode, subCompany, action: 'INCLUDE' | 'EXCLUDE' | 'AUTO' }
 *  - AUTO: 해당 행 삭제 (자동 판단으로 복귀)
 *  - INCLUDE/EXCLUDE: UPSERT
 */
function setPayoutOverride(data, user) {
  try {
    if (!data || !data.tourCode || !data.subCompany) {
      return { ok: false, error: 'tourCode + subCompany 필수' };
    }
    const tourCode = String(data.tourCode).trim();
    const subCompany = String(data.subCompany).trim();
    const subKey = subCompany.toUpperCase().replace(/\s+/g, ' ');
    const action = String(data.action || '').trim().toUpperCase();

    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('PayoutOverrides');
    if (!sheet) {
      sheet = ss.insertSheet('PayoutOverrides');
      sheet.appendRow(MASTER_HEADERS.PayoutOverrides);
      sheet.getRange(1, 1, 1, MASTER_HEADERS.PayoutOverrides.length)
        .setFontWeight('bold').setBackground('#f3f4f6');
      sheet.setFrozenRows(1);
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const tcI = headers.indexOf('TourCode');
    const scI = headers.indexOf('SubCompany');
    const acI = headers.indexOf('Action');
    const upI = headers.indexOf('UpdatedAt');
    const ubI = headers.indexOf('UpdatedBy');

    // 기존 행 검색
    let existingRow = -1;
    if (lastRow >= 2) {
      const data2 = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      for (let i = 0; i < data2.length; i++) {
        const t = String(data2[i][tcI] || '').trim();
        const s = String(data2[i][scI] || '').trim().toUpperCase().replace(/\s+/g, ' ');
        if (t === tourCode && s === subKey) {
          existingRow = i + 2;
          break;
        }
      }
    }

    if (action === 'AUTO') {
      // 자동 복귀 → 행 삭제
      if (existingRow > 0) {
        sheet.deleteRow(existingRow);
        return { ok: true, deleted: true };
      }
      return { ok: true, deleted: false };
    }

    if (action !== 'INCLUDE' && action !== 'EXCLUDE') {
      return { ok: false, error: 'action은 INCLUDE/EXCLUDE/AUTO 중 하나여야 함' };
    }

    const now = new Date();
    const ts = Utilities.formatDate(now, 'Australia/Sydney', "yyyy-MM-dd'T'HH:mm:ss");

    if (existingRow > 0) {
      sheet.getRange(existingRow, acI + 1).setValue(action);
      if (upI >= 0) sheet.getRange(existingRow, upI + 1).setValue(ts);
      if (ubI >= 0) sheet.getRange(existingRow, ubI + 1).setValue(user || '');
      return { ok: true, updated: true, rowIndex: existingRow };
    } else {
      const newRow = new Array(lastCol).fill('');
      if (tcI >= 0) newRow[tcI] = tourCode;
      if (scI >= 0) newRow[scI] = subCompany;
      if (acI >= 0) newRow[acI] = action;
      if (upI >= 0) newRow[upI] = ts;
      if (ubI >= 0) newRow[ubI] = user || '';
      sheet.appendRow(newRow);
      return { ok: true, inserted: true, rowIndex: sheet.getLastRow() };
    }
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

/**
 * 일회성 마이그레이션 — Schedule 시트에 BillingEntity 컬럼 백필
 *
 * 사용법: GAS 편집기에서 실행하거나 'migrate_schedule_billing_entity' 액션 호출
 *
 * 동작:
 *  1) Schedule 시트에 BillingEntity 컬럼이 없으면 추가
 *  2) 기존 행의 BillingEntity가 비어있으면 'DC'로 백필
 *
 * 반복 실행해도 안전 (멱등)
 */
function migrateScheduleBillingEntity() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'Schedule'); // ensureSheet이 누락 컬럼 자동 추가

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const beIdx = headers.indexOf('BillingEntity');
    if (beIdx < 0) {
      return { ok: false, error: 'BillingEntity 컬럼 추가 실패 — ensureSheet 점검 필요' };
    }

    if (lastRow < 2) {
      return { ok: true, filled: 0, skipped: 0, total: 0, note: 'no data rows' };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    let filled = 0;
    let skipped = 0;
    const updates = []; // { rowIndex: int, value: str }

    for (let i = 0; i < data.length; i++) {
      const current = String(data[i][beIdx] || '').trim();
      if (current) { skipped++; continue; }
      updates.push({ rowIndex: i + 2, value: 'DC' });
      filled++;
    }

    if (filled > 0) {
      // 일괄 업데이트
      updates.forEach(u => {
        sheet.getRange(u.rowIndex, beIdx + 1).setValue(u.value);
      });
    }

    Logger.log('✅ Schedule.BillingEntity 백필 완료: 채움 ' + filled + '건, 기존값 유지 ' + skipped + '건, 총 ' + data.length + '행');
    return { ok: true, filled: filled, skipped: skipped, total: data.length };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

/**
 * 일회성 정리 — BillingEntity == SubCompany 인 자동등록 DRSUB 거래 삭제
 *
 * dryRun=true (기본): 삭제 후보만 반환, 실제 삭제 안 함
 * dryRun=false: 후보를 실제 삭제 (지급된 CR이 있는 그룹은 보존)
 *
 * 안전장치:
 *  - DRSUB: prefix 행만 대상 (PAID:.. / PAID_TC:.. / 수동 등록 행은 절대 안 건드림)
 *  - 같은 TourCode + SubCompany 그룹에 CR 지급된 행이 하나라도 있으면 그 그룹 전체 보존
 *  - Schedule에 BillingEntity가 없는 TourCode는 판단 불가 → 건드리지 않음
 */
function cleanupSelfOwnedSubTxns(dryRun) {
  try {
    dryRun = (dryRun !== false);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const subSheet = ss.getSheetByName('SUB_Txn');
    if (!subSheet) return { ok: false, error: 'SUB_Txn 시트 없음' };
    const scheduleSheet = ss.getSheetByName('Schedule');
    if (!scheduleSheet) return { ok: false, error: 'Schedule 시트 없음' };

    // 1) Schedule.BillingEntity 맵 구축
    const sLastRow = scheduleSheet.getLastRow();
    const sLastCol = scheduleSheet.getLastColumn();
    const billingMap = {};
    if (sLastRow >= 2 && sLastCol >= 1) {
      const sHeaders = scheduleSheet.getRange(1, 1, 1, sLastCol).getValues()[0];
      const tcIdx = sHeaders.indexOf('TourCode');
      const beIdx = sHeaders.indexOf('BillingEntity');
      if (tcIdx >= 0 && beIdx >= 0) {
        const sData = scheduleSheet.getRange(2, 1, sLastRow - 1, sLastCol).getValues();
        sData.forEach(row => {
          const tc = String(row[tcIdx] || '').trim();
          if (!tc) return;
          let be = String(row[beIdx] || '').trim();
          if (!be) be = 'DC';
          billingMap[tc] = be.toUpperCase().replace(/\s+/g, ' ');
        });
      }
    }

    // 2) SUB_Txn 스캔
    const lastRow = subSheet.getLastRow();
    const lastCol = subSheet.getLastColumn();
    if (lastRow < 2) return { ok: true, dryRun: dryRun, candidates: [], blocked: [], candidateCount: 0, blockedCount: 0 };

    const headers = subSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const scI = headers.indexOf('SubCompany');
    const tcI = headers.indexOf('TourCode');
    const dcI = headers.indexOf('Description');
    const drI = headers.indexOf('DR');
    const crI = headers.indexOf('CR');
    const dtI = headers.indexOf('Date');
    if (scI < 0 || dcI < 0) return { ok: false, error: 'SubCompany 또는 Description 컬럼 없음' };

    const data = subSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const candidates = [];          // 삭제 후보
    const groupHasPayment = {};     // key = tc + '|' + scKey → CR 있으면 true

    data.forEach((row, i) => {
      const rowIndex = i + 2;
      const sc = String(row[scI] || '').trim();
      const scKey = sc.toUpperCase().replace(/\s+/g, ' ');
      const desc = String(row[dcI] || '');
      const dr = Number(row[drI] || 0);
      const cr = Number(row[crI] || 0);

      // TourCode 추출 (시트 컬럼 우선, Description fallback)
      let tc = tcI >= 0 ? String(row[tcI] || '').trim() : '';
      if (!tc) {
        const m = desc.match(/^DRSUB:\d{4}-\d{2}-\d{2}_[^_]+_(.+)$/);
        if (m && m[1]) tc = m[1].trim();
      }
      if (!tc || !sc) return;

      const key = tc + '|' + scKey;
      if (cr > 0) groupHasPayment[key] = true;

      if (dr <= 0) return;
      if (!desc.startsWith('DRSUB:')) return;

      const be = billingMap[tc];
      if (!be) return; // Schedule에 없으면 건드리지 않음

      if (be === scKey) {
        // BillingEntity == SubCompany → 자기 차로 자기 손님 운행 → 삭제 후보
        candidates.push({
          rowIndex: rowIndex,
          tourCode: tc,
          subCompany: sc,
          dr: dr,
          date: dtI >= 0 ? String(row[dtI] || '') : '',
          desc: desc
        });
      }
    });

    // CR 있는 그룹의 후보 제거 (이미 일부 지급된 그룹은 데이터 보존)
    const safe = candidates.filter(c => {
      const key = c.tourCode + '|' + c.subCompany.toUpperCase().replace(/\s+/g, ' ');
      return !groupHasPayment[key];
    });
    const blocked = candidates.filter(c => {
      const key = c.tourCode + '|' + c.subCompany.toUpperCase().replace(/\s+/g, ' ');
      return groupHasPayment[key];
    });

    if (dryRun) {
      // 합계도 같이 반환
      const totalDR = safe.reduce((s, c) => s + Number(c.dr || 0), 0);
      return {
        ok: true,
        dryRun: true,
        candidates: safe,
        blocked: blocked,
        candidateCount: safe.length,
        blockedCount: blocked.length,
        totalDR: totalDR
      };
    }

    // 실제 삭제 — 아래에서 위로 (인덱스 안 꼬임)
    const sortedRows = safe.map(c => c.rowIndex).sort((a, b) => b - a);
    sortedRows.forEach(r => subSheet.deleteRow(r));

    return {
      ok: true,
      dryRun: false,
      deleted: sortedRows.length,
      blocked: blocked.length,
      deletedRows: sortedRows.length
    };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

// ───────────────────────────────────────────────────────────────────────
// 🩺 cleanupSelfOwnedSubTxns 결과를 Logger에 출력하는 wrapper들
// ───────────────────────────────────────────────────────────────────────
function cleanupSelfOwnedSubTxns_preview() {
  const r = cleanupSelfOwnedSubTxns(true);
  const log = [];
  log.push('═══ cleanupSelfOwnedSubTxns [DRY RUN] ═══');
  if (!r.ok) {
    log.push('❌ 실패: ' + r.error);
    Logger.log(log.join('\n'));
    return r;
  }
  log.push('삭제 후보: ' + r.candidateCount + '건 (합계 DR $' + (r.totalDR||0).toLocaleString() + ')');
  log.push('보존 (CR 있어 안 건드림): ' + r.blockedCount + '건');
  log.push('');
  if (r.candidates && r.candidates.length) {
    log.push('── 삭제 후보 ──');
    r.candidates.forEach(c => {
      log.push('  row ' + c.rowIndex + ' | ' + c.date + ' | ' + c.subCompany + ' | TC=' + c.tourCode + ' | DR=$' + c.dr + ' | ' + c.desc);
    });
  }
  if (r.blocked && r.blocked.length) {
    log.push('');
    log.push('── 보존 (이미 CR 지급된 그룹) ──');
    r.blocked.forEach(c => {
      log.push('  row ' + c.rowIndex + ' | ' + c.date + ' | ' + c.subCompany + ' | TC=' + c.tourCode + ' | DR=$' + c.dr);
    });
  }
  log.push('');
  log.push('확정 삭제하려면: cleanupSelfOwnedSubTxns_commit() 실행');
  Logger.log(log.join('\n'));
  return r;
}

function cleanupSelfOwnedSubTxns_commit() {
  const r = cleanupSelfOwnedSubTxns(false);
  const log = [];
  log.push('═══ cleanupSelfOwnedSubTxns [COMMIT] ═══');
  if (!r.ok) {
    log.push('❌ 실패: ' + r.error);
  } else {
    log.push('✅ 삭제 완료: ' + (r.deleted || 0) + '행');
    log.push('보존 (CR 있어 안 건드림): ' + (r.blocked || 0) + '건');
  }
  Logger.log(log.join('\n'));
  return r;
}

// 진단: 18건 q4 운행이 Schedule에 등록되어 있는지 확인 (BillingEntity 기준)
function diagEGq4ScheduleCoverage() {
  const log = [];
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    // SUB_Txn에서 EG TRAVEL의 DRSUB 행 TourCode 추출
    const subSheet = ss.getSheetByName('SUB_Txn');
    const subData = subSheet.getDataRange().getValues();
    const subHeaders = subData[0];
    const scI = subHeaders.indexOf('SubCompany');
    const tcI = subHeaders.indexOf('TourCode');
    const dcI = subHeaders.indexOf('Description');
    const drI = subHeaders.indexOf('DR');
    const crI = subHeaders.indexOf('CR');
    const egTourCodes = new Set();
    const egTcCrPaid = {};
    for (let i = 1; i < subData.length; i++) {
      const row = subData[i];
      const sc = String(row[scI]||'').toUpperCase();
      if (sc.indexOf('EG TRAVEL') < 0) continue;
      const desc = String(row[dcI]||'');
      let tc = tcI >= 0 ? String(row[tcI]||'').trim() : '';
      if (!tc) {
        const m = desc.match(/^DRSUB:\d{4}-\d{2}-\d{2}_[^_]+_(.+)$/);
        if (m) tc = m[1].trim();
      }
      if (!tc) continue;
      if (desc.startsWith('DRSUB:') && Number(row[drI]||0) > 0) egTourCodes.add(tc);
      if (Number(row[crI]||0) > 0) egTcCrPaid[tc] = (egTcCrPaid[tc]||0) + Number(row[crI]);
    }
    // Schedule.BillingEntity 매핑
    const schSheet = ss.getSheetByName('Schedule');
    const schData = schSheet ? schSheet.getDataRange().getValues() : [];
    const schMap = {};
    if (schData.length > 1) {
      const h = schData[0];
      const tIdx = h.indexOf('TourCode');
      const bIdx = h.indexOf('BillingEntity');
      for (let i = 1; i < schData.length; i++) {
        const tc = String(schData[i][tIdx]||'').trim();
        const be = String(schData[i][bIdx]||'').trim();
        if (tc) schMap[tc] = be || '(빈 값)';
      }
    }
    log.push('═══ EG TRAVEL DRSUB TourCodes vs Schedule.BillingEntity ═══');
    log.push('EG TRAVEL DRSUB 자동 등록 TourCode: ' + egTourCodes.size + '개');
    log.push('');
    log.push('TC | Schedule.BillingEntity | CR 지급 여부 | 정리 가능?');
    log.push('---');
    let canClean = 0, cantClean = 0, hasCR = 0, notInSched = 0;
    Array.from(egTourCodes).sort().forEach(tc => {
      const be = schMap[tc];
      const crSum = egTcCrPaid[tc] || 0;
      let status;
      if (!be) { status = '⚠️ Schedule에 없음 → 건드릴 수 없음'; notInSched++; }
      else if (crSum > 0) { status = '🔒 CR $' + crSum + ' 있어 보존'; hasCR++; }
      else if (be.toUpperCase().indexOf('EG TRAVEL') >= 0) { status = '✅ 정리 가능 (BE=' + be + ' == EG TRAVEL)'; canClean++; }
      else { status = '⏭ 정리 안 됨 (BE=' + be + ' ≠ EG TRAVEL)'; cantClean++; }
      log.push(tc + ' | ' + (be||'(없음)') + ' | $' + crSum + ' | ' + status);
    });
    log.push('---');
    log.push('정리 가능: ' + canClean + ' | CR로 보존: ' + hasCR + ' | Schedule에 없음: ' + notInSched + ' | BE 불일치: ' + cantClean);
    Logger.log(log.join('\n'));
    return { canClean, hasCR, notInSched, cantClean };
  } catch (e) {
    Logger.log('error: ' + e);
    return { error: String(e) };
  }
}

function setupScheduleTrigger() {
  removeScheduleTrigger();
  ScriptApp.newTrigger('runScheduleStatusUpdate')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .inTimezone('Australia/Sydney')
    .create();
  Logger.log('✅ 운행 일정 자동 상태 전환 트리거 등록: 매일 새벽 1시 (Sydney)');
  return 'Schedule trigger created.';
}

function removeScheduleTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'runScheduleStatusUpdate') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  Logger.log('Removed ' + removed + ' schedule trigger(s).');
  return removed;
}

/**
 * 진단: SUB 인보이스 → SUB_Txn 동기화 상태 점검
 *
 * 목적: Sub 인보이스가 발행됐는데 잔액 화면에 안 나오는 원인 파악
 *  - Invoices 시트의 Source='SUB' 행과 SUB_Txn의 SUBINV: 행을 1:1 대조
 *  - sync 누락 / 중복 / 금액 불일치 / SubCompany 누락 등 케이스 식별
 *
 * 옵션: subCompanyFilter (선택) — 특정 SUB 업체만 점검
 * 사용법:
 *   - 전체 점검: diagnoseSubInvoiceSync()
 *   - 특정 업체: diagnoseSubInvoiceSync('Sydney Edu Tours P/L')
 *
 * 반환: 진단 결과 객체 (Logger.log로 사람이 읽기 좋은 형식도 출력)
 */
function diagnoseSubInvoiceSync(subCompanyFilter) {
  const result = {
    ok: true,
    summary: {},
    issues: [],
    invoices: [],
    matches: []
  };
  const log = [];
  const filter = subCompanyFilter ? String(subCompanyFilter).trim() : '';
  log.push('═══ SUB 인보이스 ↔ SUB_Txn 동기화 진단 ═══');
  if (filter) log.push('필터: SubCompany = "' + filter + '"');

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);

    // ── 1) Invoices 시트에서 SUB 인보이스 추출 ──
    const invSheet = ss.getSheetByName('Invoices');
    if (!invSheet) {
      result.ok = false;
      log.push('❌ Invoices 시트 없음');
      Logger.log(log.join('\n'));
      return result;
    }
    const invLastRow = invSheet.getLastRow();
    const invLastCol = invSheet.getLastColumn();
    const subInvs = []; // { rowIndex, invNum, subCompany, grandTotal, source, status, issuedDate }
    if (invLastRow >= 2 && invLastCol >= 1) {
      const invHeaders = invSheet.getRange(1, 1, 1, invLastCol).getValues()[0];
      const idx = {};
      ['InvNumber','Source','SubCompany','GrandTotal','Status','IssuedDate','PeriodFrom','PeriodTo'].forEach(h => {
        idx[h] = invHeaders.indexOf(h);
      });
      const invData = invSheet.getRange(2, 1, invLastRow - 1, invLastCol).getValues();
      invData.forEach((row, i) => {
        const source = idx.Source >= 0 ? String(row[idx.Source] || '').trim() : '';
        const invNum = idx.InvNumber >= 0 ? String(row[idx.InvNumber] || '').trim() : '';
        // SUB 인보이스 식별: Source='SUB' 또는 패턴 (INV-가 아니고 알파벳1~3자+숫자)
        const isSub = source === 'SUB' ||
          (invNum && !/^INV-/i.test(invNum) && /^[A-Z]{1,3}\d+$/.test(invNum));
        if (!isSub) return;
        const subCo = idx.SubCompany >= 0 ? String(row[idx.SubCompany] || '').trim() : '';
        if (filter && subCo !== filter) return;
        subInvs.push({
          rowIndex: i + 2,
          invNum: invNum,
          subCompany: subCo,
          grandTotal: idx.GrandTotal >= 0 ? Number(row[idx.GrandTotal] || 0) : 0,
          source: source,
          status: idx.Status >= 0 ? String(row[idx.Status] || '').trim() : '',
          issuedDate: idx.IssuedDate >= 0 ? String(row[idx.IssuedDate] || '').trim() : '',
          periodFrom: idx.PeriodFrom >= 0 ? String(row[idx.PeriodFrom] || '').trim() : '',
          periodTo: idx.PeriodTo >= 0 ? String(row[idx.PeriodTo] || '').trim() : ''
        });
      });
    }
    log.push('\n--- 1) Invoices 시트 SUB 인보이스 ---');
    log.push('총 ' + subInvs.length + '건' + (filter ? ' (필터 적용)' : ''));

    // ── 2) SUB_Txn에서 SUBINV: 거래 추출 ──
    const subSheet = ss.getSheetByName('SUB_Txn');
    if (!subSheet) {
      result.ok = false;
      log.push('❌ SUB_Txn 시트 없음');
      Logger.log(log.join('\n'));
      return result;
    }
    const sLastRow = subSheet.getLastRow();
    const sLastCol = subSheet.getLastColumn();
    const subTxns = []; // SUBINV: 거래만
    const allTxns = []; // 같은 SubCompany 전체 거래 (CR 합계용)
    if (sLastRow >= 2 && sLastCol >= 1) {
      const sHeaders = subSheet.getRange(1, 1, 1, sLastCol).getValues()[0];
      const idx2 = {};
      ['SubCompany','Description','DR','CR','InvoiceNo','Date'].forEach(h => {
        idx2[h] = sHeaders.indexOf(h);
      });
      const sData = subSheet.getRange(2, 1, sLastRow - 1, sLastCol).getValues();
      sData.forEach((row, i) => {
        const sc = idx2.SubCompany >= 0 ? String(row[idx2.SubCompany] || '').trim() : '';
        if (filter && sc !== filter) return;
        const desc = idx2.Description >= 0 ? String(row[idx2.Description] || '').trim() : '';
        const dr = idx2.DR >= 0 ? Number(row[idx2.DR] || 0) : 0;
        const cr = idx2.CR >= 0 ? Number(row[idx2.CR] || 0) : 0;
        const invNo = idx2.InvoiceNo >= 0 ? String(row[idx2.InvoiceNo] || '').trim() : '';
        const tx = {
          rowIndex: i + 2, subCompany: sc, description: desc,
          dr: dr, cr: cr, invoiceNo: invNo,
          date: idx2.Date >= 0 ? String(row[idx2.Date] || '').trim() : ''
        };
        allTxns.push(tx);
        if (desc.indexOf('SUBINV:') === 0) {
          subTxns.push(tx);
        }
      });
    }
    log.push('\n--- 2) SUB_Txn 시트 SUBINV: 거래 ---');
    log.push('총 ' + subTxns.length + '건' + (filter ? ' (필터 적용)' : ''));

    // ── 3) 대조 분석 ──
    log.push('\n--- 3) 인보이스 vs SUB_Txn 대조 ---');
    const txnByInvNum = {};
    subTxns.forEach(t => {
      // SUBINV:invNum 또는 InvoiceNo 칼럼으로 매칭
      const m = String(t.description).match(/^SUBINV:(.+)$/);
      const key = m ? m[1].trim() : (t.invoiceNo || '');
      if (!key) return;
      if (!txnByInvNum[key]) txnByInvNum[key] = [];
      txnByInvNum[key].push(t);
    });

    let missing = 0, duplicate = 0, mismatch = 0, ok = 0, noSubCo = 0;
    subInvs.forEach(inv => {
      const item = {
        invNum: inv.invNum,
        subCompany: inv.subCompany,
        grandTotal: inv.grandTotal,
        status: inv.status,
        issuedDate: inv.issuedDate,
        matched: false,
        txns: []
      };

      // 진단: SubCompany 누락
      if (!inv.subCompany) {
        item.issue = 'SubCompany 비어있음 (sync 대상 제외)';
        noSubCo++;
        result.issues.push(item);
        log.push('⚠️ ' + inv.invNum + ' — SubCompany 비어있음 (row ' + inv.rowIndex + ')');
        result.invoices.push(item);
        return;
      }

      // 진단: 금액 0
      if (inv.grandTotal <= 0) {
        item.issue = 'GrandTotal 0 이하 (sync 대상 제외)';
        result.issues.push(item);
        log.push('⚠️ ' + inv.invNum + ' — GrandTotal=' + inv.grandTotal + ' (sync 안 됨)');
        result.invoices.push(item);
        return;
      }

      const matchedTxns = txnByInvNum[inv.invNum] || [];
      item.txns = matchedTxns;

      if (matchedTxns.length === 0) {
        item.issue = 'SUB_Txn에 sync 안 됨';
        missing++;
        result.issues.push(item);
        log.push('❌ ' + inv.invNum + ' (' + inv.subCompany + ') $' + inv.grandTotal + ' — SUB_Txn에 sync 안 됨');
      } else if (matchedTxns.length > 1) {
        item.issue = '중복 ' + matchedTxns.length + '건';
        duplicate++;
        result.issues.push(item);
        log.push('⚠️ ' + inv.invNum + ' — SUB_Txn에 ' + matchedTxns.length + '건 중복');
      } else {
        const t = matchedTxns[0];
        // 금액 일치 확인
        if (Math.abs(t.dr - inv.grandTotal) > 0.01) {
          item.issue = '금액 불일치: 인보이스 $' + inv.grandTotal + ' vs SUB_Txn DR $' + t.dr;
          mismatch++;
          result.issues.push(item);
          log.push('⚠️ ' + inv.invNum + ' — 금액 불일치 $' + inv.grandTotal + ' vs DR $' + t.dr);
        } else {
          item.matched = true;
          ok++;
          result.matches.push(item);
        }
      }
      result.invoices.push(item);
    });

    // ── 4) 같은 SubCompany 잔액 시뮬레이션 ──
    log.push('\n--- 4) SubCompany별 잔액 시뮬레이션 ---');
    const balByCompany = {};
    allTxns.forEach(t => {
      const sc = t.subCompany || '(없음)';
      if (!balByCompany[sc]) balByCompany[sc] = { dr: 0, cr: 0, drCount: 0, crCount: 0 };
      balByCompany[sc].dr += t.dr;
      balByCompany[sc].cr += t.cr;
      if (t.dr > 0) balByCompany[sc].drCount++;
      if (t.cr > 0) balByCompany[sc].crCount++;
    });
    Object.keys(balByCompany).sort().forEach(sc => {
      const b = balByCompany[sc];
      log.push('  ' + sc + ' → DR $' + b.dr.toFixed(2) + ' (' + b.drCount + '건) / CR $' + b.cr.toFixed(2) + ' (' + b.crCount + '건) = $' + (b.dr - b.cr).toFixed(2));
    });

    result.summary = {
      totalInvoices: subInvs.length,
      totalSubTxns: subTxns.length,
      ok: ok,
      missing: missing,
      duplicate: duplicate,
      mismatch: mismatch,
      noSubCompany: noSubCo
    };

    log.push('\n═══ 요약 ═══');
    log.push('  ✅ 정상: ' + ok + '건');
    log.push('  ❌ Sync 누락: ' + missing + '건');
    log.push('  ⚠️ 중복: ' + duplicate + '건');
    log.push('  ⚠️ 금액 불일치: ' + mismatch + '건');
    log.push('  ⚠️ SubCompany 누락: ' + noSubCo + '건');
    log.push('');

    Logger.log(log.join('\n'));
    return result;
  } catch (e) {
    result.ok = false;
    result.error = String(e);
    Logger.log('진단 오류: ' + e);
    return result;
  }
}

/**
 * 누락된 SUB 인보이스를 SUB_Txn에 재동기화
 *
 * diagnoseSubInvoiceSync()에서 missing 으로 식별된 인보이스를 SUB_Txn에 등록
 * 안전장치: SubCompany 비어있거나 GrandTotal=0 이면 스킵
 * 멱등: 이미 SUB_Txn에 있으면 추가 등록 안 함
 *
 * 사용법:
 *   - 전체: resyncMissingSubInvoices()
 *   - 특정 업체: resyncMissingSubInvoices('Sydney Edu Tours P/L')
 */
function resyncMissingSubInvoices(subCompanyFilter) {
  const result = { ok: true, registered: 0, skipped: 0, errors: [] };
  try {
    const diag = diagnoseSubInvoiceSync(subCompanyFilter);
    if (!diag.ok) return { ok: false, error: diag.error || 'diagnose failed' };

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const subSheet = ss.getSheetByName('SUB_Txn');
    if (!subSheet) return { ok: false, error: 'SUB_Txn 시트 없음' };

    const sLastCol = subSheet.getLastColumn();
    const headers = subSheet.getRange(1, 1, 1, sLastCol).getValues()[0];

    // missing 건만 추출
    const missing = diag.invoices.filter(i => i.issue === 'SUB_Txn에 sync 안 됨');
    Logger.log('재동기화 대상: ' + missing.length + '건');

    missing.forEach(inv => {
      try {
        const issuedDate = (inv.issuedDate || '').slice(0, 10);
        const newRow = {};
        newRow.SubCompany = inv.subCompany;
        newRow.Category = 'Outsourcing';
        newRow.Date = issuedDate;
        newRow.InvoiceNo = inv.invNum;
        newRow.Description = 'SUBINV:' + inv.invNum;
        newRow.DR = inv.grandTotal;
        newRow.CR = 0;
        newRow.Remark = inv.invNum + ' (재동기화)';

        const rowArr = headers.map(h => newRow[h] !== undefined ? newRow[h] : '');
        subSheet.appendRow(rowArr);
        result.registered++;
        Logger.log('  ✅ ' + inv.invNum + ' 등록됨');
      } catch (e) {
        result.errors.push({ invNum: inv.invNum, error: String(e) });
        Logger.log('  ❌ ' + inv.invNum + ' 실패: ' + e);
      }
    });

    result.skipped = diag.invoices.length - missing.length;
    Logger.log('\n재동기화 완료: 등록 ' + result.registered + '건, 스킵 ' + result.skipped + '건, 오류 ' + result.errors.length + '건');
    return result;
  } catch (e) {
    result.ok = false;
    result.error = String(e);
    return result;
  }
}

/**
 * 진단: Schedule 시트 + 트리거 상태를 한 번에 점검
 * Apps Script 에디터에서 직접 실행 → Logger 확인
 */
function diagnoseScheduleSystem() {
  const log = [];
  log.push('═══ 운행 일정 시스템 진단 ═══');

  // 1. 트리거 상태
  log.push('\n--- 1) 자동 트리거 등록 상태 ---');
  const triggers = ScriptApp.getProjectTriggers();
  const scheduleTriggers = triggers.filter(t => t.getHandlerFunction() === 'runScheduleStatusUpdate');
  if (scheduleTriggers.length === 0) {
    log.push('❌ 등록된 트리거 없음. setupScheduleTrigger() 함수를 실행해야 합니다.');
  } else {
    scheduleTriggers.forEach(t => {
      log.push('✅ 트리거 등록됨: ' + t.getEventType() + ' (' + t.getTriggerSource() + ')');
    });
  }

  // 2. 오늘 날짜 (Sydney 기준)
  const now = new Date();
  const today = Utilities.formatDate(now, 'Australia/Sydney', 'yyyy-MM-dd');
  log.push('\n--- 2) 오늘 날짜 (Sydney) ---');
  log.push('today = ' + today);

  // 3. Schedule 시트 데이터
  log.push('\n--- 3) Schedule 시트 데이터 ---');
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Schedule');
  if (!sheet) {
    log.push('❌ Schedule 시트 없음');
    Logger.log(log.join('\n'));
    return log.join('\n');
  }
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    log.push('(데이터 없음)');
    Logger.log(log.join('\n'));
    return log.join('\n');
  }

  const headers = data[0];
  const idCol = headers.indexOf('TourID');
  const tcCol = headers.indexOf('TourCode');
  const sdCol = headers.indexOf('StartDate');
  const edCol = headers.indexOf('EndDate');
  const stCol = headers.indexOf('Status');
  const agCol = headers.indexOf('Agency');

  function _toISODate(v) {
    if (!v && v !== 0) return '';
    if (v instanceof Date && !isNaN(v.getTime())) {
      return Utilities.formatDate(v, 'Australia/Sydney', 'yyyy-MM-dd');
    }
    const s = String(v).trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
    if (m) return `${m[3]}-${m[2]}-${m[1]}`;
    return '';
  }

  // 진행중이거나 예정인 일정만 표시
  log.push('상태가 scheduled / in_progress 인 일정:');
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const st = String(data[i][stCol]||'').trim();
    if (st !== 'scheduled' && st !== 'in_progress') continue;

    const rawSd = data[i][sdCol];
    const rawEd = data[i][edCol];
    const sd = _toISODate(rawSd);
    const ed = _toISODate(rawEd);
    const id = data[i][idCol] || '';
    const tc = data[i][tcCol] || '';
    const ag = data[i][agCol] || '';

    let suggestion = '';
    if (today > ed && ed) suggestion = ' → completed (종료일 지남)';
    else if (sd && ed && today >= sd && today <= ed) suggestion = ' → in_progress (오늘 일정 중)';
    else if (sd && today < sd) suggestion = ' (시작일 미도래, scheduled 유지)';

    // 원본 셀 타입 함께 출력 (디버깅용)
    const sdType = rawSd instanceof Date ? 'Date' : typeof rawSd;
    log.push(`[${st}] ${id} (${tc}) | ${ag} | ${sd} ~ ${ed} (raw start: ${sdType} "${rawSd}")${suggestion}`);
    count++;
    if (count > 20) { log.push('... (이하 생략)'); break; }
  }
  if (count === 0) log.push('(scheduled/in_progress 일정 없음)');

  // 4. 시뮬레이션 — 지금 trigger 돌리면 몇 건 바뀔지
  log.push('\n--- 4) 만약 지금 runScheduleStatusUpdate를 실행하면 ---');
  let wouldUpdate = 0;
  for (let i = 1; i < data.length; i++) {
    const sd = _toISODate(data[i][sdCol]);
    const ed = _toISODate(data[i][edCol]);
    const st = String(data[i][stCol]||'').trim();
    if ((st === 'scheduled' || st === 'in_progress') && today > ed && ed) wouldUpdate++;
    else if (st === 'scheduled' && sd && ed && today >= sd && today <= ed) wouldUpdate++;
  }
  log.push(wouldUpdate + ' 건이 상태 변경 대상');

  Logger.log(log.join('\n'));
  return log.join('\n');
}

/**
 * 강화된 자동 상태 전환 — DR 매칭 기반 보너스 룰 추가
 *
 * 기존 룰:
 *   - StartDate <= today <= EndDate → in_progress
 *   - today > EndDate → completed
 *
 * 새 룰 (DR-driven):
 *   - 일정 기간 내 Daily_Report에 매칭 row가 1건이라도 있으면 → in_progress
 *     (시작일 도래 안 했어도 드라이버가 일찍 출근한 케이스 처리)
 *
 * 매칭 키: Date in [StartDate, EndDate] AND (Agency match OR TourCode match)
 */
function runScheduleStatusUpdateV2() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Schedule');
    if (!sheet) {
      Logger.log('Schedule sheet not found, skipping');
      return { ok: true, updated: 0 };
    }
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { ok: true, updated: 0 };

    const headers = data[0];
    const sdCol = headers.indexOf('StartDate');
    const edCol = headers.indexOf('EndDate');
    const stCol = headers.indexOf('Status');
    const upCol = headers.indexOf('UpdatedAt');
    const agCol = headers.indexOf('Agency');
    const tcCol = headers.indexOf('TourCode');

    const now = new Date();
    const today = Utilities.formatDate(now, 'Australia/Sydney', 'yyyy-MM-dd');
    const sydNow = Utilities.formatDate(now, 'Australia/Sydney', 'yyyy-MM-dd HH:mm:ss');

    // 시트 셀이 Date 객체/문자열 모두 처리
    function _toISODate(v) {
      if (!v && v !== 0) return '';
      if (v instanceof Date && !isNaN(v.getTime())) {
        return Utilities.formatDate(v, 'Australia/Sydney', 'yyyy-MM-dd');
      }
      const s = String(v).trim();
      if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
      const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
      if (m) return `${m[3]}-${m[2]}-${m[1]}`;
      return '';
    }

    // Daily_Report 한 번만 로드해서 메모리에서 처리
    const drSheet = ss.getSheetByName('Daily_Report');
    let drRows = [];
    if (drSheet && drSheet.getLastRow() > 1) {
      const drData = drSheet.getDataRange().getValues();
      const drHeaders = drData[0];
      const drDateCol = drHeaders.indexOf('Date');
      const drAgCol = drHeaders.indexOf('Agency');
      const drTcCol = drHeaders.indexOf('Tour_Code');
      drRows = drData.slice(1).map(r => ({
        date: _toISODate(r[drDateCol]),
        agency: String(r[drAgCol]||'').trim(),
        tourCode: String(r[drTcCol]||'').trim()
      })).filter(r => r.date);
    }

    function hasMatchingDR(sd, ed, agency, tourCode) {
      if (!sd || !ed) return false;
      const ag = String(agency||'').trim();
      const tc = String(tourCode||'').trim();
      return drRows.some(dr => {
        if (dr.date < sd || dr.date > ed) return false;
        if (tc && dr.tourCode === tc) return true;
        if (ag && dr.agency === ag) return true;
        return false;
      });
    }

    let updated = 0;
    const updateLog = [];
    for (let i = 1; i < data.length; i++) {
      const sd = _toISODate(data[i][sdCol]);
      const ed = _toISODate(data[i][edCol]);
      const st = String(data[i][stCol]||'').trim();
      const ag = data[i][agCol] || '';
      const tc = tcCol >= 0 ? (data[i][tcCol] || '') : '';
      let newSt = '';

      // 룰 1: 종료일 지남 → completed
      if ((st === 'scheduled' || st === 'in_progress') && today > ed && ed) {
        newSt = 'completed';
      }
      // 룰 2: 오늘이 일정 기간 내 → in_progress
      else if (st === 'scheduled' && sd && ed && today >= sd && today <= ed) {
        newSt = 'in_progress';
      }
      // 룰 3 (NEW): DR이 매칭되면 시작일 전이라도 → in_progress
      else if (st === 'scheduled' && hasMatchingDR(sd, ed, ag, tc)) {
        newSt = 'in_progress';
      }

      if (newSt && newSt !== st) {
        sheet.getRange(i + 1, stCol + 1).setValue(newSt);
        if (upCol >= 0) sheet.getRange(i + 1, upCol + 1).setValue(sydNow);
        updated++;
        updateLog.push(`Row ${i+1}: ${st} → ${newSt} (${ag} ${sd}~${ed})`);
      }
    }
    Logger.log(`runScheduleStatusUpdateV2: ${updated} 건 상태 변경`);
    if (updateLog.length) Logger.log(updateLog.join('\n'));
    return { ok: true, updated: updated, details: updateLog };
  } catch (err) {
    Logger.log('runScheduleStatusUpdateV2 error: ' + err.toString());
    return { ok: false, error: err.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// ONE-TIME MIGRATION: 인보이스 번호 변경 (옵션 C — 시스템을 PDF에 맞춤)
// ═══════════════════════════════════════════════════════════════════════════
//
// 사용법:
//   1) Apps Script 에디터에서 이 파일을 열고
//   2) 함수 선택 드롭다운에서 'fixInvoiceNumber_001to002' 선택
//   3) ▶ Run 클릭 → 실행 권한 승인
//   4) Logger 로그(보기 → 실행)에서 결과 확인
//   5) 실행 후 이 함수는 다시 실행하지 말 것 (멱등성 가드 있음)
//
// 작동:
//   - Invoices 시트의 InvNumber 'INV-202605-001' → 'INV-202605-002'
//   - Agency_Txn 시트의 InvoiceID 'INV-202605-001' → 'INV-202605-002'
//   - Agency_Txn의 Remark에 포함된 'INV-202605-001' 문자열도 모두 치환
//   - 002가 이미 존재하면 충돌 방지를 위해 중단 (안전 가드)
//
function fixInvoiceNumber_001to002() {
  const OLD_NUM = 'INV-202605-001';
  const NEW_NUM = 'INV-202605-002';
  return _renameInvoiceNumber(OLD_NUM, NEW_NUM);
}

function _renameInvoiceNumber(OLD_NUM, NEW_NUM) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const log = [];
  log.push(`▶ 마이그레이션 시작: ${OLD_NUM} → ${NEW_NUM}`);

  // ─── 1. Invoices 시트 ───
  const invSheet = ss.getSheetByName('Invoices');
  if (!invSheet) {
    return { ok: false, error: 'Invoices 시트 없음', log: log.join('\n') };
  }
  const invData = invSheet.getDataRange().getValues();
  const invHeaders = invData[0];
  const invNumCol = invHeaders.indexOf('InvNumber');
  if (invNumCol < 0) {
    return { ok: false, error: 'InvNumber 열 없음', log: log.join('\n') };
  }

  // 안전 가드: NEW_NUM이 이미 존재하면 충돌
  let oldRow = -1;
  let newExists = false;
  for (let i = 1; i < invData.length; i++) {
    const v = String(invData[i][invNumCol]).trim();
    if (v === OLD_NUM) oldRow = i + 1; // 1-based row
    if (v === NEW_NUM) newExists = true;
  }
  if (oldRow < 0) {
    log.push(`⚠️ Invoices 시트에 ${OLD_NUM}이 없음 — 이미 변경됐거나 삭제됨. 중단.`);
    Logger.log(log.join('\n'));
    return { ok: false, error: 'OLD_NUM not found', log: log.join('\n') };
  }
  if (newExists) {
    log.push(`❌ 충돌: Invoices 시트에 ${NEW_NUM}이 이미 존재함. 중단.`);
    Logger.log(log.join('\n'));
    return { ok: false, error: 'NEW_NUM already exists', log: log.join('\n') };
  }

  // 실제 변경
  invSheet.getRange(oldRow, invNumCol + 1).setValue(NEW_NUM);
  log.push(`✅ Invoices: row ${oldRow} InvNumber ${OLD_NUM} → ${NEW_NUM}`);

  // ─── 2. Agency_Txn 시트 ───
  const txnSheet = ss.getSheetByName('Agency_Txn');
  if (!txnSheet) {
    log.push(`⚠️ Agency_Txn 시트 없음 — 스킵.`);
    Logger.log(log.join('\n'));
    return { ok: true, log: log.join('\n') };
  }
  const txnData = txnSheet.getDataRange().getValues();
  const txnHeaders = txnData[0];
  const invIdCol = txnHeaders.indexOf('InvoiceID');
  const remarkCol = txnHeaders.indexOf('Remark');

  let txnUpdated = 0;
  for (let i = 1; i < txnData.length; i++) {
    let rowChanged = false;
    // InvoiceID 정확 매칭
    if (invIdCol >= 0 && String(txnData[i][invIdCol]).trim() === OLD_NUM) {
      txnSheet.getRange(i + 1, invIdCol + 1).setValue(NEW_NUM);
      rowChanged = true;
    }
    // Remark에 포함된 OLD_NUM 문자열 치환 (예: "전액결제 완료 (INV-202605-001)")
    if (remarkCol >= 0) {
      const remark = String(txnData[i][remarkCol] || '');
      if (remark.indexOf(OLD_NUM) >= 0) {
        const newRemark = remark.split(OLD_NUM).join(NEW_NUM);
        txnSheet.getRange(i + 1, remarkCol + 1).setValue(newRemark);
        rowChanged = true;
      }
    }
    if (rowChanged) txnUpdated++;
  }
  log.push(`✅ Agency_Txn: ${txnUpdated}개 행 갱신 (InvoiceID + Remark)`);

  log.push(`▶ 마이그레이션 완료.`);
  Logger.log(log.join('\n'));
  return { ok: true, log: log.join('\n') };
}

// ═══════════════════════════════════════════════════════════════════════════
// Daily Report Draft — 서버 백업
// localStorage가 비워진 상황(앱 재설치, PWA 캐시 정리, 다른 기기 접속)에도
// 작성 중인 Daily Report를 복원할 수 있도록 서버에 보조 저장한다.
//
// 시트: Daily_Draft
// 컬럼: [Driver, Updated_At, DraftJSON]
// — 드라이버당 1행 (덮어쓰기). 제출 / 명시적 clear 시 행 삭제.
// — 48시간 지나면 무효(서버에서도 제거).
// ═══════════════════════════════════════════════════════════════════════════
const DAILY_DRAFT_SHEET = 'Daily_Draft';
const DAILY_DRAFT_HEADERS = ['Driver', 'Updated_At', 'DraftJSON'];
const DAILY_DRAFT_TTL_MS = 48 * 60 * 60 * 1000; // 48h

function _getDailyDraftSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(DAILY_DRAFT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DAILY_DRAFT_SHEET);
    sh.getRange(1, 1, 1, DAILY_DRAFT_HEADERS.length)
      .setValues([DAILY_DRAFT_HEADERS])
      .setBackground('#1a56db').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(3, 600); // DraftJSON 넓게
  }
  return sh;
}

// 드라이버명 정확 일치 행을 찾아 row index(1-based) 반환. 없으면 -1.
function _findDailyDraftRow_(sh, driverName) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const driverCol = sh.getRange(2, 1, last - 1, 1).getValues();
  const target = String(driverName || '').trim();
  for (let i = 0; i < driverCol.length; i++) {
    if (String(driverCol[i][0] || '').trim() === target) return i + 2;
  }
  return -1;
}

function saveDailyDraftServer(driverName, draftJSON) {
  try {
    const name = String(driverName || '').trim();
    if (!name) return { ok: false, error: 'driver required' };
    if (typeof draftJSON !== 'string' || !draftJSON) {
      return { ok: false, error: 'draftJSON required' };
    }
    // GAS 셀 한도(50,000자) 안전 마진
    if (draftJSON.length > 45000) {
      return { ok: false, error: 'draft too large' };
    }

    const sh = _getDailyDraftSheet_();
    const now = new Date();
    const row = _findDailyDraftRow_(sh, name);
    if (row > 0) {
      sh.getRange(row, 1, 1, 3).setValues([[name, now, draftJSON]]);
    } else {
      sh.appendRow([name, now, draftJSON]);
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

function getDailyDraftServer(driverName) {
  try {
    const name = String(driverName || '').trim();
    if (!name) return { ok: false, error: 'driver required' };

    const sh = _getDailyDraftSheet_();
    const row = _findDailyDraftRow_(sh, name);
    if (row < 0) return { ok: true, draft: null };

    const vals = sh.getRange(row, 1, 1, 3).getValues()[0];
    const updatedAt = vals[1];
    const json = String(vals[2] || '');

    // TTL 검사
    const tsMs = (updatedAt instanceof Date) ? updatedAt.getTime()
               : (typeof updatedAt === 'number' ? updatedAt : Date.parse(updatedAt));
    if (tsMs && (Date.now() - tsMs) > DAILY_DRAFT_TTL_MS) {
      sh.deleteRow(row);
      return { ok: true, draft: null };
    }

    if (!json) return { ok: true, draft: null };

    return {
      ok: true,
      draft: json,
      updatedAt: tsMs || null
    };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

function clearDailyDraftServer(driverName) {
  try {
    const name = String(driverName || '').trim();
    if (!name) return { ok: false, error: 'driver required' };

    const sh = _getDailyDraftSheet_();
    const row = _findDailyDraftRow_(sh, name);
    if (row > 0) sh.deleteRow(row);
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}


/**
 * ─────────────────────────────────────────────────────────────────
 *  Bulk Sync All Vehicle Current_KM (매시간 트리거)
 * ─────────────────────────────────────────────────────────────────
 *  목적: Pre_Departure / Daily_Report / End_of_Shift 시트를 스캔해서
 *       각 차량(Rego)의 최신 KM을 찾아 M_Vehicles.Current_KM 컬럼에 반영.
 *
 *  트리거 등록: setupBulkSyncKMTrigger() 한 번만 실행
 *  트리거 제거: removeBulkSyncKMTrigger()
 *  수동 실행:   _bulkSyncAllVehicleCurrentKM()
 * ─────────────────────────────────────────────────────────────────
 */
function _bulkSyncAllVehicleCurrentKM() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const vSheet = ss.getSheetByName('M_Vehicles');
    if (!vSheet) {
      Logger.log('❌ M_Vehicles 시트 없음');
      return { ok: false, error: 'M_Vehicles not found' };
    }

    const lastRow = vSheet.getLastRow();
    const lastCol = vSheet.getLastColumn();
    if (lastRow < 2) {
      Logger.log('M_Vehicles 데이터 없음');
      return { ok: true, updated: 0, msg: 'no vehicles' };
    }

    const vHeaders = vSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const vRegoIdx = vHeaders.indexOf('Rego');
    const vKMIdx = vHeaders.indexOf('Current_KM');
    if (vRegoIdx < 0 || vKMIdx < 0) {
      Logger.log('❌ M_Vehicles에 Rego/Current_KM 컬럼 없음');
      return { ok: false, error: 'Rego or Current_KM column missing' };
    }

    // 1) 각 시트에서 Rego별 최대 KM 수집
    const kmMap = {};
    const scanForKM = (sheetName, kmFields) => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      const lr = sheet.getLastRow();
      const lc = sheet.getLastColumn();
      if (lr < 2 || lc < 1) return;
      const data = sheet.getRange(1, 1, lr, lc).getValues();
      const headers = data[0];
      const regoIdx = headers.indexOf('Rego');
      if (regoIdx < 0) return;
      const colIdxs = kmFields.map(f => headers.indexOf(f)).filter(i => i >= 0);
      if (colIdxs.length === 0) return;
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rego = String(row[regoIdx] || '').trim();
        if (!rego) continue;
        colIdxs.forEach(ci => {
          const v = parseFloat(row[ci]);
          if (!isNaN(v) && v > 0) {
            if (!kmMap[rego] || v > kmMap[rego]) kmMap[rego] = v;
          }
        });
      }
    };
    scanForKM('Pre_Departure', ['Start_KM']);
    scanForKM('Daily_Report',  ['KM_Start', 'KM_End']);
    scanForKM('End_of_Shift',  ['Start_KM', 'End_KM']);

    // 2) M_Vehicles 일괄 업데이트 (변동분만)
    const vData = vSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const kmColValues = vSheet.getRange(2, vKMIdx + 1, lastRow - 1, 1).getValues();
    let updated = 0;
    const updates = []; // {row, newKM}

    for (let i = 0; i < vData.length; i++) {
      const rego = String(vData[i][vRegoIdx] || '').trim();
      if (!rego) continue;
      const latest = kmMap[rego];
      if (latest == null) continue;
      const cur = parseFloat(kmColValues[i][0]);
      // 현재 값보다 새로 발견된 KM이 더 클 때만 업데이트
      if (isNaN(cur) || latest > cur) {
        updates.push({ row: i + 2, newKM: latest });
      }
    }

    // 3) 일괄 setValue (개별 호출 최소화)
    updates.forEach(u => {
      vSheet.getRange(u.row, vKMIdx + 1).setValue(u.newKM);
      updated++;
    });

    Logger.log('✅ Current_KM 동기화 완료: ' + updated + '대 업데이트 (전체 ' + vData.length + '대 중)');
    return { ok: true, updated: updated, total: vData.length };
  } catch (err) {
    Logger.log('❌ _bulkSyncAllVehicleCurrentKM 실패: ' + err);
    return { ok: false, error: err.toString() };
  }
}

function setupBulkSyncKMTrigger() {
  removeBulkSyncKMTrigger();
  ScriptApp.newTrigger('_bulkSyncAllVehicleCurrentKM')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('✅ Current_KM 자동 동기화 트리거 등록: 매시간');
  return 'BulkSyncKM trigger created.';
}

function removeBulkSyncKMTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  triggers.forEach(t => {
    if (t.getHandlerFunction() === '_bulkSyncAllVehicleCurrentKM') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  Logger.log('Removed ' + removed + ' bulkSyncKM trigger(s).');
  return removed;
}

// ═══════════════════════════════════════════════════════════════════════════
// EG TRAVEL 자동 리포트 발송 모듈
// ───────────────────────────────────────────────────────────────────────────
// - 매일 06:00: 전날 EG 관련 DR 정리 + 종료된 투어코드 별도 섹션 (중복 방지)
// - 매주 월요일 06:00: 지난주 EG 운행 요약 + 드라이버별 지급액
// - 수신자: EG TRAVEL 등록 이메일 + Branden (안전장치)
// - 발송 이력: EG_Report_Log 시트에 기록 (종료 투어 중복 발송 방지)
// ═══════════════════════════════════════════════════════════════════════════

const EG_REPORT_KEYWORD = 'EG TRAVEL';     // 매칭 키워드 (대소문자 무시)
const EG_REPORT_ADMIN_BCC = 'branden.dongchoi@gmail.com'; // 안전장치 — Branden에게 항상 BCC
const EG_REPORT_DAILY_HOUR = 6;            // 매일 발송 시각 (시드니 06:00)
const EG_REPORT_WEEKLY_HOUR = 6;           // 매주 월요일 06:00

// ── 헬퍼: 한 행에서 EG TRAVEL 관련 키워드 매칭 ─────────────────────────────
// Entity 정규화 — 풀네임/짧은코드/한글 모두 인식하여 'EG' / 'DC' / '' 로 변환
function _egNormEntity(s){
  const v = String(s||'').toUpperCase().trim();
  if(!v) return '';
  // EG TRAVEL, EG TRAVEL PTY LTD, EG 등 모두 매칭
  if(/\bEG\b/.test(v) || v.indexOf('EG TRAVEL') >= 0) return 'EG';
  // DONG CHOI PTY LTD, DC, 동초이 등
  if(v === 'DC' || v.indexOf('DONG CHOI') >= 0 || v.indexOf('DONGCHOI') >= 0) return 'DC';
  return v; // 기타 (제3자)
}

// Trailer 번호 → Owner 매핑 캐시
let _egTrailerOwnerCache = null;
function _egLoadTrailerOwners(){
  if(_egTrailerOwnerCache !== null) return _egTrailerOwnerCache;
  _egTrailerOwnerCache = {};
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('M_Trailers');
    if(!sheet) return _egTrailerOwnerCache;
    const data = sheet.getDataRange().getValues();
    if(data.length < 2) return _egTrailerOwnerCache;
    const headers = data[0].map(String);
    const tnIdx = headers.indexOf('Trailer_Number');
    const ownerIdx = headers.indexOf('Owner');
    if(tnIdx < 0 || ownerIdx < 0) return _egTrailerOwnerCache;
    for(let i=1; i<data.length; i++){
      const t = String(data[i][tnIdx]||'').trim();
      const o = String(data[i][ownerIdx]||'').trim();
      if(t) _egTrailerOwnerCache[t] = o;
    }
  } catch(e){
    Logger.log('_egLoadTrailerOwners error: ' + e);
  }
  return _egTrailerOwnerCache;
}

// M_PriceSub 로더 — SubCo별/Course별/좌석별 rate
let _egPriceSubCache = null;
function _egLoadPriceSub(){
  if(_egPriceSubCache !== null) return _egPriceSubCache;
  _egPriceSubCache = {};
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('M_PriceSub');
    if(!sheet) return _egPriceSubCache;
    const data = sheet.getDataRange().getValues();
    if(data.length < 2) return _egPriceSubCache;
    const headers = data[0].map(String);
    const subIdx = headers.indexOf('SubCo');
    const courseIdx = headers.indexOf('Course');
    const mhIdx = headers.indexOf('max_hours');
    const seatCols = {};
    ['21','25','40','50'].forEach(s => {
      seatCols[s] = {
        rate: headers.indexOf('seats_' + s + '_rate'),
        ot: headers.indexOf('seats_' + s + '_ot')
      };
    });
    for(let i=1; i<data.length; i++){
      const row = data[i];
      const sub = String(row[subIdx]||'').trim();
      const course = String(row[courseIdx]||'').trim();
      if(!sub || !course) continue;
      if(!_egPriceSubCache[sub]) _egPriceSubCache[sub] = {};
      const entry = { max_hours: Number(row[mhIdx])||0 };
      ['21','25','40','50'].forEach(s => {
        entry[s] = {
          rate: Number(row[seatCols[s].rate])||0,
          ot:   Number(row[seatCols[s].ot])||0
        };
      });
      _egPriceSubCache[sub][course] = entry;
    }
  } catch(e){
    Logger.log('_egLoadPriceSub error: ' + e);
  }
  return _egPriceSubCache;
}

// M_PriceDriver 로더 (드라이버 base rate)
let _egPriceDriverCache = null;
function _egLoadPriceDriver(){
  if(_egPriceDriverCache !== null) return _egPriceDriverCache;
  _egPriceDriverCache = {};
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('M_PriceDriver');
    if(!sheet) return _egPriceDriverCache;
    const data = sheet.getDataRange().getValues();
    if(data.length < 2) return _egPriceDriverCache;
    const headers = data[0].map(String);
    const courseIdx = headers.indexOf('Course');
    const mhIdx = headers.indexOf('max_hours');
    const seatCols = {};
    ['21','25','40','50'].forEach(s => {
      seatCols[s] = {
        base: headers.indexOf('seats_' + s + '_base'),
        ot: headers.indexOf('seats_' + s + '_ot')
      };
    });
    for(let i=1; i<data.length; i++){
      const row = data[i];
      const course = String(row[courseIdx]||'').trim();
      if(!course) continue;
      const entry = { max_hours: Number(row[mhIdx])||0 };
      ['21','25','40','50'].forEach(s => {
        entry[s] = {
          base: Number(row[seatCols[s].base])||0,
          ot: Number(row[seatCols[s].ot])||0
        };
      });
      _egPriceDriverCache[course] = entry;
    }
  } catch(e){
    Logger.log('_egLoadPriceDriver error: ' + e);
  }
  return _egPriceDriverCache;
}

// 드라이버 지급액 breakdown — {total, items}
function _egCalcDriverPay(r){
  const PD = _egLoadPriceDriver();
  const attraction = String(r.Attraction||r.tour||'').trim();
  const seatsRaw = String(r.Seats||r.seats||'').replace(/S/i,'').trim();
  const capNum = parseInt(seatsRaw)||25;
  const capKey = capNum>=50?'50':capNum>=40?'40':capNum>=25?'25':'21';

  // Base rate from M_PriceDriver
  function _findCourse(cn){
    if(!PD || !cn) return null;
    if(PD[cn]) return PD[cn];
    const lc = cn.toLowerCase();
    const keys = Object.keys(PD);
    for(let i=0; i<keys.length; i++){
      if(keys[i].toLowerCase() === lc) return PD[keys[i]];
    }
    return null;
  }
  const ce = _findCourse(attraction);
  let baseRate = 0;
  if(ce){ const sd = ce[capKey] || ce['21']; baseRate = Number(sd && sd.base) || 0; }

  // 각 항목 (DR 값 그대로)
  const ot   = Number(r.OT||0);
  const htl  = Number(r.Hotel_Surcharge||0);
  const dst  = Number(r.Dist_Surcharge||0);
  const erl  = Number(r.Early||0);
  const trl  = Number(r.Trailer||0);
  const ngt  = Number(r.Night_DR||0);
  const ngo  = Number(r.Night_Owner||0);
  const wash = Number(r.Wash||0);
  const meal = Number(r.Meal||0);
  const tip  = Number(r.Tip||0);
  const etc  = Number(r.Etc||0);
  const etcDesc = String(r.Etc_Desc||'').trim();
  const tollP = String(r.Toll_Personal||'').toUpperCase() === 'Y' ? Number(r.Toll||0) : 0;
  const fuelP = String(r.Fuel_Personal||'').toUpperCase() === 'Y' ? Number(r.Fuel||0) : 0;

  const items = [];
  if(baseRate !== 0){
    items.push({label: 'Base (' + (attraction||'코스') + ' · ' + capNum + 'S)', amount: baseRate});
  } else {
    // fallback: 시트의 DR_Cost를 base로 표시할 수도 있지만 일단 0 처리
    const drStored = Number(r.DR_Cost || r.Total || 0);
    if(drStored !== 0){
      items.push({label: 'Base (저장값 사용)', amount: drStored, note: 'M_PriceDriver 매칭 없음'});
      return {
        total: drStored,
        items: items,
        valueOf: function(){ return this.total; },
        toString: function(){ return String(this.total); }
      };
    }
  }
  if(ot !== 0)  items.push({label: 'OT', amount: ot});
  if(htl !== 0) items.push({label: '호텔 서차지', amount: htl});
  if(dst !== 0) items.push({label: '거리 서차지', amount: dst});
  if(erl !== 0) items.push({label: '조기 서차지', amount: erl});
  if(trl !== 0) items.push({label: '트레일러', amount: trl});
  if(ngt !== 0) items.push({label: '야간 운행', amount: ngt});
  if(wash !== 0) items.push({label: '세차비', amount: wash});
  if(meal !== 0) items.push({label: '식비', amount: meal});
  if(tip !== 0) items.push({label: '팁', amount: tip});
  if(tollP !== 0) items.push({label: '톨비 (개인)', amount: tollP});
  if(fuelP !== 0) items.push({label: '연료 (개인)', amount: fuelP});
  if(etc !== 0) items.push({label: '기타' + (etcDesc?' ('+etcDesc+')':''), amount: etc});
  if(ngo !== 0) items.push({label: '차주 납입 차감', amount: -Math.abs(ngo)});

  const total = items.reduce((s, it) => s + it.amount, 0);

  return {
    total: total,
    items: items,
    valueOf: function(){ return this.total; },
    toString: function(){ return String(this.total); }
  };
}

// ── 환산 헬퍼 (admin.html과 동일한 로직) ──
function _egHotelDRtoTA(dr, sn){
  if(!dr || dr===0) return 0;
  if(sn>=50) return dr===15?80:dr===30?160:dr*4;
  if(sn>=40) return dr===15?75:dr===30?150:dr*4;
  return dr===10?40:dr===20?80:dr*4;
}
function _egDistDRtoTA(dr, sn){
  if(!dr || dr===0) return 0;
  if(sn>=50) return dr===40?160:dr*4;
  if(sn>=40) return dr===40?150:Math.round(dr*3.75);
  return dr===30?80:Math.round(dr*2.67);
}
function _egTrailerSurchargeDRtoTA(dr, sn){
  // 트레일러 서차지 환산: 21/25S DR$30→$80, 40S+ 청구 없음
  if(!dr || dr===0) return 0;
  if(sn>=40) return 0;
  return dr===30?80:Math.round(dr*2.67);
}

// EG SUB 청구액 계산 — calcSubReport와 정합하는 로직
// 반환: {total, items: [{label, amount, note}]}
function _egCalcEgSubAmount(r){
  const PS = _egLoadPriceSub();
  const attraction = String(r.Attraction||r.tour||'').trim();
  const seatsRaw = String(r.Seats||r.seats||'').replace(/S/i,'').trim();
  const capNum = parseInt(seatsRaw)||25;
  const capKey = capNum>=50?'50':capNum>=40?'40':capNum>=25?'25':'21';
  const isLarge = capNum>=40;
  const agency = String(r.Agency||r.agency||'').trim();

  // 1) M_PriceSub의 EG TRAVEL 행에서 base rate
  function _findSub(){
    const keys = Object.keys(PS);
    for(let i=0; i<keys.length; i++){
      if(_egNormEntity(keys[i]) === 'EG') return PS[keys[i]];
    }
    return null;
  }
  function _findCourse(pc, cn){
    if(!pc || !cn) return null;
    if(pc[cn]) return pc[cn];
    const lc = cn.toLowerCase();
    const keys = Object.keys(pc);
    for(let i=0; i<keys.length; i++){
      if(keys[i].toLowerCase() === lc) return pc[keys[i]];
    }
    return null;
  }
  const subEntity = _findSub();
  const ce = subEntity ? _findCourse(subEntity, attraction) : null;
  let baseRate = 0;
  let baseSource = '';
  if(ce){
    const sd = ce[capKey] || ce['21'];
    baseRate = Number(sd && sd.rate) || 0;
    baseSource = 'M_PriceSub';
  }
  if(baseRate === 0){
    baseRate = Number(r.SVC_Charge)||0;
    baseSource = 'SVC_Charge';
  }

  // 2) 서차지 — TA 환산식 적용 (calcSubReport와 동일)
  const hotelDR = Number(r.Hotel_Surcharge||0);
  const distDR  = Number(r.Dist_Surcharge||0);
  const trailerDR = Number(r.Trailer||0);
  const otDR  = Number(r.OT||0);
  const earlyDR = Number(r.Early||0);
  const toll = Number(r.Toll||0);

  const hotelTA   = _egHotelDRtoTA(hotelDR, capNum);
  const distTA    = _egDistDRtoTA(distDR, capNum);
  const trailerTA = _egTrailerSurchargeDRtoTA(trailerDR, capNum);

  // OT 환산 — 호주로(Tour Hojuro)/Plus Australia 21~25S: 30분 UNIT
  const otRateTA = capNum>=50?160:capNum>=40?150:80;
  const otRateDR = capNum>=40?40:30;
  const isHojuroOT = /호주로|hojuro|plus\s*australia/i.test(agency);
  const otTA = isHojuroOT
    ? Math.round((otDR / (otRateDR/2)) * (otRateTA/2))
    : Math.round((otRateDR>0 ? otDR/otRateDR : 0) * otRateTA);

  // Early 환산 — Hojuro 21/25S: $80 / 그 외: Airport Transfer rate × 0.3
  let earlyTA = 0;
  if(earlyDR > 0){
    const isHojuroEarly = /호주로|hojuro|plus\s*australia/i.test(agency);
    if(isHojuroEarly && capNum < 40){
      earlyTA = 80;
    } else {
      // M_PriceClient에서 같은 여행사의 Airport Transfer rate 찾기 (fallback: 다른 여행사)
      const PC = _egLoadPriceClient();
      let atE = null;
      const agPC = PC[agency];
      if(agPC){
        atE = _findCourse(agPC, 'Airport Transfer');
      }
      if(!atE){
        const allAgs = Object.keys(PC);
        for(let i=0; i<allAgs.length; i++){
          atE = _findCourse(PC[allAgs[i]], 'Airport Transfer');
          if(atE) break;
        }
      }
      if(atE){
        const sd2 = atE[capKey] || atE['21'];
        earlyTA = Math.round((Number(sd2 && sd2.rate)||0) * 0.3);
      }
    }
  }

  // 3) 공항 픽업 주차비 (EG는 항상 부담)
  // ★ 단, Tour Hojuro / Plus Australia + 21/25S: 여행사가 청구 안 받음 → EG도 부담 안 함
  const apPat = /\b(airport|syd|kingsford|mascot|international|domestic|terminal)\b/i;
  const pickup = String(r.Pickup||'');
  const isHojuroParking = /호주로|hojuro|plus\s*australia/i.test(agency);
  const _excludeParkingForAgency = isHojuroParking && capNum < 40;  // 21/25S만
  const parking = (apPat.test(pickup) && !_excludeParkingForAgency) ? (isLarge ? 40 : 30) : 0;

  // 4) Toll (대형만)
  const tollAmt = isLarge ? toll : 0;

  // 5) 트레일러 대여비 - 트레일러 소유주가 식별되면 -$30 (소유주에게 지급)
  let trailerRental = 0;
  let trailerOwnerName = '';
  if(trailerDR > 0){
    const trNum = String(r.Trailer_Number||'').trim();
    if(trNum){
      const tOwners = _egLoadTrailerOwners();
      const rawOwner = tOwners[trNum] || '';
      const trOwner = _egNormEntity(rawOwner);
      if(trOwner){
        trailerRental = -30;
        trailerOwnerName = rawOwner;
      }
    }
  }

  // Breakdown 구성
  const items = [];
  if(baseRate !== 0){
    items.push({
      label: 'Base (' + (attraction||'코스') + ' · ' + capNum + 'S)',
      amount: baseRate,
      note: baseSource === 'SVC_Charge' ? '(SVC fallback)' : ''
    });
  }
  if(otTA !== 0){
    items.push({
      label: 'OT' + (otDR ? ' (DR $' + otDR + ' → TA $' + otTA + ')' : ''),
      amount: otTA
    });
  }
  if(hotelTA !== 0){
    items.push({
      label: '호텔 서차지 (DR $' + hotelDR + ' → TA $' + hotelTA + ')',
      amount: hotelTA
    });
  }
  if(distTA !== 0){
    items.push({
      label: '거리 서차지 (DR $' + distDR + ' → TA $' + distTA + ')',
      amount: distTA
    });
  }
  if(earlyTA !== 0){
    items.push({
      label: '조기 서차지 (DR $' + earlyDR + ' → TA $' + earlyTA + ')',
      amount: earlyTA
    });
  }
  if(parking !== 0) items.push({label: '공항 픽업 주차비', amount: parking});
  if(tollAmt !== 0) items.push({label: '톨비', amount: tollAmt});
  if(trailerTA !== 0){
    items.push({
      label: '트레일러 서차지 (DR $' + trailerDR + ' → TA $' + trailerTA + ')',
      amount: trailerTA
    });
  }
  if(trailerRental !== 0){
    items.push({
      label: '트레일러 대여비',
      amount: trailerRental,
      note: trailerOwnerName ? '(소유주: ' + trailerOwnerName + ')' : ''
    });
  }

  const total = baseRate + otTA + hotelTA + distTA + earlyTA + parking + tollAmt + trailerTA + trailerRental;

  return {
    total: total,
    items: items,
    valueOf: function(){ return this.total; },
    toString: function(){ return String(this.total); }
  };
}

// 운행 분류 — 'EG_BILLS' (Billing=EG, EG가 여행사에 청구) / 'DC_BILLS_EG_VEH' (Billing=DC, EG차량 sub) / null
function _egClassifyRow(r){
  const billing = _egNormEntity(r.Billing_Entity || r.BillingEntity || 'DC');
  if(billing === 'EG'){
    // EG가 청구하는 일정 — 차량이 DC면 DC가 EG에게 sub로 청구
    const owners = _egLoadVehicleOwners();
    const vehOwner = _egNormEntity(owners[String(r.Rego||'').trim()] || '');
    if(vehOwner === 'DC') return 'EG_BILLS_DC_VEH';  // DC → EG 청구
    return 'EG_BILLS_OWN';  // EG 자체 운행 (참고용, 보통 리포트 제외)
  }
  if(billing === 'DC'){
    const owners = _egLoadVehicleOwners();
    const vehOwner = _egNormEntity(owners[String(r.Rego||'').trim()] || '');
    if(vehOwner === 'EG') return 'DC_BILLS_EG_VEH';  // EG → DC 청구
  }
  return null;
}

// 차량(Rego) → Owner 매핑 캐시 (요청당 1회 로드)
let _egVehicleOwnerCache = null;
function _egLoadVehicleOwners(){
  if(_egVehicleOwnerCache !== null) return _egVehicleOwnerCache;
  _egVehicleOwnerCache = {};
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('M_Vehicles');
    if(!sheet) return _egVehicleOwnerCache;
    const data = sheet.getDataRange().getValues();
    if(data.length < 2) return _egVehicleOwnerCache;
    const headers = data[0].map(String);
    const regoIdx = headers.indexOf('Rego');
    const ownerIdx = headers.indexOf('Owner');
    if(regoIdx < 0 || ownerIdx < 0) return _egVehicleOwnerCache;
    for(let i=1; i<data.length; i++){
      const r = String(data[i][regoIdx]||'').trim();
      const o = String(data[i][ownerIdx]||'').trim();
      if(r) _egVehicleOwnerCache[r] = o;
    }
  } catch(e){
    Logger.log('_egLoadVehicleOwners error: ' + e);
  }
  return _egVehicleOwnerCache;
}

function _egRowMatches(row){
  if(!row || typeof row !== 'object') return false;
  // 새 분류 로직: 3가지 케이스 모두 리포트 포함
  // EG_BILLS_DC_VEH: EG 빌링, DC 차량  → EG가 DC에 지급
  // DC_BILLS_EG_VEH: DC 빌링, EG 차량  → DC가 EG에 지급 (EG가 받음)
  // EG_BILLS_OWN:    EG 빌링, EG 차량  → EG가 여행사 직접 청구
  const cls = _egClassifyRow(row);
  return cls === 'EG_BILLS_DC_VEH' || cls === 'DC_BILLS_EG_VEH' || cls === 'EG_BILLS_OWN';
}

// ── 날짜 헬퍼 ─────────────────────────────────────────────────────────────
function _egToISO(dStr){
  if(!dStr) return '';
  if(dStr instanceof Date){
    return Utilities.formatDate(dStr, 'Australia/Sydney', 'yyyy-MM-dd');
  }
  const s = String(dStr).trim();
  // ISO already
  let m = s.match(/^(\d{4}-\d{2}-\d{2})/);
  if(m) return m[1];
  // dd/MM/yyyy
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if(m) return m[3]+'-'+m[2].padStart(2,'0')+'-'+m[1].padStart(2,'0');
  // try parsing
  try {
    const d = new Date(s);
    if(!isNaN(d.getTime())) return Utilities.formatDate(d, 'Australia/Sydney', 'yyyy-MM-dd');
  } catch(e){}
  return '';
}
function _egFmtDate(iso){
  if(!iso) return '—';
  const m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(m) return m[3]+'/'+m[2]+'/'+m[1];
  return String(iso);
}
function _egTodaySydney(){
  return Utilities.formatDate(new Date(), 'Australia/Sydney', 'yyyy-MM-dd');
}
function _egYesterdaySydney(){
  const d = new Date();
  d.setDate(d.getDate() - 1);
  return Utilities.formatDate(d, 'Australia/Sydney', 'yyyy-MM-dd');
}
function _egMondayOf(iso){
  // ISO 날짜의 같은 주 월요일 ISO 반환 (호주식: 월요일 시작)
  const d = new Date(iso + 'T00:00:00');
  if(isNaN(d.getTime())) return iso;
  const day = d.getDay(); // 0=일, 1=월
  const diff = (day === 0) ? -6 : 1 - day;
  d.setDate(d.getDate() + diff);
  return Utilities.formatDate(d, 'Australia/Sydney', 'yyyy-MM-dd');
}

// ── 수신자 결정 ───────────────────────────────────────────────────────────
function _egGetRecipients(){
  // EG TRAVEL의 M_Clients 등록 이메일 + Branden 본인 (BCC)
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const cliSheet = ss.getSheetByName('M_Clients');
  let toList = [];
  let ccList = [];
  if(cliSheet){
    const data = cliSheet.getDataRange().getValues();
    const headers = data[0].map(String);
    const nameIdx = headers.indexOf('Name');
    const emailIdx = headers.indexOf('Email');
    const ccIdx = headers.indexOf('Email_CC');
    for(let i=1; i<data.length; i++){
      const name = String(data[i][nameIdx]||'').trim();
      if(name.toUpperCase().indexOf(EG_REPORT_KEYWORD) >= 0){
        const em = String(data[i][emailIdx]||'').trim();
        const cc = String(data[i][ccIdx]||'').trim();
        if(em) toList.push(em);
        if(cc) ccList.push(cc);
        break;
      }
    }
  }
  return {
    to: toList.join(', '),
    cc: ccList.join(', '),
    bcc: EG_REPORT_ADMIN_BCC  // Branden 안전장치
  };
}

// ── 이미 발송된 종료 투어코드 조회 (중복 방지) ────────────────────────────
function _egGetAlreadySentTourCodes(){
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ensureSheet(ss, 'EG_Report_Log');
  const data = sheet.getDataRange().getValues();
  if(data.length < 2) return new Set();
  const headers = data[0].map(String);
  const tcIdx = headers.indexOf('TourCodes');
  const statusIdx = headers.indexOf('Status');
  const sent = new Set();
  for(let i=1; i<data.length; i++){
    const status = statusIdx >= 0 ? String(data[i][statusIdx]||'').trim().toUpperCase() : 'OK';
    if(status === 'FAILED' || status === 'ERROR') continue;
    const tcs = String(data[i][tcIdx]||'').trim();
    if(!tcs) continue;
    tcs.split(',').forEach(t => {
      const tt = t.trim().toUpperCase();
      if(tt) sent.add(tt);
    });
  }
  return sent;
}

// ── 발송 이력 기록 ─────────────────────────────────────────────────────────
function _egLogReportSent(reportType, periodFrom, periodTo, recipients, tourCodes, subject, status, notes){
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSheet(ss, 'EG_Report_Log');
    sheet.appendRow([
      Utilities.formatDate(new Date(), 'Australia/Sydney', 'yyyy-MM-dd HH:mm:ss'),
      reportType, periodFrom||'', periodTo||'',
      recipients||'', (tourCodes||[]).join(','),
      subject||'', status||'OK', notes||''
    ]);
  } catch(e){
    Logger.log('EG report log error: ' + e);
  }
}

// ── 데이터 로더 ────────────────────────────────────────────────────────────
function _egLoadDRs(fromISO, toISO){
  // Daily_Report에서 fromISO~toISO 기간의 EG 관련 행만 추출
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Daily_Report');
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if(data.length < 2) return [];
  const headers = data[0].map(String);
  const rows = [];
  for(let i=1; i<data.length; i++){
    const row = {};
    headers.forEach((h, ci) => { row[h] = data[i][ci]; });
    if(!_egRowMatches(row)) continue;
    const iso = _egToISO(row.Date);
    if(!iso) continue;
    if(iso < fromISO || iso > toISO) continue;
    row._iso = iso;
    rows.push(row);
  }
  return rows;
}

function _egLoadSchedule(){
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Schedule');
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if(data.length < 2) return [];
  const headers = data[0].map(String);
  const rows = [];
  for(let i=1; i<data.length; i++){
    const row = {};
    headers.forEach((h, ci) => { row[h] = data[i][ci]; });
    if(!_egRowMatches(row)) continue;
    rows.push(row);
  }
  return rows;
}

// ── 종료된 투어코드 판정 ──────────────────────────────────────────────────
// 다음 조건 중 하나라도 충족하면 종료:
//   1) Schedule.EndDate < 오늘
//   2) Schedule.Status === 'completed' || 'invoiced' || 'paid'
//   3) 해당 TourCode의 모든 DR이 제출됐고 마지막 DR.Date < 오늘
function _egFindCompletedTourCodes(todayISO){
  const sched = _egLoadSchedule();
  const drs = _egLoadDRs('2020-01-01', todayISO); // 모든 과거
  const completed = [];
  const _drByTC = {};
  drs.forEach(r => {
    const tc = String(r.Tour_Code || r.TourCode || '').trim().toUpperCase();
    if(!tc) return;
    if(!_drByTC[tc]) _drByTC[tc] = [];
    _drByTC[tc].push(r);
  });
  sched.forEach(s => {
    const tc = String(s.TourCode || '').trim().toUpperCase();
    if(!tc) return;
    const endISO = _egToISO(s.EndDate);
    const status = String(s.Status || '').trim().toLowerCase();
    let isDone = false;
    let reason = '';
    // 조건 1
    if(endISO && endISO < todayISO){ isDone = true; reason = '일정 종료일 경과'; }
    // 조건 2
    if(!isDone && (status === 'completed' || status === 'invoiced' || status === 'paid')){
      isDone = true; reason = '상태: ' + status;
    }
    // 조건 3
    if(!isDone){
      const tcDRs = _drByTC[tc] || [];
      if(tcDRs.length > 0){
        const lastDR = tcDRs.map(r=>r._iso).sort().reverse()[0];
        if(lastDR && lastDR < todayISO){
          isDone = true; reason = 'DR 마지막일 경과 (' + lastDR + ')';
        }
      }
    }
    if(isDone){
      completed.push({
        tourCode: tc,
        agency: s.Agency || '',
        startDate: _egToISO(s.StartDate),
        endDate: endISO,
        status: s.Status,
        guide: s.Guide || '',
        pax: s.Pax || '',
        reason: reason,
        drs: _drByTC[tc] || []
      });
    }
  });
  return completed;
}

// ── HTML 빌더 공통 스타일 ─────────────────────────────────────────────────
function _egCommonStyle(){
  return `
    <style>
      body{font-family:Arial,'Malgun Gothic','맑은 고딕',sans-serif;color:#1f2937;margin:0;padding:18px;font-size:11pt;line-height:1.4;}
      .hdr{border-bottom:3px solid #7c3aed;padding-bottom:12px;margin-bottom:20px;}
      .hdr h1{margin:0 0 4px;font-size:18pt;color:#7c3aed;}
      .hdr .sub{color:#6b7280;font-size:10pt;}
      .sec-title{font-size:13pt;font-weight:bold;color:#1f2937;background:#f3f4f6;
                 padding:8px 12px;border-left:4px solid #7c3aed;margin:18px 0 10px;}
      .meta{display:table;width:100%;background:#faf5ff;border:1px solid #e9d5ff;
            border-radius:6px;padding:10px 14px;margin-bottom:12px;font-size:10pt;}
      .meta div{display:table-cell;width:25%;text-align:center;padding:0 6px;}
      .meta div + div{border-left:1px solid #e9d5ff;}
      .tc-badge{display:inline-block;background:#ede9fe;color:#5b21b6;padding:1px 7px;
                border-radius:4px;font-weight:bold;font-size:9.5pt;}
      .date-badge{display:inline-block;background:#1e40af;color:#fff;font-size:9pt;
                  font-weight:800;padding:2px 8px;border-radius:4px;letter-spacing:.3px;}
      .sub-badge{display:inline-block;background:#7c3aed;color:white;font-size:8.5pt;
                 font-weight:700;padding:2px 7px;border-radius:8px;margin-left:5px;}
      .empty{text-align:center;padding:24px;color:#9ca3af;font-style:italic;font-size:10pt;}

      /* 드라이버 그룹 헤더 */
      .driver-grp{margin:12px 0;border-radius:10px;overflow:hidden;
                  border:1px solid #d1d5db;background:white;}
      .driver-grp .hdr-bar{background:#1f2937;color:white;padding:8px 14px;
                           font-weight:bold;font-size:11pt;display:table;width:100%;}
      .driver-grp .hdr-bar > div{display:table-cell;}
      .driver-grp .hdr-bar .right{text-align:right;color:#fbbf24;}

      /* 운행 카드 (드라이버 급여 스타일) */
      .trip-card{background:#f8fafc;border-radius:8px;padding:10px 12px;margin:6px 8px;
                 border:1px solid #e5e7eb;}
      .trip-card.sub{background:#f5f3ff;border-left:3px solid #7c3aed;}
      .trip-card .top{display:table;width:100%;}
      .trip-card .top .info{display:table-cell;vertical-align:top;}
      .trip-card .top .amt{display:table-cell;vertical-align:top;text-align:right;
                           min-width:110px;padding-left:8px;}
      .trip-card .title{font-size:11pt;font-weight:700;color:#1f2937;margin-top:4px;}
      .trip-card .meta-line{font-size:9pt;color:#6b7280;margin-top:2px;}
      .trip-card .time-line{font-size:9pt;color:#4f46e5;margin-top:1px;}
      .trip-card .amount{font-size:13pt;font-weight:800;color:#16a34a;}
      .trip-card .amount.neg{color:#dc2626;}
      .trip-card .km-badge{background:#dcfce7;color:#16a34a;padding:0 5px;
                           border-radius:3px;font-weight:700;font-size:9pt;}
      .trip-card .night-own{font-size:9pt;color:#dc2626;font-weight:700;margin-top:1px;}
      .trip-card .night-own.sub{color:#7c3aed;}

      /* 서차지 배지 */
      .surcharge-row{margin-top:6px;font-size:8.5pt;}
      .sur-badge{display:inline-block;padding:1px 6px;border-radius:4px;font-weight:700;
                 margin-right:3px;margin-top:2px;border:1px solid;}

      /* 종료 투어 카드 */
      .tour-card{background:white;border:1.5px solid #d1d5db;border-radius:10px;
                 padding:0;margin:12px 0;overflow:hidden;}
      .tour-card .head{background:linear-gradient(135deg,#7c3aed,#5b21b6);color:white;
                       padding:10px 14px;}
      .tour-card .head .row1{display:table;width:100%;}
      .tour-card .head .row1 .left{display:table-cell;}
      .tour-card .head .row1 .right{display:table-cell;text-align:right;}
      .tour-card .head .tc-name{font-size:13pt;font-weight:bold;}
      .tour-card .head .agency{font-size:9.5pt;opacity:.9;margin-top:2px;}
      .tour-card .head .period{font-size:10pt;font-weight:700;}
      .tour-card .head .meta-row{display:table;width:100%;margin-top:8px;
                                  border-top:1px solid rgba(255,255,255,.3);padding-top:6px;}
      .tour-card .head .meta-row > div{display:table-cell;font-size:9pt;}
      .tour-card .body{padding:10px 14px;}
      .tour-card .reason{font-size:9pt;color:#7c3aed;background:#ede9fe;
                         padding:4px 10px;border-radius:4px;margin-bottom:8px;font-weight:600;}
      .tour-card .dr-list{margin-top:8px;}
      .tour-card .dr-row{display:table;width:100%;font-size:9.5pt;padding:5px 0;
                         border-bottom:1px solid #f3f4f6;}
      .tour-card .dr-row > div{display:table-cell;padding:0 4px;}
      .tour-card .dr-row .dt{width:80px;color:#6b7280;}
      .tour-card .dr-row .info{color:#1f2937;}
      .tour-card .dr-row .amt{text-align:right;width:95px;font-weight:700;color:#7c3aed;
                              font-family:Consolas,monospace;}
      .tour-card .dr-row .amt-dr{text-align:right;width:80px;font-weight:700;color:#16a34a;
                                  font-family:Consolas,monospace;}
      .tour-card .totals{margin-top:8px;padding-top:8px;border-top:2px solid #e5e7eb;
                         background:#f9fafb;padding:10px 14px;margin:8px -14px -10px;
                         display:table;width:calc(100% + 28px);}
      .tour-card .totals > div{display:table-cell;vertical-align:middle;}
      .tour-card .totals .label{font-size:10.5pt;font-weight:700;color:#5b21b6;line-height:1.4;}
      .tour-card .totals .val{text-align:right;font-size:9pt;color:#6b7280;width:140px;line-height:1.3;}
      .tour-card .totals .val-ta{text-align:right;font-size:9pt;color:#6b7280;width:140px;line-height:1.3;border-right:1px solid #e5e7eb;padding-right:12px;}

      .ftr{margin-top:30px;padding-top:12px;border-top:1px solid #e5e7eb;
           color:#6b7280;font-size:8.5pt;text-align:center;}
      table{border-collapse:collapse;width:100%;font-size:9.5pt;margin-bottom:12px;}
      th{background:#7c3aed;color:white;padding:7px 6px;text-align:left;font-weight:600;}
      td{padding:6px;border-bottom:1px solid #e5e7eb;}
      tr:nth-child(even) td{background:#fafafa;}
      .num{text-align:right;font-family:Consolas,monospace;}
      .summary-box{background:#f0fdf4;border:1px solid #86efac;border-radius:6px;
                   padding:10px 14px;margin:8px 0;}
      .summary-row{display:table;width:100%;padding:3px 0;}
      .summary-row > div{display:table-cell;}
      .summary-row > div + div{text-align:right;font-weight:bold;}
      .summary-row.tot{border-top:2px solid #16a34a;margin-top:6px;padding-top:8px;}
    </style>
  `;
}

// ── 운행 카드 빌더 (관리자 급여 탭 _reportRow 스타일 — 정적 PDF에 맞게 펼친 형태) ──
function _egTripCardHTML(r){
  const drCostStored = Number(r.DR_Cost || r.Total || 0) || 0;
  const drCalc = _egCalcDriverPay(r);   // {total, items}
  const drCost = drCalc.total !== 0 ? drCalc.total : drCostStored;
  const drBreakdown = drCalc.items;
  const taCalc = _egCalcEgSubAmount(r);  // {total, items}
  const taAmount = taCalc.total;
  const breakdown = taCalc.items;
  const cls = _egClassifyRow(r);  // EG_BILLS_DC_VEH / DC_BILLS_EG_VEH / EG_BILLS_OWN
  const nightOwn = Number(r.Night_Owner || 0) || 0;
  const tS = _egFmtTime(r.Time_Start || r.Start_Time);
  const tE = _egFmtTime(r.Time_End || r.End_Time);
  const timeStr = tS ? (tS + (tE ? ' ~ ' + tE : '')) : '';
  const kmS = String(r.KM_Start || r.Start_KM || '').trim();
  const kmE = String(r.KM_End || r.End_KM || '').trim();
  let kmDiff = 0;
  if(kmS && kmE){
    const ks = parseInt(kmS), ke = parseInt(kmE);
    if(!isNaN(ks) && !isNaN(ke)) kmDiff = Math.abs(ke - ks);
  }
  const hotel = r.Hotel || r.Accommodation || '';
  const guide = r.Guide || '';
  const tc = r.Tour_Code || r.TourCode || '';
  const rego = r.Rego || '';
  const seats = r.Seats || r.Pax || '';
  const agency = r.Agency || r.Tour_Agency || '';
  const attraction = r.Attraction || r.Course || '';

  // 청구/지급 대상 뱃지 (한 운행에 여러 방향 표시 가능)
  const badges = [];  // [{text, color}]
  if(cls === 'DC_BILLS_EG_VEH'){
    // DC가 빌링, EG 차량 → EG가 DC에 청구 (받을 돈)
    badges.push({text: '→ DC 청구', color: '#7c3aed'});
  } else if(cls === 'EG_BILLS_DC_VEH'){
    // EG가 빌링, DC 차량 → 여행사에 청구 + DC에 지급 (양방향)
    badges.push({text: '→ ' + (agency || '여행사') + ' 청구', color: '#0891b2'});
    badges.push({text: '← DC 지급', color: '#dc2626'});
  } else if(cls === 'EG_BILLS_OWN'){
    // EG 자체 운행 → 여행사 직접 청구
    badges.push({text: '→ ' + (agency || '여행사') + ' 청구', color: '#0891b2'});
  }

  // 서차지 배지
  const sur = [];
  if(Number(r.OT) > 0) sur.push({l:'OT', v:Number(r.OT), c:'#3b82f6'});
  if(Number(r.Hotel_Surcharge) > 0) sur.push({l:'호텔', v:Number(r.Hotel_Surcharge), c:'#8b5cf6'});
  if(Number(r.Dist_Surcharge) > 0) sur.push({l:'거리', v:Number(r.Dist_Surcharge), c:'#0ea5e9'});
  if(Number(r.Early) > 0) sur.push({l:'조기', v:Number(r.Early), c:'#f59e0b'});
  if(Number(r.Trailer) > 0) sur.push({l:'트레일러', v:Number(r.Trailer), c:'#64748b'});
  if(Number(r.Night_DR) > 0) sur.push({l:'야간', v:Number(r.Night_DR), c:'#a5b4fc'});
  if(Number(r.Wash) > 0) sur.push({l:'세차', v:Number(r.Wash), c:'#10b981'});
  if(Number(r.Meal) > 0) sur.push({l:'식비', v:Number(r.Meal), c:'#10b981'});
  if(Number(r.Tip) > 0) sur.push({l:'팁', v:Number(r.Tip), c:'#eab308'});
  if(String(r.Toll_Personal||'').toUpperCase() === 'Y' && Number(r.Toll) > 0)
    sur.push({l:'톨비', v:Number(r.Toll), c:'#78716c'});
  if(String(r.Fuel_Personal||'').toUpperCase() === 'Y' && Number(r.Fuel) > 0)
    sur.push({l:'연료', v:Number(r.Fuel), c:'#78716c'});

  const _fmtAmt = (v) => (v < 0 ? '-$' : '$') + Math.abs(v).toLocaleString('en-AU',{minimumFractionDigits:2, maximumFractionDigits:2});
  const dateStr = _egFmtDate(r._iso);
  const isPaid = cls === 'EG_BILLS_DC_VEH';

  let html = '<div class="trip-card' + (isPaid ? ' sub' : '') + '">';
  html += '<div class="top">';
  html += '<div class="info">';
  html += '<div><span class="date-badge">📅 ' + _egEsc(dateStr) + '</span>';
  badges.forEach(b => {
    html += '<span class="sub-badge" style="background:' + b.color + ';margin-left:6px;">' + _egEsc(b.text) + '</span>';
  });
  html += '</div>';
  html += '<div class="title">' + _egEsc(agency) + ' · ' + _egEsc(attraction) + '</div>';

  // 거래 흐름 설명
  let flowNote = '';
  if(cls === 'EG_BILLS_DC_VEH'){
    // EG가 여행사 청구 + DC 차량 sub → 받은 돈을 DC에 패스스루
    flowNote = '💡 EG가 ' + _egEsc(agency || '여행사') + '에 청구 → 받은 금액을 DC에 지급';
  } else if(cls === 'DC_BILLS_EG_VEH'){
    // DC가 여행사 청구 + EG 차량 sub → EG가 DC에 청구
    flowNote = '💡 DC가 ' + _egEsc(agency || '여행사') + '에 청구 → EG는 sub로 운행, DC에 청구';
  } else if(cls === 'EG_BILLS_OWN'){
    flowNote = '💡 EG 자체 운행 — ' + _egEsc(agency || '여행사') + '에 직접 청구';
  }
  if(flowNote){
    html += '<div style="font-size:8.5pt;color:#6b7280;font-style:italic;margin:2px 0 4px;">' + flowNote + '</div>';
  }

  html += '<div class="meta-line">🚐 ' + _egEsc(rego) + (seats ? ' · ' + _egEsc(seats) + '석' : '') + '</div>';
  if(timeStr) html += '<div class="time-line">⏱ ' + _egEsc(timeStr) + '</div>';
  if(kmS && kmE){
    html += '<div class="meta-line">🛣 ' + _egEsc(kmS) + ' → ' + _egEsc(kmE);
    if(kmDiff > 0) html += ' <span class="km-badge">+' + kmDiff + ' km</span>';
    html += '</div>';
  }
  if(hotel) html += '<div class="meta-line">🏨 ' + _egEsc(hotel) + '</div>';
  if(guide || tc) html += '<div class="meta-line">👤 ' + _egEsc(guide) + (tc ? ' · <span class="tc-badge">' + _egEsc(tc) + '</span>' : '') + '</div>';
  html += '</div>';
  html += '<div class="amt">';
  // EG 청구/지급 금액
  if(taAmount !== 0){
    const amtLabel = (cls === 'EG_BILLS_DC_VEH') ? 'EG 지급액' : 'EG 청구액';
    const amtColor = (cls === 'EG_BILLS_DC_VEH') ? '#dc2626' : '#7c3aed';
    html += '<div style="font-size:8.5pt;color:#6b7280;margin-bottom:2px;">' + amtLabel + '</div>';
    html += '<div style="font-size:13pt;font-weight:800;color:' + amtColor + ';">$' + taAmount.toLocaleString('en-AU',{minimumFractionDigits:2, maximumFractionDigits:2}) + '</div>';
    html += '<div style="font-size:8pt;color:#9ca3af;margin-top:4px;border-top:1px solid #e5e7eb;padding-top:4px;">드라이버 지급액</div>';
  }
  html += '<div class="amount' + (drCost < 0 ? ' neg' : '') + '" style="' + (taAmount !== 0 ? 'font-size:11pt;' : '') + '">' + _fmtAmt(drCost) + '</div>';
  if(nightOwn > 0){
    html += '<div class="night-own' + (isPaid ? ' sub' : '') + '">';
    html += (isPaid ? '차주 납입' : '회사 납입') + ' -$' + nightOwn.toLocaleString('en-AU',{minimumFractionDigits:2, maximumFractionDigits:2});
    html += '</div>';
  }
  html += '</div>';
  html += '</div>';

  // Breakdown 테이블 — EG 청구액 + 드라이버 지급액 2-column
  const hasEG = taAmount !== 0 && breakdown.length > 0;
  const hasDR = drBreakdown.length > 0;
  if(hasEG || hasDR){
    const _renderBreakdown = (title, items, total, totalColor) => {
      let h = '<div style="font-size:8pt;color:#6b7280;font-weight:600;margin-bottom:4px;">' + title + '</div>';
      h += '<table style="width:100%;font-size:9pt;border-collapse:collapse;">';
      items.forEach(it => {
        const sign = it.amount < 0 ? '-' : '+';
        const absAmt = Math.abs(it.amount);
        const color = it.amount < 0 ? '#dc2626' : '#374151';
        h += '<tr>';
        h += '<td style="padding:2px 4px 2px 0;color:#4b5563;">' + _egEsc(it.label);
        if(it.note) h += ' <span style="color:#9ca3af;font-size:8pt;">' + _egEsc(it.note) + '</span>';
        h += '</td>';
        h += '<td style="padding:2px 0;text-align:right;color:' + color + ';font-variant-numeric:tabular-nums;white-space:nowrap;">' + sign + '$' + absAmt.toLocaleString('en-AU',{minimumFractionDigits:2,maximumFractionDigits:2}) + '</td>';
        h += '</tr>';
      });
      h += '<tr style="border-top:1px solid #d1d5db;"><td style="padding:4px 0 2px;font-weight:700;color:#1f2937;">합계</td>';
      h += '<td style="padding:4px 0 2px;text-align:right;font-weight:700;color:' + totalColor + ';font-variant-numeric:tabular-nums;white-space:nowrap;">$' + total.toLocaleString('en-AU',{minimumFractionDigits:2,maximumFractionDigits:2}) + '</td></tr>';
      h += '</table>';
      return h;
    };

    html += '<div style="background:#f9fafb;border-top:1px solid #e5e7eb;padding:8px 12px;margin-top:8px;border-radius:0 0 6px 6px;display:flex;gap:12px;flex-wrap:wrap;">';
    if(hasEG){
      const egTitle = (cls === 'EG_BILLS_DC_VEH') ? '📊 EG 지급액 산출 근거' : '📊 EG 청구액 산출 근거';
      const egColor = (cls === 'EG_BILLS_DC_VEH') ? '#dc2626' : '#7c3aed';
      html += '<div style="flex:1;min-width:240px;">';
      html += _renderBreakdown(egTitle, breakdown, taAmount, egColor);
      html += '</div>';
    }
    if(hasDR){
      html += '<div style="flex:1;min-width:240px;">';
      html += _renderBreakdown('📊 드라이버 지급액 산출 근거', drBreakdown, drCost, '#16a34a');
      html += '</div>';
    }
    html += '</div>';
  }
  html += '</div>';
  return html;
}

// ── 종료 투어 카드 빌더 (TourCode별로 일정/금액/포함 DR 상세) ──
function _egTourCompletionCardHTML(t){
  // t = {tourCode, agency, startDate, endDate, status, guide, pax, reason, drs}
  const drs = (t.drs || []).slice().sort((a,b) => (a._iso||'').localeCompare(b._iso||''));
  const totalDR = drs.reduce((s,r) => s + (Number(r.DR_Cost||r.Total||0) || 0), 0);
  const totalTA = drs.reduce((s,r) => s + _egCalcEgSubAmount(r).total, 0);
  const days = drs.length > 0
    ? (function(){
        const isos = drs.map(r=>r._iso).filter(Boolean).sort();
        return isos.length > 0 ? (isos[isos.length-1] === isos[0] ? 1 :
          Math.round((new Date(isos[isos.length-1]) - new Date(isos[0])) / 86400000) + 1) : 0;
      })()
    : 0;

  let html = '<div class="tour-card">';
  // 헤더 (보라 그라디언트)
  html += '<div class="head">';
  html += '<div class="row1">';
  html += '<div class="left">';
  html += '<div class="tc-name">🎫 ' + _egEsc(t.tourCode) + '</div>';
  html += '<div class="agency">' + _egEsc(t.agency || '—') + '</div>';
  html += '</div>';
  html += '<div class="right">';
  html += '<div class="period">' + _egFmtDate(t.startDate) + ' ~ ' + _egFmtDate(t.endDate) + '</div>';
  if(days > 0) html += '<div class="agency">총 ' + days + '일 · DR ' + drs.length + '건</div>';
  html += '</div>';
  html += '</div>';

  // 부가 정보 행 (가이드, Pax, 상태)
  html += '<div class="meta-row">';
  if(t.guide) html += '<div>👤 ' + _egEsc(t.guide) + '</div>';
  if(t.pax) html += '<div>👥 Pax ' + _egEsc(t.pax) + '</div>';
  if(t.status) html += '<div>📌 ' + _egEsc(t.status) + '</div>';
  html += '</div>';
  html += '</div>';

  // 본문 (종료 사유 + DR 리스트 + 합계)
  html += '<div class="body">';
  html += '<div class="reason">✅ 종료 사유: ' + _egEsc(t.reason) + '</div>';

  if(drs.length > 0){
    html += '<div class="dr-list">';
    // 컬럼 헤더
    html += '<div class="dr-row" style="background:#f3f4f6;font-weight:700;font-size:8.5pt;">';
    html += '<div class="dt">날짜</div>';
    html += '<div class="info">차량 · 드라이버 · 코스 · 시간</div>';
    html += '<div class="amt" style="color:#7c3aed;">EG 인보이스</div>';
    html += '<div class="amt-dr" style="color:#16a34a;">드라이버</div>';
    html += '</div>';
    drs.forEach(r => {
      const dr = Number(r.DR_Cost || r.Total || 0) || 0;
      const ta = _egCalcEgSubAmount(r).total;
      const rego = r.Rego || '';
      const driver = r.Driver || '';
      const attraction = r.Attraction || r.Course || '';
      const tS = _egFmtTime(r.Time_Start || r.Start_Time);
      const tE = _egFmtTime(r.Time_End || r.End_Time);
      const timeStr = tS ? (tS + (tE ? '~' + tE : '')) : '';
      html += '<div class="dr-row">';
      html += '<div class="dt">' + _egFmtDate(r._iso) + '</div>';
      html += '<div class="info"><b>' + _egEsc(rego) + '</b> · ' + _egEsc(driver) +
              (attraction ? ' · ' + _egEsc(attraction) : '') +
              (timeStr ? ' <span style="color:#4f46e5;">' + _egEsc(timeStr) + '</span>' : '') + '</div>';
      html += '<div class="amt">$' + ta.toLocaleString('en-AU',{minimumFractionDigits:2, maximumFractionDigits:2}) + '</div>';
      html += '<div class="amt-dr">$' + dr.toLocaleString('en-AU',{minimumFractionDigits:2, maximumFractionDigits:2}) + '</div>';
      html += '</div>';
    });
    html += '</div>';
  } else {
    html += '<div class="empty" style="padding:14px;">이 투어코드에는 DR 기록이 없습니다.</div>';
  }

  // 합계 (EG 인보이스 + 드라이버 둘 다)
  html += '<div class="totals">';
  html += '<div class="label">💰 투어 합계<br><span style="font-size:8.5pt;color:#9ca3af;font-weight:400;">DR ' + drs.length + '건</span></div>';
  html += '<div class="val-ta">EG 인보이스<br><span style="font-size:13pt;color:#7c3aed;font-weight:800;">$' + totalTA.toLocaleString('en-AU',{minimumFractionDigits:2, maximumFractionDigits:2}) + '</span></div>';
  html += '<div class="val">드라이버<br><span style="font-size:13pt;color:#16a34a;font-weight:800;">$' + totalDR.toLocaleString('en-AU',{minimumFractionDigits:2, maximumFractionDigits:2}) + '</span></div>';
  html += '</div>';

  html += '</div>'; // body
  html += '</div>'; // tour-card
  return html;
}

// ── 일일 리포트 HTML 빌더 (재작성 — 카드 스타일) ─────────────────────────
function _egBuildDailyReportHTML(targetDateISO, drs, newlyCompletedTours){
  _egResetTACache();  // 매 요청 신선한 캐시
  
  // 분류별 그룹핑
  const claims = [];   // 섹션 1: EG 청구 (EG가 받을 돈) — DC_BILLS_EG_VEH + EG_BILLS_OWN
  const payments = []; // 섹션 2: EG 지급 (EG가 줄 돈) — EG_BILLS_DC_VEH
  drs.forEach(r => {
    const cls = _egClassifyRow(r);
    if(cls === 'DC_BILLS_EG_VEH' || cls === 'EG_BILLS_OWN'){
      claims.push(r);
    } else if(cls === 'EG_BILLS_DC_VEH'){
      payments.push(r);
    }
  });

  const totalClaim = claims.reduce((s,r) => s + _egCalcEgSubAmount(r).total, 0);
  const totalPay = payments.reduce((s,r) => s + _egCalcEgSubAmount(r).total, 0);
  const totalDR = drs.reduce((s,r) => s + (Number(r.DR_Cost||r.Total||0) || 0), 0);
  const tcCount = new Set(drs.map(r => String(r.Tour_Code||r.TourCode||'').trim()).filter(Boolean)).size;

  let html = '<html><head><meta charset="UTF-8">' + _egCommonStyle() + '</head><body>';
  html += '<div class="hdr"><h1>📋 EG TRAVEL 일일 운행 리포트</h1>';
  html += '<div class="sub">대상일: <b>' + _egFmtDate(targetDateISO) + '</b> · 발행: ' + _egFmtDate(_egTodaySydney()) + '</div></div>';

  // 메타 박스 — 4컬럼
  html += '<div class="meta">';
  html += '<div><div style="font-size:9pt;color:#6b7280;">운행 건수</div><div style="font-size:14pt;font-weight:bold;color:#1f2937;">' + drs.length + '건</div></div>';
  html += '<div><div style="font-size:9pt;color:#6b7280;">투어코드</div><div style="font-size:14pt;font-weight:bold;color:#5b21b6;">' + tcCount + '개</div></div>';
  html += '<div><div style="font-size:9pt;color:#6b7280;">EG 청구액 (받을 돈)</div><div style="font-size:14pt;font-weight:bold;color:#7c3aed;">$' + totalClaim.toLocaleString('en-AU',{minimumFractionDigits:2, maximumFractionDigits:2}) + '</div></div>';
  html += '<div><div style="font-size:9pt;color:#6b7280;">EG 지급액 (줄 돈)</div><div style="font-size:14pt;font-weight:bold;color:#dc2626;">$' + totalPay.toLocaleString('en-AU',{minimumFractionDigits:2, maximumFractionDigits:2}) + '</div></div>';
  html += '</div>';

  // 시간순 정렬 헬퍼
  const _byTime = (a,b) => String(a.Time_Start||a.Start_Time||'').localeCompare(String(b.Time_Start||b.Start_Time||''));

  // ── 안전망: 트레일러 정보 불일치 감지 ──
  // PD에 트레일러 있는데 DR에 없음 → 청구 누락 의심
  // DR에 트레일러 비용 있는데 PD/DR Trailer_Number 없음 → 소유주 미상
  const trailerWarnings = [];
  drs.forEach(r => {
    const drTrailerCost = Number(r.Trailer||0);
    const drTrailerNum = String(r.Trailer_Number||'').trim();
    // DR의 Trailer_Number가 비어있으면 PD에서 조회
    let pdTrailer = '';
    try {
      const pdResult = lookupTrailerForDR({
        date: r._iso,
        driver: String(r.Driver||'').trim(),
        rego: String(r.Rego||'').trim()
      });
      if(pdResult && pdResult.ok) pdTrailer = pdResult.pdTrailer || '';
    } catch(e){}

    // 케이스 A: PD에 트레일러 있는데 DR Trailer 비용 0 → 청구 누락 의심
    if(pdTrailer && drTrailerCost === 0){
      trailerWarnings.push({
        type: 'missing_charge',
        date: r._iso,
        driver: String(r.Driver||''),
        rego: String(r.Rego||''),
        agency: String(r.Agency||''),
        pdTrailer: pdTrailer,
        msg: 'PD에 트레일러 [' + pdTrailer + '] 픽업 기록이 있으나 DR에 트레일러 비용이 없습니다. 청구 누락 가능성.'
      });
    }
    // 케이스 B: DR에 트레일러 비용 있는데 Trailer_Number 미상
    if(drTrailerCost > 0 && !drTrailerNum && !pdTrailer){
      trailerWarnings.push({
        type: 'missing_number',
        date: r._iso,
        driver: String(r.Driver||''),
        rego: String(r.Rego||''),
        agency: String(r.Agency||''),
        cost: drTrailerCost,
        msg: '트레일러 비용 $' + drTrailerCost + ' 입력됐으나 트레일러 번호 미상. 소유주 식별 불가 → 대여비 차감 누락.'
      });
    }
  });

  if(trailerWarnings.length > 0){
    html += '<div style="background:#fef3c7;border:2px solid #f59e0b;border-radius:8px;padding:12px;margin:12px 0;">';
    html += '<div style="font-size:11pt;font-weight:700;color:#92400e;margin-bottom:6px;">⚠️ 검토 필요 (트레일러 정보 불일치) — ' + trailerWarnings.length + '건</div>';
    html += '<div style="font-size:9pt;color:#78350f;margin-bottom:8px;">아래 운행은 트레일러 정보가 일치하지 않습니다. 정산 전 확인이 필요합니다.</div>';
    trailerWarnings.forEach(w => {
      html += '<div style="background:#fffbeb;padding:6px 10px;margin-bottom:4px;border-radius:4px;font-size:9.5pt;">';
      html += '<b>' + _egFmtDate(w.date) + '</b> · ' + _egEsc(w.driver) + ' / ' + _egEsc(w.rego);
      if(w.agency) html += ' · ' + _egEsc(w.agency);
      html += '<br><span style="color:#92400e;">' + _egEsc(w.msg) + '</span>';
      html += '</div>';
    });
    html += '</div>';
  }

  // ── 섹션 1: EG 청구 내역 (받을 돈) ──
  html += '<div class="sec-title">💰 EG 청구 내역 (받을 돈)</div>';
  if(claims.length === 0){
    html += '<div class="empty">청구 대상 운행이 없습니다.</div>';
  } else {
    claims.sort(_byTime).forEach(r => { html += _egTripCardHTML(r); });
  }

  // ── 섹션 2: EG 지급 내역 (줄 돈) ──
  html += '<div class="sec-title">💸 EG 지급 내역 (줄 돈)</div>';
  if(payments.length === 0){
    html += '<div class="empty">지급 대상 운행이 없습니다.</div>';
  } else {
    payments.sort(_byTime).forEach(r => { html += _egTripCardHTML(r); });
  }

  // ── 섹션 3: 차주별 드라이버 지급 요약 ──
  html += '<div class="sec-title">👥 차주별 드라이버 지급 요약</div>';
  const vehOwners = _egLoadVehicleOwners();
  const ownerDriverSum = {};  // {ownerKey: {driverName: amount}}
  const ownerDisplay = {};    // {ownerKey: displayName}
  drs.forEach(r => {
    const rego = String(r.Rego||'').trim();
    const rawOwner = vehOwners[rego] || 'Unknown';
    const ownerNorm = _egNormEntity(rawOwner) || rawOwner;
    const driver = String(r.Driver || 'Unknown').trim();
    const cost = Number(r.DR_Cost || r.Total || 0) || 0;
    if(!ownerDriverSum[ownerNorm]) ownerDriverSum[ownerNorm] = {};
    if(!ownerDriverSum[ownerNorm][driver]) ownerDriverSum[ownerNorm][driver] = 0;
    ownerDriverSum[ownerNorm][driver] += cost;
    // 표시명 (풀네임 우선)
    if(!ownerDisplay[ownerNorm] || ownerDisplay[ownerNorm].length < rawOwner.length){
      ownerDisplay[ownerNorm] = rawOwner;
    }
  });
  const ownerKeys = Object.keys(ownerDriverSum);
  if(ownerKeys.length === 0){
    html += '<div class="empty">드라이버 지급 내역이 없습니다.</div>';
  } else {
    ownerKeys.sort().forEach(ownerKey => {
      const drivers = ownerDriverSum[ownerKey];
      const ownerTot = Object.keys(drivers).reduce((s,d) => s + drivers[d], 0);
      html += '<div class="driver-grp">';
      html += '<div class="hdr-bar">';
      html += '<div>🏢 ' + _egEsc(ownerDisplay[ownerKey]) + '</div>';
      html += '<div class="right"><span style="color:#fbbf24;">합계: $' + ownerTot.toLocaleString('en-AU',{minimumFractionDigits:2, maximumFractionDigits:2}) + '</span></div>';
      html += '</div>';
      // 드라이버별 줄
      const driverNames = Object.keys(drivers).sort((a,b) => drivers[b] - drivers[a]);
      driverNames.forEach(dn => {
        html += '<div style="display:flex;justify-content:space-between;padding:8px 14px;border-bottom:1px solid #f3f4f6;font-size:10pt;">';
        html += '<div>👤 ' + _egEsc(dn) + '</div>';
        html += '<div style="font-weight:700;color:#16a34a;">$' + drivers[dn].toLocaleString('en-AU',{minimumFractionDigits:2, maximumFractionDigits:2}) + '</div>';
        html += '</div>';
      });
      html += '</div>';
    });
  }

  // ── 섹션 4: 새로 종료된 투어코드 ──
  html += '<div class="sec-title">✅ 새로 종료된 투어코드</div>';
  if(newlyCompletedTours.length === 0){
    html += '<div class="empty">새로 종료된 투어가 없습니다.</div>';
  } else {
    newlyCompletedTours.sort((a,b) => (b.endDate||'').localeCompare(a.endDate||''));
    newlyCompletedTours.forEach(t => {
      html += _egTourCompletionCardHTML(t);
    });
  }

  html += '<div class="ftr">Dong Choi Pty Ltd · 자동 생성 리포트 · 문의: ' + EG_REPORT_ADMIN_BCC + '</div>';
  html += '</body></html>';
  return html;
}

// ── 주간 리포트 HTML 빌더 ─────────────────────────────────────────────────
// 1) 운행 통계, 2) 운행 상세 테이블, 3) 드라이버별 지급액
function _egBuildWeeklyReportHTML(monISO, sunISO, drs){
  _egResetTACache();

  // 분류
  const claims = [];   // 받을 돈: DC_BILLS_EG_VEH + EG_BILLS_OWN
  const payments = []; // 줄 돈: EG_BILLS_DC_VEH
  drs.forEach(r => {
    const cls = _egClassifyRow(r);
    if(cls === 'DC_BILLS_EG_VEH' || cls === 'EG_BILLS_OWN') claims.push(r);
    else if(cls === 'EG_BILLS_DC_VEH') payments.push(r);
  });

  const totalClaim = claims.reduce((s,r)=>s + _egCalcEgSubAmount(r).total, 0);
  const totalPay = payments.reduce((s,r)=>s + _egCalcEgSubAmount(r).total, 0);
  const totalDR = drs.reduce((s,r)=>s+(Number(r.DR_Cost||r.Total||0)||0), 0);
  const tcSet = new Set();
  const vehCount = {};
  drs.forEach(r => {
    const tc = String(r.Tour_Code||r.TourCode||'').trim();
    if(tc) tcSet.add(tc);
    const veh = String(r.Rego||'').trim() || '?';
    vehCount[veh] = (vehCount[veh]||0) + 1;
  });

  let html = `<html><head><meta charset="UTF-8">${_egCommonStyle()}</head><body>`;
  html += `<div class="hdr"><h1>📊 EG TRAVEL 주간 운행 요약</h1>
            <div class="sub">기간: ${_egFmtDate(monISO)} ~ ${_egFmtDate(sunISO)} · 발행: ${_egFmtDate(_egTodaySydney())}</div></div>`;

  // 통계 박스
  html += `<div class="summary-box">
    <div class="summary-row"><div>📋 운행 건수</div><div>${drs.length}건</div></div>
    <div class="summary-row"><div>🎫 투어코드 수</div><div>${tcSet.size}개</div></div>
    <div class="summary-row"><div>🚐 차량 수</div><div>${Object.keys(vehCount).length}대</div></div>
    <div class="summary-row tot"><div style="color:#7c3aed;">💰 EG 청구액 (받을 돈)</div>
      <div style="color:#7c3aed;font-size:13pt;">$${totalClaim.toLocaleString('en-AU',{minimumFractionDigits:2,maximumFractionDigits:2})}</div></div>
    <div class="summary-row tot"><div style="color:#dc2626;">💸 EG 지급액 (줄 돈)</div>
      <div style="color:#dc2626;font-size:13pt;">$${totalPay.toLocaleString('en-AU',{minimumFractionDigits:2,maximumFractionDigits:2})}</div></div>
    <div class="summary-row tot"><div style="color:#16a34a;">💵 드라이버 지급액 합계</div>
      <div style="color:#16a34a;font-size:13pt;">$${totalDR.toLocaleString('en-AU',{minimumFractionDigits:2,maximumFractionDigits:2})}</div></div>
  </div>`;

  // 운행 행 렌더 헬퍼 (청구/지급 공통)
  const _renderTripRow = (r) => {
    const ta = _egCalcEgSubAmount(r).total;
    const dr = Number(r.DR_Cost||r.Total||0)||0;
    const cls = _egClassifyRow(r);
    const agency = String(r.Agency||r.Tour_Agency||'').trim();
    let target = '';
    if(cls === 'DC_BILLS_EG_VEH') target = 'DC';
    else if(cls === 'EG_BILLS_DC_VEH') target = agency || '여행사';
    else if(cls === 'EG_BILLS_OWN') target = agency || '여행사';
    const amtColor = (cls === 'EG_BILLS_DC_VEH') ? '#dc2626' : '#7c3aed';
    return `<tr>
      <td>${_egFmtDate(r._iso)}</td>
      <td><span class="tc-badge">${_egEsc(r.Tour_Code||r.TourCode||'')}</span></td>
      <td>${_egEsc(r.Rego||'')}</td>
      <td>${_egEsc(r.Driver||'')}</td>
      <td>${_egEsc(target)}</td>
      <td>${_egEsc(r.Attraction||r.Course||'')}</td>
      <td class="num" style="color:${amtColor};font-weight:700;">$${ta.toLocaleString('en-AU',{minimumFractionDigits:2,maximumFractionDigits:2})}</td>
      <td class="num" style="color:#16a34a;font-weight:700;">$${dr.toLocaleString('en-AU',{minimumFractionDigits:2,maximumFractionDigits:2})}</td>
    </tr>`;
  };

  const _tableHead = (amtLabel, amtColor) => `<table>
      <tr><th>날짜</th><th>TourCode</th><th>차량</th><th>드라이버</th><th>대상</th>
          <th>코스</th><th class="num" style="color:${amtColor};">${amtLabel}</th><th class="num">드라이버</th></tr>`;
  const _totalRow = (label, ta, dr, color) => `<tr style="background:#f9fafb;">
      <td colspan="6"><b>${label}</b></td>
      <td class="num" style="color:${color};font-weight:800;">$${ta.toLocaleString('en-AU',{minimumFractionDigits:2,maximumFractionDigits:2})}</td>
      <td class="num" style="color:#16a34a;font-weight:800;">$${dr.toLocaleString('en-AU',{minimumFractionDigits:2,maximumFractionDigits:2})}</td>
    </tr>`;

  // ── 섹션 1: EG 청구 내역 (받을 돈) ──
  html += `<div class="sec-title">💰 EG 청구 내역 (받을 돈)</div>`;
  if(claims.length === 0){
    html += '<div class="empty">청구 대상 운행이 없습니다.</div>';
  } else {
    html += _tableHead('EG 청구액', '#7c3aed');
    claims.slice().sort((a,b)=>(a._iso||'').localeCompare(b._iso||'')).forEach(r => {
      html += _renderTripRow(r);
    });
    const claimDR = claims.reduce((s,r)=>s+(Number(r.DR_Cost||r.Total||0)||0), 0);
    html += _totalRow('합계', totalClaim, claimDR, '#7c3aed');
    html += '</table>';
  }

  // ── 섹션 2: EG 지급 내역 (줄 돈) ──
  html += `<div class="sec-title">💸 EG 지급 내역 (줄 돈)</div>`;
  if(payments.length === 0){
    html += '<div class="empty">지급 대상 운행이 없습니다.</div>';
  } else {
    html += _tableHead('EG 지급액', '#dc2626');
    payments.slice().sort((a,b)=>(a._iso||'').localeCompare(b._iso||'')).forEach(r => {
      html += _renderTripRow(r);
    });
    const payDR = payments.reduce((s,r)=>s+(Number(r.DR_Cost||r.Total||0)||0), 0);
    html += _totalRow('합계', totalPay, payDR, '#dc2626');
    html += '</table>';
  }

  // ── 섹션 3: 차주별 드라이버 지급 요약 ──
  html += `<div class="sec-title">👥 차주별 드라이버 지급 요약</div>`;
  const vehOwners = _egLoadVehicleOwners();
  const ownerDriverSum = {};  // {ownerKey: {driverName: {amount, count}}}
  const ownerDisplay = {};
  drs.forEach(r => {
    const rego = String(r.Rego||'').trim();
    const rawOwner = vehOwners[rego] || 'Unknown';
    const ownerNorm = _egNormEntity(rawOwner) || rawOwner;
    const driver = String(r.Driver||'').trim() || 'Unknown';
    const cost = Number(r.DR_Cost||r.Total||0)||0;
    if(!ownerDriverSum[ownerNorm]) ownerDriverSum[ownerNorm] = {};
    if(!ownerDriverSum[ownerNorm][driver]) ownerDriverSum[ownerNorm][driver] = {amount:0, count:0};
    ownerDriverSum[ownerNorm][driver].amount += cost;
    ownerDriverSum[ownerNorm][driver].count += 1;
    if(!ownerDisplay[ownerNorm] || ownerDisplay[ownerNorm].length < rawOwner.length){
      ownerDisplay[ownerNorm] = rawOwner;
    }
  });
  const ownerKeys = Object.keys(ownerDriverSum).sort();
  if(ownerKeys.length === 0){
    html += '<div class="empty">드라이버 지급 내역이 없습니다.</div>';
  } else {
    html += `<table>
      <tr><th>차주</th><th>드라이버</th><th class="num">운행 건수</th><th class="num">지급액</th></tr>`;
    ownerKeys.forEach(ownerKey => {
      const drivers = ownerDriverSum[ownerKey];
      const driverNames = Object.keys(drivers).sort((a,b) => drivers[b].amount - drivers[a].amount);
      const ownerTot = driverNames.reduce((s,d)=>s+drivers[d].amount, 0);
      const ownerCnt = driverNames.reduce((s,d)=>s+drivers[d].count, 0);
      driverNames.forEach((dn, idx) => {
        html += `<tr>
          <td>${idx === 0 ? '🏢 ' + _egEsc(ownerDisplay[ownerKey]) : ''}</td>
          <td>${_egEsc(dn)}</td>
          <td class="num">${drivers[dn].count}건</td>
          <td class="num" style="color:#16a34a;font-weight:700;">$${drivers[dn].amount.toLocaleString('en-AU',{minimumFractionDigits:2,maximumFractionDigits:2})}</td>
        </tr>`;
      });
      html += `<tr style="background:#f0fdf4;">
        <td colspan="2"><b>${_egEsc(ownerDisplay[ownerKey])} 소계</b></td>
        <td class="num"><b>${ownerCnt}건</b></td>
        <td class="num" style="color:#16a34a;font-weight:800;">$${ownerTot.toLocaleString('en-AU',{minimumFractionDigits:2,maximumFractionDigits:2})}</td>
      </tr>`;
    });
    html += `</table>`;
  }

  html += `<div class="ftr">Dong Choi Pty Ltd · 자동 생성 리포트 · 문의: ${EG_REPORT_ADMIN_BCC}</div>`;
  html += '</body></html>';
  return html;
}

function _egEsc(s){
  if(s === null || s === undefined) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ── 시간 정규화 (HH:MM) ─────────────────────────────────────────────────
// Google Sheets는 시간 셀("08:30")을 Date 객체(1899-12-30T08:30:00)로 반환하므로
// String() 변환 시 "Sat Dec 30 1899 08:30:00 GMT+1000" 같이 깨짐.
// 이 헬퍼는 Date 객체에서 HH:MM만 추출하고, 이미 문자열이면 그대로 정리.
function _egFmtTime(v){
  if(v === null || v === undefined || v === '') return '';
  // Date 객체 — Sydney 타임존 기준 HH:MM 추출
  if(v instanceof Date && !isNaN(v.getTime())){
    return Utilities.formatDate(v, 'Australia/Sydney', 'HH:mm');
  }
  const s = String(v).trim();
  if(!s) return '';
  // Sheets에서 종종 "Sat Dec 30 1899 08:30:00 GMT+1000" 형태로 들어옴
  const m1 = s.match(/\b(\d{1,2}):(\d{2})(?::\d{2})?\b/);
  if(m1) return m1[1].padStart(2,'0') + ':' + m1[2];
  // 이미 정상 "08:30" 또는 "8:30"
  return s;
}

// ─── M_PriceClient 기반 TA(여행사 청구) 계산 ──────────────────────────────
// admin.html의 calcAgencyTA 로직을 GAS로 이식.
// 사용 시점에 M_PriceClient 시트를 1회 로드하여 메모리 캐시 (요청당)
let _egPriceClientCache = null;
function _egLoadPriceClient(){
  if(_egPriceClientCache !== null) return _egPriceClientCache;
  _egPriceClientCache = {};
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('M_PriceClient');
    if(!sheet) return _egPriceClientCache;
    const data = sheet.getDataRange().getValues();
    if(data.length < 2) return _egPriceClientCache;
    const headers = data[0].map(String);
    const agIdx = headers.indexOf('Agency');
    const courseIdx = headers.indexOf('Course');
    const mhIdx = headers.indexOf('max_hours');
    // 좌석별 rate / ot 컬럼
    const seatCols = {};
    ['21','25','40','50'].forEach(s => {
      seatCols[s] = {
        rate: headers.indexOf('seats_' + s + '_rate'),
        ot: headers.indexOf('seats_' + s + '_ot')
      };
    });
    for(let i=1; i<data.length; i++){
      const row = data[i];
      const ag = String(row[agIdx]||'').trim();
      const course = String(row[courseIdx]||'').trim();
      if(!ag || !course) continue;
      if(!_egPriceClientCache[ag]) _egPriceClientCache[ag] = {};
      const entry = { max_hours: Number(row[mhIdx])||0 };
      ['21','25','40','50'].forEach(s => {
        entry[s] = {
          rate: Number(row[seatCols[s].rate])||0,
          ot:   Number(row[seatCols[s].ot])||0
        };
      });
      _egPriceClientCache[ag][course] = entry;
    }
  } catch(e){
    Logger.log('_egLoadPriceClient error: ' + e);
  }
  return _egPriceClientCache;
}

// 좌석별 트레일러 DR→TA 변환 (admin.html _trailerDRtoTA 이식)
function _egTrailerDRtoTA(dr, sn){
  if(!dr || dr === 0) return 0;
  if(sn >= 40) return 0;  // 40/50석은 트레일러 TA 0
  return dr === 30 ? 80 : Math.round(dr * 2.67);
}

// 여행사 TA 청구금액 계산 — admin.html calcAgencyTA 전체 로직 이식
function _egCalcAgencyTA(r){
  const PC = _egLoadPriceClient();
  const agency = String(r.Agency||r.agency||'').trim();
  const attraction = String(r.Attraction||r.tour||'').trim();
  const seatsRaw = String(r.Seats||r.seats||'').replace('S','').trim();
  const capNum = parseInt(seatsRaw)||25;
  const capKey = capNum>=50?'50':capNum>=40?'40':capNum>=25?'25':'21';
  const isLarge = capNum>=40;
  const svc = Number(r.SVC_Charge)||0;

  // 1) M_PriceClient base rate 조회 (대소문자 무시 fallback)
  const agPC = PC[agency];
  function _findCourse(pc, cn){
    if(!pc || !cn) return null;
    if(pc[cn]) return pc[cn];
    const lc = cn.toLowerCase();
    const keys = Object.keys(pc);
    for(let i=0; i<keys.length; i++){
      if(keys[i].toLowerCase() === lc) return pc[keys[i]];
    }
    return null;
  }
  const ce = agPC ? _findCourse(agPC, attraction) : null;
  let taBase = 0;
  if(ce){ const sd = ce[capKey] || ce['21']; taBase = Number(sd && sd.rate) || 0; }
  if(taBase === 0) taBase = svc;

  // 2) 서차지 DR→TA 역산
  const htl = Number(r.Hotel_Surcharge||r.hotel||0);
  const dst = Number(r.Dist_Surcharge||r.dist||0);
  const ot  = Number(r.OT||r.ot||0);
  const erl = Number(r.Early||0);
  const toll= Number(r.Toll||0);

  const htlTA = htl===0?0:
    capNum>=50?(htl===15?80:htl===30?160:htl*4):
    capNum>=40?(htl===15?75:htl===30?150:htl*4):
    (htl===10?40:htl===20?80:htl*4);
  const dstTA = dst===0?0:
    capNum>=50?(dst===40?160:dst*4):
    capNum>=40?(dst===40?150:Math.round(dst*3.75)):
    (dst===30?80:Math.round(dst*2.67));

  // OT TA — Tour Hojuro / Plus Australia 21/25S: 30분 UNIT 단위
  const otRateTA = capNum>=50?160:capNum>=40?150:80;
  const otRateDR = capNum>=40?40:30;
  const isHojuroOT = /호주로|hojuro|plus\s*australia/i.test(agency);
  const otHrs = isHojuroOT ? (ot / (otRateDR/2)) * 0.5 : (otRateDR>0 ? ot/otRateDR : 0);
  const otTA = isHojuroOT
    ? Math.round((ot / (otRateDR/2)) * (otRateTA/2))
    : Math.round(otHrs * otRateTA);

  // Early TA
  let erlTA = 0;
  if(erl > 0){
    const isHojuroEarly = /호주로|hojuro|plus\s*australia/i.test(agency);
    if(isHojuroEarly && capNum < 40){
      erlTA = 80;
    } else {
      let atE = agPC ? _findCourse(agPC, 'Airport Transfer') : null;
      if(!atE){
        const allAgs = Object.keys(PC);
        for(let i=0; i<allAgs.length; i++){
          atE = _findCourse(PC[allAgs[i]], 'Airport Transfer');
          if(atE) break;
        }
      }
      if(atE){
        const sd2 = atE[capKey] || atE['21'];
        erlTA = Math.round((Number(sd2 && sd2.rate)||0) * 0.3);
      }
    }
  }

  // Parking (공항 픽업) — Tour Hojuro / Plus Australia 21/25S 제외
  const apPat = /\b(airport|syd|kingsford|mascot|international|domestic|terminal)\b/i;
  const isHojuro = /호주로|hojuro|plus\s*australia/i.test(agency);
  const pickup = String(r.Pickup||'');
  const parking = (apPat.test(pickup) && !(isHojuro && !isLarge))
    ? (isLarge ? 40 : 30)
    : 0;

  // Toll (대형 버스만 TA에 포함)
  const tollTA = isLarge ? toll : 0;

  // Trailer
  const trl = Number(r.Trailer||0);
  const trlTA = _egTrailerDRtoTA(trl, capNum);

  return taBase + otTA + htlTA + dstTA + erlTA + parking + tollTA + trlTA;
}

// 캐시 무효화 (매 요청 시작 시 호출 — M_PriceClient 변경 반영)
function _egResetTACache(){
  _egPriceClientCache = null;
  _egVehicleOwnerCache = null;  // 차량 소유주 캐시도 함께 무효화
  _egPriceSubCache = null;      // SUB 가격 캐시
  _egTrailerOwnerCache = null;  // 트레일러 소유주 캐시
  _egPriceDriverCache = null;   // 드라이버 가격 캐시
}

// ── 발송 (공통) — HTML을 PDF로 첨부하여 Gmail로 발송 ──────────────────────
function _egSendEmailWithPDF(subject, bodyText, docHtml, pdfName, recipients){
  if(!recipients || !recipients.to){
    return { ok: false, error: 'no_recipient', message: 'EG TRAVEL 이메일이 M_Clients에 등록되지 않음' };
  }
  try {
    const htmlBlob = Utilities.newBlob(docHtml, 'text/html', 'report.html');
    const pdfBlob = htmlBlob.getAs('application/pdf').setName(pdfName);
    const options = {
      name: 'Dong Choi Pty Ltd',
      attachments: [pdfBlob],
      bcc: recipients.bcc || EG_REPORT_ADMIN_BCC
    };
    if(recipients.cc) options.cc = recipients.cc;
    GmailApp.sendEmail(recipients.to, subject, bodyText, options);
    return { ok: true };
  } catch(err){
    Logger.log('EG send email error: ' + err);
    return { ok: false, error: err.toString() };
  }
}

// ── 메인: 일일 리포트 발송 ────────────────────────────────────────────────
function sendEGDailyReport(opts){
  opts = opts || {};
  const dryRun = !!opts.dryRun;
  const targetDate = opts.date || _egYesterdaySydney(); // 기본: 전날
  const todayISO = _egTodaySydney();

  try {
    // 1. 전날 DR 로드
    const drs = _egLoadDRs(targetDate, targetDate);

    // 2. 새로 종료된 투어코드 (이미 발송 이력에 있는 것 제외)
    const allCompleted = _egFindCompletedTourCodes(todayISO);
    const alreadySent = _egGetAlreadySentTourCodes();
    const newCompleted = allCompleted.filter(t => !alreadySent.has(t.tourCode.toUpperCase()));

    // 3. 전날 DR이 0건이면 skip (DR이 있을 때만 발송)
    if(drs.length === 0){
      Logger.log('[EG Daily] skip — 전날 DR 0건 (운행 없음)');
      return { ok: true, skipped: true, reason: 'no_dr' };
    }

    // 4. HTML 빌드
    const html = _egBuildDailyReportHTML(targetDate, drs, newCompleted);
    const pdfName = 'EG_Daily_Report_' + targetDate + '.pdf';
    const subject = `[EG TRAVEL] 일일 운행 리포트 ${_egFmtDate(targetDate)} — DR ${drs.length}건, 종료 투어 ${newCompleted.length}개`;
    const body = `안녕하세요,\n\n` +
      `${_egFmtDate(targetDate)} EG TRAVEL 관련 운행 리포트를 첨부합니다.\n\n` +
      `· 전일 운행 건수: ${drs.length}건\n` +
      `· 새로 종료된 투어: ${newCompleted.length}개\n\n` +
      `상세 내용은 첨부 PDF를 확인해주세요.\n\n` +
      `Kind regards,\nDong Choi Pty Ltd`;

    if(dryRun){
      return { ok: true, dryRun: true, html: html, subject: subject,
               drCount: drs.length, completedCount: newCompleted.length };
    }

    // 5. 발송
    const recipients = _egGetRecipients();
    const sendResult = _egSendEmailWithPDF(subject, body, html, pdfName, recipients);

    // 6. 이력 기록
    const tcCodes = newCompleted.map(t => t.tourCode);
    _egLogReportSent(
      'daily', targetDate, targetDate,
      recipients.to + (recipients.cc?' (cc:'+recipients.cc+')':''),
      tcCodes, subject,
      sendResult.ok ? 'OK' : 'FAILED',
      sendResult.ok ? `DR ${drs.length}건 / 종료투어 ${newCompleted.length}개` : sendResult.error
    );

    return {
      ok: sendResult.ok, error: sendResult.error,
      drCount: drs.length, completedCount: newCompleted.length,
      recipients: recipients
    };
  } catch(err){
    Logger.log('sendEGDailyReport error: ' + err);
    _egLogReportSent('daily', targetDate, targetDate, '', [], '', 'ERROR', err.toString());
    return { ok: false, error: err.toString() };
  }
}

// ── 메인: 주간 리포트 발송 ────────────────────────────────────────────────
function sendEGWeeklyReport(opts){
  opts = opts || {};
  const dryRun = !!opts.dryRun;
  // 기본: 지난주 월~일 (오늘이 월요일이면 지난주 월~일)
  const todayISO = _egTodaySydney();
  const todayMon = _egMondayOf(todayISO);
  const lastMonD = new Date(todayMon + 'T00:00:00');
  lastMonD.setDate(lastMonD.getDate() - 7);
  const lastMonISO = Utilities.formatDate(lastMonD, 'Australia/Sydney', 'yyyy-MM-dd');
  const lastSunD = new Date(lastMonISO + 'T00:00:00');
  lastSunD.setDate(lastSunD.getDate() + 6);
  const lastSunISO = Utilities.formatDate(lastSunD, 'Australia/Sydney', 'yyyy-MM-dd');

  const fromISO = opts.from || lastMonISO;
  const toISO = opts.to || lastSunISO;

  try {
    const drs = _egLoadDRs(fromISO, toISO);
    if(drs.length === 0 && !opts.forceEmpty){
      Logger.log('[EG Weekly] skip — 운행 0건');
      return { ok: true, skipped: true, reason: 'no_data' };
    }

    const html = _egBuildWeeklyReportHTML(fromISO, toISO, drs);
    const pdfName = 'EG_Weekly_Report_' + fromISO + '_to_' + toISO + '.pdf';
    _egResetTACache();
    const totDR = drs.reduce((s,r)=>s+(Number(r.DR_Cost||r.Total||0)||0), 0);
    // 청구/지급 분리 합계
    let wkClaim = 0, wkPay = 0;
    drs.forEach(r => {
      const c = _egClassifyRow(r);
      const amt = _egCalcEgSubAmount(r).total;
      if(c === 'DC_BILLS_EG_VEH' || c === 'EG_BILLS_OWN') wkClaim += amt;
      else if(c === 'EG_BILLS_DC_VEH') wkPay += amt;
    });
    const totTA = wkClaim + wkPay;
    const subject = `[EG TRAVEL] 주간 운행 요약 ${_egFmtDate(fromISO)}~${_egFmtDate(toISO)} — ${drs.length}건, 청구 $${wkClaim.toLocaleString('en-AU',{minimumFractionDigits:2,maximumFractionDigits:2})} / 지급 $${wkPay.toLocaleString('en-AU',{minimumFractionDigits:2,maximumFractionDigits:2})}`;
    const body = `안녕하세요,\n\n` +
      `${_egFmtDate(fromISO)} ~ ${_egFmtDate(toISO)} EG TRAVEL 주간 운행 요약을 첨부합니다.\n\n` +
      `· 총 운행 건수: ${drs.length}건\n` +
      `· EG 청구액 (받을 돈): $${wkClaim.toLocaleString('en-AU',{minimumFractionDigits:2,maximumFractionDigits:2})}\n` +
      `· EG 지급액 (줄 돈): $${wkPay.toLocaleString('en-AU',{minimumFractionDigits:2,maximumFractionDigits:2})}\n` +
      `· 드라이버 지급액 합계: $${totDR.toLocaleString('en-AU',{minimumFractionDigits:2,maximumFractionDigits:2})}\n\n` +
      `상세 운행 내역과 차주별 드라이버 지급액은 첨부 PDF를 참고하세요.\n\n` +
      `Kind regards,\nDong Choi Pty Ltd`;

    if(dryRun){
      return { ok: true, dryRun: true, html: html, subject: subject, drCount: drs.length };
    }

    const recipients = _egGetRecipients();
    const sendResult = _egSendEmailWithPDF(subject, body, html, pdfName, recipients);

    _egLogReportSent(
      'weekly', fromISO, toISO,
      recipients.to + (recipients.cc?' (cc:'+recipients.cc+')':''),
      [], subject,
      sendResult.ok ? 'OK' : 'FAILED',
      sendResult.ok ? `DR ${drs.length}건 · EG $${totTA.toFixed(2)} / 드라이버 $${totDR.toFixed(2)}` : sendResult.error
    );

    return { ok: sendResult.ok, error: sendResult.error,
             drCount: drs.length, totalAmount: totTA, driverTotal: totDR, recipients: recipients };
  } catch(err){
    Logger.log('sendEGWeeklyReport error: ' + err);
    _egLogReportSent('weekly', fromISO, toISO, '', [], '', 'ERROR', err.toString());
    return { ok: false, error: err.toString() };
  }
}

// ── 트리거 설정 (한 번만 실행) ────────────────────────────────────────────
function setupEGReportTriggers(){
  // 기존 EG 리포트 트리거 제거
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  triggers.forEach(t => {
    const fn = t.getHandlerFunction();
    if(fn === 'sendEGDailyReport' || fn === 'sendEGWeeklyReport'){
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  // 매일 06:00 (시드니)
  ScriptApp.newTrigger('sendEGDailyReport')
    .timeBased().atHour(EG_REPORT_DAILY_HOUR).everyDays(1)
    .inTimezone('Australia/Sydney').create();
  // 매주 월요일 06:00
  ScriptApp.newTrigger('sendEGWeeklyReport')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(EG_REPORT_WEEKLY_HOUR)
    .inTimezone('Australia/Sydney').create();
  return { ok: true, removed: removed, created: 2,
           message: '매일 06:00 일일 리포트, 매주 월요일 06:00 주간 리포트 트리거 설정됨' };
}

// ── 트리거 제거 (필요 시) ─────────────────────────────────────────────────
function removeEGReportTriggers(){
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  triggers.forEach(t => {
    const fn = t.getHandlerFunction();
    if(fn === 'sendEGDailyReport' || fn === 'sendEGWeeklyReport'){
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  return { ok: true, removed: removed };
}

// ═══════════════════════════════════════════════════════════════════════════
// 진단 도구 — EG 리포트 누락 케이스 디버깅
// 사용법:
//   1. GAS 에디터에서 debugEGMay22 (또는 debugEGYesterday) 실행 후 로그 확인
//   2. 또는 _egDebugDate('2026-05-22') 직접 호출
// ═══════════════════════════════════════════════════════════════════════════
function debugEGMay22(){ return _egDebugDate('2026-05-22'); }
function debugEGYesterday(){ return _egDebugDate(_egYesterdaySydney()); }
function _egDebugDate(targetISO){
  if(!targetISO){
    // 인자 없이 실행됐을 때 안전 폴백 — 어제 날짜로 설정
    targetISO = _egYesterdaySydney();
    Logger.log('⚠️ 인자가 전달되지 않아 어제 날짜(' + targetISO + ')로 자동 설정됩니다.');
  }
  _egResetTACache();  // 캐시 무효화
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Daily_Report');
  if(!sheet){ Logger.log('Daily_Report 시트 없음'); return; }
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(String);
  Logger.log('=== EG Debug for ' + targetISO + ' ===');
  Logger.log('헤더: ' + headers.join(', '));

  // 차량 캐시 확인
  const owners = _egLoadVehicleOwners();
  Logger.log('M_Vehicles에서 EG 차량들:');
  Object.keys(owners).forEach(rego => {
    if(/eg\s*travel/i.test(owners[rego])){
      Logger.log('  ' + rego + ' → ' + owners[rego]);
    }
  });

  // 대상 일자의 모든 행 검사
  let matched = 0, unmatched = 0;
  let foundOnDate = 0;
  for(let i=1; i<data.length; i++){
    const row = {};
    headers.forEach((h, ci) => { row[h] = data[i][ci]; });
    const iso = _egToISO(row.Date);
    if(iso !== targetISO) continue;
    foundOnDate++;

    const rego = String(row.Rego||'').trim();
    const owner = owners[rego] || '(매핑 없음)';
    const matches = _egRowMatches(row);
    const summary = [
      'Row ' + (i+1),
      'Date=' + iso,
      'Rego=' + rego,
      'Driver=' + (row.Driver||''),
      'Agency=' + (row.Agency||''),
      'Billing=' + (row.Billing_Entity||row.BillingEntity||''),
      'VehOwner=' + owner,
      'Match=' + (matches ? 'YES ✅' : 'NO ❌')
    ].join(' · ');
    Logger.log(summary);
    if(matches) matched++; else unmatched++;
  }
  Logger.log('=== 결과: ' + targetISO + '에 ' + foundOnDate + '건 발견 / 매칭 ' + matched + '건 / 비매칭 ' + unmatched + '건 ===');
  if(foundOnDate === 0){
    Logger.log('⚠️ 해당 날짜에 Daily_Report 행이 없습니다. 날짜 형식 또는 저장 여부 확인 필요.');
    // 가까운 날짜들 샘플 출력
    Logger.log('--- 최근 10일 Daily_Report 날짜 분포 ---');
    const dateCount = {};
    for(let i=Math.max(1, data.length-100); i<data.length; i++){
      const row = {};
      headers.forEach((h, ci) => { row[h] = data[i][ci]; });
      const iso = _egToISO(row.Date);
      if(iso) dateCount[iso] = (dateCount[iso]||0) + 1;
    }
    Object.keys(dateCount).sort().slice(-10).forEach(d => {
      Logger.log('  ' + d + ': ' + dateCount[d] + '건');
    });
  }
  return { date: targetISO, found: foundOnDate, matched: matched, unmatched: unmatched };
}

// ═══════════════════════════════════════════════════════════════════════════
// 🚀 시트 read 캐싱 인프라 (2026-05-23)
// ═══════════════════════════════════════════════════════════════════════════
// 목적: 같은 시트를 1분 안에 여러 번 읽으면 시트 IO 안 거치고 캐시 응답
//      → 관리자/드라이버 앱 페이지 전환 속도 개선
//
// 동작 방식:
//   1. _cachedRead(sheetName, computeFn): 캐시에 있으면 즉시 반환,
//      없으면 computeFn() 실행 후 60초 캐싱
//   2. _invalidateSheetCache(sheetName): write 후 호출. 해당 시트 캐시 삭제
//   3. CacheService 값 한계(100KB) 우회 — 메타 키에 chunk 개수 저장,
//      chunk_0, chunk_1, ... 로 분할 저장
//
// 안전 가드:
//   - TTL 60초 (최악의 경우 1분 지연)
//   - 모든 save_*/update_*/delete_* 액션 후 해당 시트 캐시 자동 무효화
//   - 클라이언트가 ?force_refresh=1 또는 ?nocache=1 주면 캐시 무시
//   - 100KB 초과 시 분할 저장. 그래도 6MB(100KB × 60 chunk) 한계는 있음
//     → 1MB 이상 시트는 그냥 캐시 안 함 (시트 IO가 캐시 IO와 비슷해짐)
// ═══════════════════════════════════════════════════════════════════════════

const _SHEET_CACHE_TTL = 60;              // 60초
const _SHEET_CACHE_MAX_CHUNKS = 60;       // 최대 60 chunks (~6MB)
const _SHEET_CACHE_MAX_TOTAL_KB = 1024;   // 1MB 이상은 캐시 안 함
const _SHEET_CACHE_CHUNK_SIZE = 95 * 1024; // 95KB (100KB 한계 안전 마진)

// 캐시 활성 시트 화이트리스트 — 자주 읽지만 변경 빈도 낮은 것만
const _CACHE_ENABLED_SHEETS = {
  'Daily_Report': 1, 'Pre_Departure': 1, 'End_of_Shift': 1,
  'Invoices': 1, 'Agency_Txn': 1, 'SUB_Txn': 1, 'Schedule': 1,
  'Wages': 1, 'Ledger': 1, 'Driver_Roster': 1,
  'M_Vehicles': 1, 'M_Drivers': 1, 'M_Clients': 1, 'M_Guides': 1,
  'M_Hotels': 1, 'M_Trailers': 1, 'M_PriceClient': 1, 'M_PriceDriver': 1,
  'M_PriceSub': 1, 'M_SUB': 1, 'M_NightRates': 1, 'M_SvcOptions': 1,
  'M_HotelOptions': 1, 'M_DistOptions': 1, 'M_Attractions': 1,
  'Notices': 1, 'Company_Profile': 1, 'Leave_Requests': 1,
  'Defect_Reports': 1, 'MOT_Report': 1, 'Bus_Damage': 1,
  'HVIS_Bookings': 1, 'Maint_Records': 1,
  // 가상 키 (다중 시트 응답)
  'all_masters': 1
};

function _sheetCacheKey(sheetName) { return 'shc:' + sheetName; }

/**
 * 캐시된 read. 캐시 hit이면 즉시 반환, miss면 computeFn 실행 후 캐싱.
 * @param {string} sheetName  시트 이름 (또는 가상 키)
 * @param {function} computeFn  () => result. 캐시 miss 시에만 실행
 * @returns {*}  computeFn의 결과 (캐시 hit이면 deserialize한 동일 값)
 */
function _cachedRead(sheetName, computeFn) {
  if (!_CACHE_ENABLED_SHEETS[sheetName]) {
    return computeFn();
  }
  try {
    const cache = CacheService.getScriptCache();
    const metaKey = _sheetCacheKey(sheetName);
    const meta = cache.get(metaKey);
    if (meta) {
      // 캐시 hit — chunks 모아서 reconstruct
      try {
        const m = JSON.parse(meta);
        if (m.chunks === 1) {
          // 단일 chunk
          return JSON.parse(m.data);
        } else {
          // 다중 chunk
          const keys = [];
          for (let i = 0; i < m.chunks; i++) keys.push(metaKey + ':c' + i);
          const chunks = cache.getAll(keys);
          let combined = '';
          for (let i = 0; i < m.chunks; i++) {
            const part = chunks[metaKey + ':c' + i];
            if (!part) { combined = null; break; }
            combined += part;
          }
          if (combined !== null) {
            return JSON.parse(combined);
          }
          // chunk 일부 만료된 경우 fall-through (다시 계산)
        }
      } catch(e) {
        Logger.log('[cache] hit deserialize fail ' + sheetName + ': ' + e);
      }
    }
    // 캐시 miss — 계산 후 저장
    const result = computeFn();
    try {
      const serialized = JSON.stringify(result);
      const sizeKb = Math.ceil(serialized.length / 1024);
      if (sizeKb > _SHEET_CACHE_MAX_TOTAL_KB) {
        // 너무 크면 캐시 안 함 (다음 요청도 시트 IO하지만 메모리 절약)
        Logger.log('[cache] skip large sheet ' + sheetName + ' (' + sizeKb + 'KB)');
        return result;
      }
      if (serialized.length <= _SHEET_CACHE_CHUNK_SIZE) {
        // 단일 chunk
        cache.putAll({
          [metaKey]: JSON.stringify({ chunks: 1, data: serialized, savedAt: Date.now() })
        }, _SHEET_CACHE_TTL);
      } else {
        // 분할 저장
        const numChunks = Math.ceil(serialized.length / _SHEET_CACHE_CHUNK_SIZE);
        if (numChunks > _SHEET_CACHE_MAX_CHUNKS) {
          Logger.log('[cache] too many chunks for ' + sheetName);
          return result;
        }
        const toStore = {};
        for (let i = 0; i < numChunks; i++) {
          const start = i * _SHEET_CACHE_CHUNK_SIZE;
          toStore[metaKey + ':c' + i] = serialized.slice(start, start + _SHEET_CACHE_CHUNK_SIZE);
        }
        toStore[metaKey] = JSON.stringify({ chunks: numChunks, savedAt: Date.now() });
        cache.putAll(toStore, _SHEET_CACHE_TTL);
      }
    } catch(e) {
      Logger.log('[cache] put fail ' + sheetName + ': ' + e);
    }
    return result;
  } catch(e) {
    // 캐시 자체 실패 시 fallback — computeFn 직접 실행
    Logger.log('[cache] fatal ' + sheetName + ': ' + e);
    return computeFn();
  }
}

/**
 * 시트 변경 후 호출 — 해당 시트와 관련 캐시를 무효화.
 * @param {string|string[]} sheetName  시트 이름 또는 배열
 */
function _invalidateSheetCache(sheetName) {
  if (!sheetName) return;
  const names = Array.isArray(sheetName) ? sheetName : [sheetName];
  try {
    const cache = CacheService.getScriptCache();
    const keysToRemove = [];
    names.forEach(n => {
      const meta = cache.get(_sheetCacheKey(n));
      if (meta) {
        try {
          const m = JSON.parse(meta);
          if (m.chunks && m.chunks > 1) {
            for (let i = 0; i < m.chunks; i++) keysToRemove.push(_sheetCacheKey(n) + ':c' + i);
          }
        } catch(e) {}
        keysToRemove.push(_sheetCacheKey(n));
      }
      // 마스터 변경은 all_masters 종합 캐시도 무효화
      if (n && n.indexOf('M_') === 0) keysToRemove.push(_sheetCacheKey('all_masters'));
    });
    if (keysToRemove.length) {
      cache.removeAll(keysToRemove);
    }
  } catch(e) {
    Logger.log('[cache] invalidate fail: ' + e);
  }
}

/**
 * 전체 캐시 강제 무효화 (관리자 디버그 / 수동 동기화용)
 */
function _flushAllSheetCache() {
  try {
    const cache = CacheService.getScriptCache();
    const allKeys = Object.keys(_CACHE_ENABLED_SHEETS).map(_sheetCacheKey);
    // chunks 키들은 정확히 알 수 없으므로 메타와 c0~c{MAX}까지 일괄 삭제
    const expanded = [];
    allKeys.forEach(k => {
      expanded.push(k);
      for (let i = 0; i < _SHEET_CACHE_MAX_CHUNKS; i++) expanded.push(k + ':c' + i);
    });
    // removeAll은 한 번에 1000개까지 가능
    const CHUNK = 500;
    for (let i = 0; i < expanded.length; i += CHUNK) {
      cache.removeAll(expanded.slice(i, i + CHUNK));
    }
    Logger.log('[cache] flushed ' + expanded.length + ' keys');
    return { ok: true, keys: expanded.length };
  } catch(e) {
    Logger.log('[cache] flush fail: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// 🩺 EG Daily Report 발송 누락 진단 (2026-05-23)
// ═══════════════════════════════════════════════════════════════════════════
// 사용법: GAS Editor에서 diagEGReport() 실행 → Logger에서 결과 확인.
//
// 점검 항목:
//   1) 현재 등록된 트리거 (sendEGDailyReport / sendEGWeeklyReport)
//   2) EG_Report_Log 최근 발송 이력 (성공/실패/스킵)
//   3) 수신자 설정 (M_Clients의 EG TRAVEL 행 + Email 필드)
//   4) 전날 Daily_Report 데이터 (DR이 있는지)
//   5) 알려진 종료 투어코드 vs 이미 발송된 투어코드
//   6) Dry run으로 실제 발송 시뮬레이션 (이메일은 안 보냄)
// ═══════════════════════════════════════════════════════════════════════════

function diagEGReport() {
  const log = [];
  log.push('═══ EG Daily Report 진단 — ' + Utilities.formatDate(new Date(), 'Australia/Sydney', 'yyyy-MM-dd HH:mm:ss') + ' ═══');

  // ── 1) 트리거 상태 ──
  log.push('\n──[1] 등록된 트리거 ──');
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const egTriggers = triggers.filter(t => {
      const fn = t.getHandlerFunction();
      return fn === 'sendEGDailyReport' || fn === 'sendEGWeeklyReport';
    });
    if (egTriggers.length === 0) {
      log.push('  ❌ EG 리포트 트리거가 등록되어 있지 않음!');
      log.push('     → setupEGReportTriggers() 함수를 실행하세요');
    } else {
      egTriggers.forEach(t => {
        log.push('  ✅ ' + t.getHandlerFunction() +
                 ' / type=' + t.getEventType() +
                 ' / source=' + t.getTriggerSource());
      });
    }
  } catch(e) {
    log.push('  ⚠️ 트리거 조회 실패: ' + e);
  }

  // ── 2) 최근 발송 이력 ──
  log.push('\n──[2] EG_Report_Log 최근 10건 ──');
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const logSheet = ss.getSheetByName('EG_Report_Log');
    if (!logSheet) {
      log.push('  ⚠️ EG_Report_Log 시트가 없음 (아직 한 번도 실행 안 됨?)');
    } else {
      const data = logSheet.getDataRange().getValues();
      if (data.length < 2) {
        log.push('  ⚠️ 발송 이력 없음 (시트는 있지만 비어있음)');
      } else {
        const headers = data[0];
        const recentRows = data.slice(Math.max(1, data.length - 10));
        log.push('  헤더: ' + headers.join(' | '));
        recentRows.forEach((row, i) => {
          const summary = headers.map((h, ci) => {
            let v = row[ci];
            if (v instanceof Date) v = Utilities.formatDate(v, 'Australia/Sydney', 'yyyy-MM-dd HH:mm:ss');
            const s = String(v||'');
            return s.length > 40 ? s.substring(0,40)+'...' : s;
          });
          log.push('  • ' + summary.join(' | '));
        });
      }
    }
  } catch(e) {
    log.push('  ⚠️ EG_Report_Log 조회 실패: ' + e);
  }

  // ── 3) 수신자 설정 확인 ──
  log.push('\n──[3] 수신자 설정 ──');
  try {
    const recipients = _egGetRecipients();
    log.push('  TO:  "' + recipients.to + '"');
    log.push('  CC:  "' + recipients.cc + '"');
    log.push('  BCC: "' + recipients.bcc + '"');
    if (!recipients.to) {
      log.push('  ❌ TO 수신자가 비어있음! 발송 자체가 실패합니다.');
      log.push('     → M_Clients 시트에서 Name 컬럼에 "EG TRAVEL" 또는 "EG"가 포함된 행의 Email 필드를 확인하세요');
      log.push('     키워드: ' + (typeof EG_REPORT_KEYWORD !== 'undefined' ? EG_REPORT_KEYWORD : '(상수 미정의)'));
    }
  } catch(e) {
    log.push('  ⚠️ 수신자 조회 실패: ' + e);
  }

  // ── 4) 전날 DR 데이터 ──
  log.push('\n──[4] 전날 Daily_Report 데이터 ──');
  try {
    const yesterday = _egYesterdaySydney();
    log.push('  대상 날짜: ' + yesterday);
    const drs = _egLoadDRs(yesterday, yesterday);
    log.push('  DR 건수: ' + drs.length);
    if (drs.length === 0) {
      log.push('  ⚠️ 전날 DR이 0건 → 발송이 "no_dr" 사유로 스킵됨');
      log.push('     (이건 버그가 아니라 정상 동작 — 운행이 없으면 발송 안 함)');
    } else {
      log.push('  DR 샘플 (처음 5건):');
      drs.slice(0, 5).forEach((dr, i) => {
        const tc = dr.Tour_Code || dr.TourCode || '';
        const drv = dr.Driver || '';
        const ag = dr.Agency || '';
        log.push('    ' + (i+1) + '. TC=' + tc + ' / Driver=' + drv + ' / Agency=' + ag);
      });
    }
  } catch(e) {
    log.push('  ⚠️ DR 조회 실패: ' + e + '\n' + (e.stack || ''));
  }

  // ── 5) 종료된 투어코드 vs 이미 발송된 ──
  log.push('\n──[5] 새로 종료된 투어코드 ──');
  try {
    const todayISO = _egTodaySydney();
    const allCompleted = _egFindCompletedTourCodes(todayISO);
    const alreadySent = _egGetAlreadySentTourCodes();
    log.push('  종료 감지된 TC: ' + allCompleted.length);
    log.push('  이미 발송 처리된 TC: ' + alreadySent.size);
    const newCompleted = allCompleted.filter(t => !alreadySent.has(t.tourCode.toUpperCase()));
    log.push('  신규 종료 TC (이번 발송 대상): ' + newCompleted.length);
    if (newCompleted.length > 0) {
      log.push('  샘플:');
      newCompleted.slice(0, 5).forEach(t => {
        log.push('    • ' + t.tourCode + ' (마지막 운행일 ' + t.lastDate + ')');
      });
    }
  } catch(e) {
    log.push('  ⚠️ 종료 투어코드 조회 실패: ' + e);
  }

  // ── 6) Dry run ──
  log.push('\n──[6] Dry Run (실제 발송 안 함, 시뮬레이션만) ──');
  try {
    const dry = sendEGDailyReport({ dryRun: true });
    log.push('  결과: ' + JSON.stringify({
      ok: dry.ok, dryRun: dry.dryRun, skipped: dry.skipped,
      reason: dry.reason, drCount: dry.drCount, completedCount: dry.completedCount,
      error: dry.error
    }, null, 2));
    if (dry.skipped) {
      log.push('  → 발송 스킵 사유: ' + dry.reason);
    }
    if (dry.subject) log.push('  제목: ' + dry.subject);
  } catch(e) {
    log.push('  ⚠️ Dry run 실패: ' + e + '\n' + (e.stack || ''));
  }

  // ── 종합 진단 ──
  log.push('\n═══ 종합 권장 사항 ═══');
  log.push('  - 트리거가 없으면: setupEGReportTriggers() 실행');
  log.push('  - 수신자가 비어있으면: M_Clients의 EG TRAVEL 행에 Email 등록');
  log.push('  - 전날 DR이 없는데 발송 필요하면: 정상 동작이므로 변경 불필요');
  log.push('  - 트리거는 있는데 실행이 안 됐으면: GAS Editor → Triggers (좌측 시계 아이콘) → Execution History 확인');
  log.push('  - 즉시 강제 발송하려면: sendEGDailyReport({dryRun: false}) 수동 실행');

  const output = log.join('\n');
  Logger.log(output);
  return output;
}

// 어제 날짜로 강제 발송 (수동 보내기)
function sendEGDailyReport_force() {
  return sendEGDailyReport({ dryRun: false });
}

// 특정 날짜로 강제 발송 (예: sendEGDailyReport_forDate('2026-05-22'))
function sendEGDailyReport_forDate(dateISO) {
  return sendEGDailyReport({ dryRun: false, date: dateISO });
}
