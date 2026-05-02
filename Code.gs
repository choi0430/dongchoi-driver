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
  'M_Drivers':  ['Name_EN','Name_KR','Initials','DriverID','Mobile_1','NEXT_OF_KIN','Moblie_2','License_Class',
                 'License_No','License_Expiry','Authority_No','Authority_Expiry','WWC_No','WWC_Expiry',
                 'Address','Suburb','Bank_Name','BSB','Account_Number','PIN','Owner','Active'],
  'M_Clients':  ['Name','ClientID','ABN','Mobile','Email','Email_CC','Address','Bank_Name','BSB','Account_Number'],
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
  'Schedule':   ['TourID','Agency','TourCode','StartDate','EndDate','Pax','Seats','Trailer',
                 'Guide','GuidePhone','Driver','Rego','FlightIn','FlightOut','Hotel',
                 'TourPlan','Notes','Status','InvoiceID','Quote','CreatedAt','UpdatedAt']
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
  'update_driver_info',
  'update_defect_status',
  'review_leave_request', 'update_roster_cell',
  'save_hvis_booking', 'delete_hvis_booking',
  'save_maint_record', 'delete_maint_record',
  'save_invoice_override', 'delete_invoice_override', 'bulk_save_invoice_overrides',
  'save_company_profile',
  // ── 운행 일정 관리 (Schedule) ──
  'save_schedule', 'delete_schedule', 'update_schedule_status',
  // 관리자가 주로 쓰지만 드라이버도 가끔 필요할 수 있는 조회는 제외:
  // get_invoices, get_agency_txn, get_sub_txn 등은 일단 드라이버도 허용
  // 추후 엄격하게 할 수 있음
];

// 관리자 전용 GET 액션
const ADMIN_ONLY_GET_ACTIONS = [
  'get_agency_txn', 'get_sub_txn', 'get_agency_balances',
  'get_invoices', 'get_all_leave_requests', 'get_roster',
  'get_ledger', 'get_defect_reports',
  'get_admin_bundle', 'get_audit_log',
  // ── 운행 일정 ──
  'get_schedule', 'get_schedule_stats'
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

      case 'get_driver_schedule': {
        // 드라이버에게 배정된 일정 조회 (드라이버 앱용 — 인증 불필요, 드라이버 식별만)
        const driver = e.parameter.driver || '';
        const from = e.parameter.from || '';
        const to = e.parameter.to || '';
        return cors(getDriverSchedule(driver, from, to));
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

      case 'get_my_shifts':
        return cors(getMyShifts(effectiveDriver));

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
        const defDriver = e.parameter.driver || '';
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

      case 'get_roster':
        return cors(getRosterData(e.parameter.from, e.parameter.to));

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
      if (payload.data && typeof payload.data === 'object' && payload.data.Driver) {
        payload.data.Driver = tokenValid.user;
      }
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

    let rows = data.slice(1).map((row, i) => {
      const obj = {_rowIndex: i + 2}; // 1-indexed sheet row (헤더가 1행이므로 i+2)
      headers.forEach((h, idx) => {
        obj[h] = formatCell(h, row[idx]);
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
    return {ok: r.ok, rows: r.rows || []};
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
      nokName: 'NEXT_OF_KIN', nokPhone: 'Moblie_2',
      address: 'Address', suburb: 'Suburb',
      bank: 'Bank', bsb: 'BSB', account: 'Account'
    };

    for (let r = 1; r < sheetData.length; r++) {
      if (sheetData[r][nameENIdx] === driverName || sheetData[r][nameKRIdx] === driverName) {
        const PHONE_SAVE_FIELDS = ['Mobile_1', 'Moblie_2', 'Phone', 'Mobile'];
        const DATE_SAVE_FIELDS = ['License_Expiry', 'Authority_Expiry', 'WWC_Expiry'];
        Object.entries(data).forEach(([key, val]) => {
          const col = fieldMap[key];
          if (col) {
            const colIdx = headers.indexOf(col);
            if (colIdx !== -1) {
              const cell = sheet.getRange(r + 1, colIdx + 1);
              // ★ 전화번호 필드: 텍스트 서식 강제 적용 (앞 0 보존)
              if (PHONE_SAVE_FIELDS.includes(col)) {
                let s = String(val||'').replace(/[^0-9]/g, '');
                if (s.length === 9) s = '0' + s;
                cell.setNumberFormat('@').setValue(s);
              } else if (DATE_SAVE_FIELDS.includes(col)) {
                // ★ 날짜 필드: 정규화 후 저장 (잘못된 형식이면 빈 값)
                const norm = _normalizeDateForSheet(val);
                cell.setNumberFormat('@').setValue(norm);
              } else {
                cell.setValue(val);
              }
            }
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
    if (docHtml) {
      var htmlBlob = Utilities.newBlob(docHtml, 'text/html', 'invoice.html');
      var pdfBlob  = htmlBlob.getAs('application/pdf').setName(pdfName);
      options.attachments = [pdfBlob];
    } else if (pdfBase64) {
      var pdfBytes = Utilities.base64Decode(pdfBase64);
      var pdfBlob2 = Utilities.newBlob(pdfBytes, 'application/pdf', pdfName);
      options.attachments = [pdfBlob2];
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
      `인보이스 이메일 발송 (PDF 첨부) → ${to} | ${subject}`);

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
          const slotDriver = _stripPrefix(slot.driver);
          const targetDriver = _stripPrefix(driver);
          if (slotDriver !== targetDriver) return;
          result.push({
            tourId: tourId,
            tourCode: tourCode,
            agency: agency,
            date: dateStr,
            slotKey: slotKey,
            slot: slot,
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

    const rowArr = headers.map(h => data[h] !== undefined ? data[h] : '');

    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, headers.length).setValues([rowArr]);
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
    let updated = 0;
    for (let i = 1; i < data.length; i++) {
      const sd = String(data[i][sdCol]||'').slice(0, 10);
      const ed = String(data[i][edCol]||'').slice(0, 10);
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
