/* ============================================================================
 * DC Fleet — Partner / BillingEntity Mode
 * ============================================================================
 * 단일 컴패니언 스크립트. admin.html / driver.html 에 <script src> 한 줄만 추가.
 *
 * 기능:
 *   1. 일정 모달에 BillingEntity 탭 (admin only)
 *   2. 일정 리스트 BillingEntity 배지/필터 (admin only)
 *   3. 드라이버 앱 — 일정 cache에서 billingEntity를 silent하게 추출 → 제출 시 전송
 *   4. Partner 모드 자동 감지 (sessionStorage.dc_role==='partner') → UI 잠금
 *
 * 작성: 2026-05
 * 라이센스: 내부용
 * ========================================================================== */

(function(global){
  'use strict';

  // ─── 상수 ─────────────────────────────────────────────────────────────────
  let PARTNER_COMPANIES = [
    { id: 'DC',                       label: '🇰🇷 DC (자사)',          color: '#1e40af', bg: '#dbeafe', border: '#3b82f6' },
    { id: 'EG TRAVEL PTY LTD',        label: '🚌 EG TRAVEL (파트너)',  color: '#6d28d9', bg: '#ede9fe', border: '#7c3aed' }
    // 신규 파트너 추가 시 이 배열에만 행 추가 (M_PriceSub에서 자동 하이드레이션됨)
  ];

  const isAdminPage  = !!document.getElementById('sch-modal')
                     || !!document.getElementById('sm-tourcode')
                     || /admin/.test(location.pathname);
  const isDriverPage = !!document.getElementById('page-report')
                     || !!document.getElementById('dr-date');

  // ─── 공통 헬퍼 ────────────────────────────────────────────────────────────
  function _getCfg(beId){
    const norm = String(beId||'DC').trim().toUpperCase();
    if(norm === 'DC' || norm === 'DONG CHOI PTY LTD') return PARTNER_COMPANIES[0];
    return PARTNER_COMPANIES.find(p => p.id.toUpperCase() === norm) || PARTNER_COMPANIES[0];
  }

  function _isDC(be){
    const n = String(be||'').trim().toUpperCase();
    return n === 'DC' || n === 'DONG CHOI PTY LTD' || n === '';
  }

  // 안전 적용: 함수가 정의되어 있을 때만 monkey-patch
  function _wrap(name, wrapper){
    if(typeof global[name] !== 'function') return false;
    const original = global[name];
    global[name] = function(){ return wrapper.call(this, original, arguments); };
    return true;
  }

  // ═════════════════════════════════════════════════════════════════════════
  // ADMIN.HTML 통합
  // ═════════════════════════════════════════════════════════════════════════

  function initAdmin(){
    if(!isAdminPage) return;

    // 스타일 주입
    const style = document.createElement('style');
    style.textContent = `
      .be-tab-row{display:flex;gap:8px;margin-bottom:12px;}
      .be-tab{flex:1;padding:10px;border-radius:8px;border:2px solid #e5e7eb;
              background:#f9fafb;color:#6b7280;font-weight:700;font-size:13px;
              cursor:pointer;opacity:.5;transition:all .15s;}
      .be-tab.on{opacity:1;}
      .be-info{margin-top:6px;font-size:10px;line-height:1.5;}
      .be-badge{padding:2px 6px;border-radius:4px;font-size:9px;font-weight:700;margin-left:4px;}
      .sch-flt-be{display:inline-flex;gap:4px;margin-left:8px;}
      [data-partner-mode] [data-admin-only]{display:none !important;}
      .partner-banner{background:#ede9fe;border-bottom:2px solid #7c3aed;padding:8px 14px;
                       font-size:12px;color:#6d28d9;font-weight:700;text-align:center;}
    `;
    document.head.appendChild(style);

    _injectBillingEntityTab();
    _hookSaveSchedule();
    _hookRenderScheduleList();
    _injectScheduleFilters();
    _detectPartnerMode();
    _hookBalDetailEGGroup();
  }

  // ─── 일정 모달 — BillingEntity 탭 주입 ─────────────────────────────────
  function _injectBillingEntityTab(){
    // 일정 모달이 열릴 때마다 탭 삽입 (멱등)
    const orig = global.openScheduleModal;
    if(typeof orig !== 'function') return;
    global.openScheduleModal = function(tourId){
      const ret = orig.apply(this, arguments);
      setTimeout(() => _ensureBeTabInModal(tourId), 50);
      return ret;
    };
  }

  function _ensureBeTabInModal(tourId){
    const modal = document.getElementById('sch-modal');
    if(!modal) return;
    // 신규 모달 — stale 값 방지
    if(tourId === undefined){ global._schEditBillingEntity = 'DC'; }
    // 동적 하이드레이션: M_PriceSub 에서 서브 회사 추가
    _hydratePartnersFromCache();
    if(modal.querySelector('.be-select-wrap')) {
      // 이미 있음 — 값만 갱신
      _refreshBeTab(tourId);
      return;
    }

    // 삽입 위치: TourCode 입력 위쪽
    const tcInput = document.getElementById('sm-tourcode');
    const insertBefore = tcInput?.closest('div[style*="margin"]') || tcInput?.parentElement;
    if(!insertBefore) return;

    const wrap = document.createElement('div');
    wrap.className = 'be-select-wrap';
    wrap.style.cssText = 'margin-bottom:12px;';
    wrap.innerHTML = `
      <div style="font-size:11px;font-weight:700;color:var(--t3,#6b7280);margin-bottom:6px;">
        💰 인보이스 발행사 (Billing Entity)
      </div>
      <select class="be-select" style="width:100%;padding:10px 12px;border-radius:8px;border:2px solid #e5e7eb;background:#f9fafb;font-weight:700;font-size:13px;cursor:pointer;transition:all .15s;">
        ${PARTNER_COMPANIES.map(p => `
          <option value="${p.id}">${p.label}</option>
        `).join('')}
      </select>
      <div class="be-info" id="sm-be-info" style="margin-top:6px;font-size:10px;line-height:1.5;"></div>
    `;
    insertBefore.parentElement.insertBefore(wrap, insertBefore);

    // 변경 핸들러
    const sel = wrap.querySelector('.be-select');
    sel.addEventListener('change', () => {
      global._schEditBillingEntity = sel.value;
      _refreshBeTab();
    });

    _refreshBeTab(tourId);
  }
  
  // M_PriceSub 캐시에서 서브 회사를 PARTNER_COMPANIES 에 추가 (중복 제외)
  function _hydratePartnersFromCache(){
    try {
      const sources = [
        global._priceSubCache,
        global._subCompanyCache,
        global._mPriceSubCache,
        (global.DB && (global.DB.PRICE_SUB || global.DB.M_PriceSub))
      ].filter(Array.isArray);
      if(!sources.length) return;
      const known = new Set(PARTNER_COMPANIES.map(p => String(p.id||'').trim().toUpperCase()));
      const seen = new Set();
      sources.forEach(arr => arr.forEach(r => {
        const company = String(r.Company || r.SubCompany || r.Sub || r.PartnerCompany || '').trim();
        if(!company) return;
        const key = company.toUpperCase();
        if(seen.has(key)) return;
        seen.add(key);
        if(known.has(key)) return;
        PARTNER_COMPANIES.push({
          id: company,
          label: '🚚 ' + company,
          color: '#475569',
          bg: '#f1f5f9',
          border: '#64748b'
        });
        known.add(key);
      }));
    } catch(e){ console.warn('[partner] hydrate failed', e); }
  }

  function _refreshBeTab(tourId){
    const modal = document.getElementById('sch-modal');
    if(!modal) return;
    // 디폴트 결정
    let current = global._schEditBillingEntity;
    if(tourId !== undefined){
      // 수정 모드 — DB에서 로드
      const cache = (global._schCache && global._schCache.length) ? global._schCache : ((global.DB && global.DB.SCH) ? global.DB.SCH : []);
      const sch = cache.find(s => String(s.TourID||'').trim() === String(tourId||'').trim()) ||
                  cache.find(s => String(s.TourCode||'').trim() === String(tourId||'').trim());
      if(sch && sch.BillingEntity){ current = String(sch.BillingEntity).trim(); global._schEditBillingEntity = current; }
      else if(!current){ current = 'DC'; global._schEditBillingEntity = current; }
    } else if(!current){
      current = 'DC';
      global._schEditBillingEntity = current;
    }
    // select 갱신
    const sel = modal.querySelector('.be-select');
    if(sel){
      // 옵션에 current 값이 없으면 추가 (legacy/unknown BillingEntity 대응)
      const has = Array.from(sel.options).some(o => o.value === current);
      if(!has && current){
        const opt = document.createElement('option');
        opt.value = current;
        opt.textContent = '❔ ' + current;
        sel.appendChild(opt);
      }
      sel.value = current || 'DC';
      const cfg = _getCfg(current);
      sel.style.borderColor = cfg.border;
      sel.style.background = cfg.bg;
      sel.style.color = cfg.color;
    }
    // 안내 문구
    const info = modal.querySelector('#sm-be-info');
    if(info){
      if(_isDC(current)){
        info.style.color = '#3b82f6';
        info.textContent = '✓ DC가 클라이언트에 인보이스를 발행합니다 (자사 운영)';
      } else {
        info.style.color = '#7c3aed';
        info.textContent = '⚠ ' + current + '가 클라이언트에 인보이스를 발행합니다 (DC는 서브 계약자)';
      }
    }
  }

  // ─── fetch hook — save_schedule 호출 시 BillingEntity 주입 ────────────
  // saveScheduleData는 apiPost 대신 fetch(APPS_SCRIPT_URL, ...)를 직접 호출하므로
  // fetch 자체를 인터셉트해서 body의 action==='save_schedule'인 경우 data에 BillingEntity 추가
  function _hookSaveSchedule(){
    if(global.__beFetchHooked) return;
    global.__beFetchHooked = true;
    const _origFetch = global.fetch;
    global.fetch = function(url, options){
      try {
        if(options && options.body && typeof options.body === 'string'){
          const body = JSON.parse(options.body);
          if(body && body.action === 'save_schedule' && body.data && typeof body.data === 'object'){
            body.data.BillingEntity = global._schEditBillingEntity || 'DC';
            options = Object.assign({}, options, { body: JSON.stringify(body) });
          }
        }
      } catch(e){ /* body가 JSON이 아니면 무시 */ }
      return _origFetch.call(this, url, options);
    };
  }

  // ─── 일정 리스트 — BillingEntity 배지 ─────────────────────────────────
  function _hookRenderScheduleList(){
    // 리스트 렌더링이 끝난 후 DOM에서 TourCode 셀에 배지 추가
    const observe = () => {
      document.querySelectorAll('[data-tour-id]:not([data-be-rendered])').forEach(row => {
        const tourId = row.dataset.tourId;
        const _cachePat = (global._schCache && global._schCache.length) ? global._schCache : ((global.DB && global.DB.SCH) ? global.DB.SCH : []);
        const tour = _cachePat.find(t => t.TourID === tourId) || null;
        if(!tour) return;
        const be = tour.BillingEntity || 'DC';
        if(_isDC(be)){
          row.dataset.beRendered = '1';
          return; // 디폴트 — 배지 생략
        }
        const cfg = _getCfg(be);
        const tcCell = row.querySelector('.tour-code, [data-field="TourCode"]')
                       || row.querySelector('td');
        if(tcCell && !tcCell.querySelector('.be-badge')){
          const badge = document.createElement('span');
          badge.className = 'be-badge';
          badge.style.cssText = `background:${cfg.bg};color:${cfg.color};border:1.5px solid ${cfg.border};padding:3px 8px;font-weight:800;`;
          badge.textContent = be === 'EG TRAVEL PTY LTD' ? '🤝 EG 발행 (자사 청구 X)' : be.split(' ')[0];
          tcCell.appendChild(badge);
        }
        // Also tint the entire row with EG color stripe (left border + bg)
        row.style.borderLeft = '4px solid ' + cfg.border;
        row.style.background = cfg.bg + '40';
        row.dataset.beRendered = '1';
      });
    };

    // 주기적으로 체크 (간단)
    setInterval(observe, 1000);
  }

  // ─── 일정 필터 칩 — 발행사 필터 ───────────────────────────────────────
  function _injectScheduleFilters(){
    const tryInject = () => {
      const dcChip = document.getElementById('sch-flt-dc');
      if(!dcChip || dcChip.dataset.beFiltered) return;

      // 필터 컨테이너에 BillingEntity 칩 추가
      const container = dcChip.parentElement;
      if(!container) return;

      const beWrap = document.createElement('div');
      beWrap.className = 'sch-flt-be';
      beWrap.innerHTML = `
        <button id="sch-flt-be-dc" class="sch-flt-chip on" data-be-filter="DC">🏢 자사 청구</button>
        <button id="sch-flt-be-eg" class="sch-flt-chip on" data-be-filter="EG TRAVEL PTY LTD">🤝 EG 청구</button>
      `;
      container.appendChild(beWrap);

      beWrap.querySelectorAll('button').forEach(btn => {
        btn.addEventListener('click', () => {
          btn.classList.toggle('on');
          if(typeof global.renderScheduleList === 'function') global.renderScheduleList();
          if(typeof global.loadSchedule === 'function') global.loadSchedule();
        });
      });

      dcChip.dataset.beFiltered = '1';
    };

    setInterval(tryInject, 800);

    // 기존 필터 함수에 BillingEntity 분기 추가 — schedule rows를 필터링
    _wrap('renderScheduleList', function(orig, args){
      const ret = orig.apply(this, args);
      setTimeout(() => {
        const dcOn = document.getElementById('sch-flt-be-dc')?.classList.contains('on');
        const egOn = document.getElementById('sch-flt-be-eg')?.classList.contains('on');
        document.querySelectorAll('[data-tour-id]').forEach(row => {
          const tourId = row.dataset.tourId;
          const _cacheP3 = (global._schCache && global._schCache.length) ? global._schCache : ((global.DB && global.DB.SCH) ? global.DB.SCH : []);
          const tour = _cacheP3.find(t => t.TourID === tourId) || null;
          if(!tour) return;
          const be = tour.BillingEntity || 'DC';
          const isDC = _isDC(be);
          const isEG = String(be).toUpperCase() === 'EG TRAVEL PTY LTD';
          let hide = false;
          if(isDC && !dcOn) hide = true;
          if(isEG && !egOn) hide = true;
          row.style.display = hide ? 'none' : '';
        });
      }, 50);
      return ret;
    });
  }

  // ─── Partner 모드 자동 감지 + UI 잠금 ─────────────────────────────────
  function _detectPartnerMode(){
    const role = (sessionStorage.getItem('dc_role') || localStorage.getItem('dc_role') || '').trim();
    const partnerCompany = (sessionStorage.getItem('dc_partner_company') ||
                            localStorage.getItem('dc_partner_company') || '').trim();
    if(role !== 'partner' || !partnerCompany) return;

    document.body.dataset.partnerMode = '1';

    // 상단 배너
    const banner = document.createElement('div');
    banner.className = 'partner-banner';
    banner.innerHTML = `🤝 ${partnerCompany} 파트너 모드 — DC 자료 접근 제한됨`;
    document.body.insertBefore(banner, document.body.firstChild);

    // BillingEntity 강제 (탭이 생성될 때마다)
    setInterval(() => {
      const modal = document.getElementById('sch-modal');
      if(!modal || modal.style.display === 'none') return;
      modal.querySelectorAll('.be-tab').forEach(btn => {
        if(btn.dataset.be !== partnerCompany){
          btn.disabled = true;
          btn.style.opacity = '0.2';
          btn.style.cursor = 'not-allowed';
          btn.title = 'Partner 모드 — 본인 회사 일정만 등록 가능';
        }
      });
      if(global._schEditBillingEntity !== partnerCompany){
        global._schEditBillingEntity = partnerCompany;
        _refreshBeTab();
      }
    }, 500);
  }

  // ═════════════════════════════════════════════════════════════════════════
  // DRIVER.HTML 통합 (silent — 드라이버는 발행사를 모름)
  // ═════════════════════════════════════════════════════════════════════════

  function initDriver(){
    if(!isDriverPage) return;

    // 일정 cache의 billingEntity를 prefill 시 글로벌에 저장
    _wrap('applySchedulePrefill', function(orig, args){
      const ret = orig.apply(this, args);
      try {
        const raw = sessionStorage.getItem('dc_dr_prefill_consumed')
                   || sessionStorage.getItem('dc_dr_prefill');
        if(raw){
          const data = JSON.parse(raw);
          global._activeBillingEntity = (data.billingEntity || data.BillingEntity || 'DC')
                                         .toString().trim();
        }
      } catch(e){}
      return ret;
    });

    // startDRFromSchedule 가 sessionStorage 에 prefill 저장하기 직전 —
    // billingEntity 같이 저장되도록 보장
    _wrap('startDRFromSchedule', function(orig, args){
      const [tourId, date, slotKey] = args;
      const cache = global._driverScheduleCache || [];
      const matched = cache.find(s =>
        s.tourId === tourId && s.date === date && s.slotKey === slotKey
      );
      if(matched){
        // billingEntity가 누락된 경우 — schedule cache에서 찾아 보강
        if(!matched.billingEntity && !matched.BillingEntity){
          // 백엔드가 새 컬럼 반영 전이면 비어있을 수 있음. 디폴트 DC.
          matched.billingEntity = 'DC';
        }
      }
      return orig.apply(this, args);
    });

    // submitDailyReport 가 만드는 data 객체에 Billing_Entity 컬럼 자동 추가
    _wrap('submitDailyReport', function(orig, args){
      // 원본 함수가 내부에서 fetch/apiPost 호출하기 전에 data를 가로채기 위한 hook
      const apiPostOrig = global.apiPost;
      let intercepted = false;

      global.apiPost = function(action, payload){
        if(!intercepted && (action === 'add_report' || action === 'save_report')){
          intercepted = true;
          try {
            const dataObj = (payload && payload.data)
              ? (typeof payload.data === 'string' ? JSON.parse(payload.data) : payload.data)
              : payload;
            if(dataObj && typeof dataObj === 'object'){
              dataObj.Billing_Entity = global._activeBillingEntity || 'DC';
              if(payload && payload.data){
                payload.data = typeof payload.data === 'string'
                              ? JSON.stringify(dataObj) : dataObj;
              }
            }
          } catch(e){ console.warn('[partner-mode] Billing_Entity injection failed:', e); }
        }
        const ret = apiPostOrig.apply(this, arguments);
        setTimeout(() => {
          global.apiPost = apiPostOrig;
          // 다음 DR이 새지 않도록
          global._activeBillingEntity = null;
        }, 200);
        return ret;
      };

      return orig.apply(this, args);
    });

    // fetch 직접 사용 시 대비 — 이미 apiPost를 안 쓰는 경로가 있으면 fetch도 가로채기
    const fetchOrig = global.fetch;
    global.fetch = function(url, opts){
      try {
        if(opts && opts.method === 'POST' && opts.body){
          let body = opts.body;
          if(typeof body === 'string'){
            try {
              const parsed = JSON.parse(body);
              if(parsed && parsed.action === 'save_report'
                 && parsed.sheet === 'Daily_Report'
                 && parsed.data && !parsed.data.Billing_Entity){
                parsed.data.Billing_Entity = global._activeBillingEntity || 'DC';
                opts = Object.assign({}, opts, { body: JSON.stringify(parsed) });
              }
            } catch(e){}
          }
        }
      } catch(e){}
      return fetchOrig.call(this, url, opts);
    };
  }

  // ═════════════════════════════════════════════════════════════════════════
  // 부트스트랩
  // ═════════════════════════════════════════════════════════════════════════

  function boot(){
    try {
      if(isAdminPage)  initAdmin();
      if(isDriverPage) initDriver();
      console.log('[partner-mode] initialized', { admin: isAdminPage, driver: isDriverPage });
    } catch(e){
      console.error('[partner-mode] boot error:', e);
    }
  }

  // DOMContentLoaded 후 부트
  if(document.readyState === 'loading'){
    document.addEventListener('DOMContentLoaded', boot);
  } else {
    boot();
  }

  // 외부 노출 (디버깅용)
  // ─── EG-billed grouping in bal-detail modal ───
  function _hookBalDetailEGGroup(){
    if(!isAdminPage) return;
    setInterval(() => {
      const modal = document.getElementById('bal-detail-modal');
      if(!modal) return;
      const visible = getComputedStyle(modal).display !== 'none';
      if(!visible){ modal.dataset.egProcessed = ''; return; }
      if(modal.dataset.egProcessed === '2') return;
      modal.dataset.egProcessed = '2';
      // Build tour code -> BillingEntity map
      const cache = _getScheduleCache();
      const tourBE = {};
      cache.forEach(t => {
        const tc = String(t.TourCode||'').trim();
        if(tc) tourBE[tc] = String(t.BillingEntity||'').trim().toUpperCase();
      });
      // Find tour code rows by text content matching
      let egTotal = 0;
      const allEls = modal.querySelectorAll('div, tr, li');
      allEls.forEach(el => {
        if(el.dataset.egMarked) return;
        const text = (el.textContent||'').trim();
        if(text.length > 500 || text.length < 5) return;
        for(const tc of Object.keys(tourBE)){
          if(text.indexOf(tc) >= 0 && tourBE[tc] === 'EG TRAVEL PTY LTD'){
            // Mark this row
            el.style.opacity = '0.55';
            el.style.background = 'rgba(237,233,254,0.4)';
            el.style.borderLeft = '3px solid #7c3aed';
            // Try to find amount in text
            const m = text.match(/\$([\d,]+(?:\.\d+)?)/);
            if(m) egTotal += parseFloat(m[1].replace(/,/g,''));
            // Add badge if not already there
            if(!el.querySelector('.eg-billed-badge')){
              const badge = document.createElement('span');
              badge.className = 'eg-billed-badge';
              badge.style.cssText = 'background:#ede9fe;color:#7c3aed;padding:2px 6px;border-radius:4px;font-size:10px;font-weight:800;margin-left:6px;display:inline-block;';
              badge.textContent = '🤝 EG 직접 (자사 지급 제외)';
              el.appendChild(badge);
            }
            el.dataset.egMarked = '1';
            break;
          }
        }
      });
      // Show subtotal at top if EG amount > 0
      if(egTotal > 0 && !modal.querySelector('.eg-subtotal-banner')){
        const banner = document.createElement('div');
        banner.className = 'eg-subtotal-banner';
        banner.style.cssText = 'background:#ede9fe;color:#5b21b6;padding:8px 12px;border-radius:6px;margin:8px 14px;font-size:12px;font-weight:700;border:1px solid #c4b5fd;';
        banner.innerHTML = '🤝 EG 자체 처리 명세 합계: 
    version: '1.0.0',
    PARTNER_COMPANIES: PARTNER_COMPANIES,
    refreshTab: () => isAdminPage && _refreshBeTab(),
    getCurrentBE: () => global._schEditBillingEntity,
    setBE: (id) => { global._schEditBillingEntity = id; _refreshBeTab(); }
  };

})(window);
 + egTotal.toLocaleString('en-AU') + ' (자사 지급 제외 — 참고용)';
        const firstChild = modal.querySelector('.bal-detail-content, [class*="detail"]') || modal.firstElementChild;
        if(firstChild) firstChild.insertBefore(banner, firstChild.firstChild);
      }
    }, 800);
  }

  global.DCPartnerMode = {
    version: '1.0.0',
    PARTNER_COMPANIES: PARTNER_COMPANIES,
    refreshTab: () => isAdminPage && _refreshBeTab(),
    getCurrentBE: () => global._schEditBillingEntity,
    setBE: (id) => { global._schEditBillingEntity = id; _refreshBeTab(); }
  };

})(window);
