/* ============================================================================
 * DC Fleet — Driver BillingEntity Helper (v1)
 * Adds BE dropdown to Daily Report form + auto-detect from admin schedule.
 * Also adds "개인일정 추가" button for ad-hoc reports.
 * ============================================================================ */
(function(){
  'use strict';
  if (typeof window === "undefined") return;

  var PARTNER_DEFAULTS = [
    { id: 'DC',                  label: '\u{1F1F0}\u{1F1F7} DC (자사)',          color: '#1e40af', bg: '#dbeafe', border: '#3b82f6' },
    { id: 'EG TRAVEL PTY LTD',   label: '\u{1F68C} EG TRAVEL (파트너)',  color: '#6d28d9', bg: '#ede9fe', border: '#7c3aed' }
  ];

  var isDriverPage = !/admin\.html?/.test(location.pathname) && !/Admin/.test(document.title || '');
  if (!isDriverPage) return;

  // Bridge let-declared globals to window via injected <script>
  function bridgeGlobals() {
    var script = document.createElement("script");
    script.textContent = [
      "(function(){",
      "  try { if (typeof _subCompanies !== \"undefined\") window.__beDdSubCompanies = _subCompanies; } catch(e) {}",
      "  try { if (typeof _driverSchedule !== \"undefined\") window.__beDdDriverSched = _driverSchedule; } catch(e) {}",
      "  try { if (typeof _driverScheduleCache !== \"undefined\") window.__beDdDriverSchedCache = _driverScheduleCache; } catch(e) {}",
      "  try { if (typeof _schedule !== \"undefined\") window.__beDdSched = _schedule; } catch(e) {}",
      "  try { if (typeof _schedules !== \"undefined\") window.__beDdScheds = _schedules; } catch(e) {}",
      "  try { if (typeof _mySchedule !== \"undefined\") window.__beDdMySched = _mySchedule; } catch(e) {}",
      "  try { if (typeof _mySchedules !== \"undefined\") window.__beDdMyScheds = _mySchedules; } catch(e) {}",
      "  try { if (typeof _agencyList !== \"undefined\") window.__beDdAgencyList = _agencyList; } catch(e) {}",
      "  try { if (typeof _attractionList !== \"undefined\") window.__beDdAttractionList = _attractionList; } catch(e) {}",
      "})();"
    ].join("");
    (document.head || document.documentElement).appendChild(script);
    if (script.parentNode) script.parentNode.removeChild(script);
  }

  function injectStyle() {
    if (document.getElementById("be-driver-style")) return;
    var s = document.createElement("style");
    s.id = "be-driver-style";
    s.textContent = [
      ".be-dr-wrap { margin: 12px 0; padding: 12px; border-radius: 10px; background: #f9fafb; border: 1.5px solid #e5e7eb; }",
      ".be-dr-wrap.be-locked { background: #f3f4f6; }",
      ".be-dr-wrap.be-personal { background: #fef3c7; border-color: #f59e0b; }",
      ".be-dr-label { font-size: 11px; font-weight: 700; color: var(--t3, #6b7280); margin-bottom: 6px; display: block; }",
      ".be-dr-select { width: 100%; padding: 10px 12px; border-radius: 8px; border: 2px solid #e5e7eb; background: #ffffff; font-weight: 700; font-size: 14px; cursor: pointer; transition: all .15s; box-sizing: border-box; }",
      ".be-dr-select:disabled { background: #f3f4f6; cursor: not-allowed; opacity: 0.85; }",
      ".be-dr-info { margin-top: 6px; font-size: 11px; line-height: 1.5; }",
      ".be-dr-badge { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 10px; font-weight: 700; margin-left: 6px; }",
      ".be-dr-badge.locked { background: #dbeafe; color: #1e40af; }",
      ".be-dr-badge.personal { background: #fbbf24; color: #78350f; }",
      ".be-personal-btn { display: inline-block; margin-left: 8px; padding: 6px 12px; border-radius: 6px; background: #fbbf24; color: #78350f; border: 1px solid #f59e0b; font-weight: 700; font-size: 13px; cursor: pointer; }",
      ".be-personal-btn:hover { background: #f59e0b; color: #fff; }"
    ].join("\n");
    document.head.appendChild(s);
  }

  function tryName(item) {
    if (item == null) return "";
    if (typeof item === "string" || typeof item === "number") return String(item);
    if (typeof item === "object") {
      return item.SubName || item.subName || item.Sub || item.sub ||
             item.Company || item.company || item.SubCompany || item.subCompany ||
             item.PartnerCompany || item.partnerCompany || item.Name || item.name || "";
    }
    return "";
  }

  function getSchedules() {
    if (Array.isArray(window.__beDdDriverSchedCache) && window.__beDdDriverSchedCache.length) return window.__beDdDriverSchedCache;
    if (Array.isArray(window._driverScheduleCache) && window._driverScheduleCache.length) return window._driverScheduleCache;
    if (Array.isArray(window.__beDdDriverSched) && window.__beDdDriverSched.length) return window.__beDdDriverSched;
    if (Array.isArray(window.__beDdSched) && window.__beDdSched.length) return window.__beDdSched;
    if (Array.isArray(window.__beDdScheds) && window.__beDdScheds.length) return window.__beDdScheds;
    if (Array.isArray(window.__beDdMySched) && window.__beDdMySched.length) return window.__beDdMySched;
    if (Array.isArray(window.__beDdMyScheds) && window.__beDdMyScheds.length) return window.__beDdMyScheds;
    if (Array.isArray(window._driverSchedule) && window._driverSchedule.length) return window._driverSchedule;
    if (Array.isArray(window._schedule)) return window._schedule;
    if (Array.isArray(window._schedules)) return window._schedules;
    return [];
  }

  function getSubArr() {
    if (Array.isArray(window.__beDdSubCompanies) && window.__beDdSubCompanies.length) return window.__beDdSubCompanies;
    if (Array.isArray(window._subCompanies) && window._subCompanies.length) return window._subCompanies;
    return [];
  }

  function getKnownPartners() {
    var out = PARTNER_DEFAULTS.slice();
    var known = new Set(out.map(function(p){ return String(p.id||"").toUpperCase(); }));
    var add = function(name){
      if (!name) return;
      var s = String(name).trim(); if (!s) return;
      var key = s.toUpperCase(); if (known.has(key)) return;
      known.add(key);
      out.push({ id: s, label: "\u{1F69A} " + s, color: "#475569", bg: "#f1f5f9", border: "#64748b" });
    };
    getSubArr().forEach(function(item){ add(tryName(item)); });
    // Also extract from schedule data
    getSchedules().forEach(function(s){ if (s && s.BillingEntity) add(s.BillingEntity); });
    return out;
  }

  function isDC(be) { return !be || String(be).toUpperCase() === "DC" || String(be).toUpperCase() === "DONG CHOI" || String(be).toUpperCase() === "DONG CHOI PTY LTD"; }

  // TourCode가 Schedule에 있으면 그 BE 반환. BE가 비어있으면 'DC'로 처리 (자사 인보이스 발행 기본)
  // 반환값: { matched: bool, be: string }
  // 슬롯 객체는 두 가지 형식 모두 지원:
  //   1) 시트 raw 헤더 형식: { TourCode, TourID, BillingEntity }
  //   2) driver app schedule slot 형식: { tourCode, tourId, billingEntity, BillingEntity }
  function lookupBEByTourCode(tourCode) {
    if (!tourCode) return { matched: false, be: "" };
    var t = String(tourCode).trim().toUpperCase();
    var schs = getSchedules();
    for (var i = 0; i < schs.length; i++) {
      var s = schs[i];
      if (!s) continue;
      var tc1 = String(s.TourCode || s.tourCode || "").trim().toUpperCase();
      var tc2 = String(s.TourID   || s.tourId   || "").trim().toUpperCase();
      if (tc1 === t || tc2 === t) {
        var be = String(s.BillingEntity || s.billingEntity || "").trim();
        return { matched: true, be: be || "DC" };  // 매칭됐는데 BE 비어있으면 DC로
      }
    }
    return { matched: false, be: "" };
  }

  function getCfg(beId, partners) {
    var found = partners.find(function(p){ return String(p.id||"").toUpperCase() === String(beId||"").toUpperCase(); });
    return found || partners[0];
  }

  // State: is current report admin-scheduled (locked) or personal (editable)
  // Reset to "auto-detect" mode by default; "personal" mode triggered by button
  function getMode() {
    return window.__beDrPersonalMode ? "personal" : "auto";
  }
  function setMode(m) {
    window.__beDrPersonalMode = (m === "personal");
  }

  function buildDropdown() {
    var tcInput = document.getElementById("dr-tourcode");
    if (!tcInput) return false;
    // Find a stable insertion point — before the agency wrap
    var agencyWrap = document.getElementById("sdd-agency-wrap");
    var insertParent = null, insertBefore = null;
    if (agencyWrap) {
      // walk up to find Agency label parent
      var node = agencyWrap.parentElement;
      while (node && !/(label|div)/i.test(node.tagName)) node = node.parentElement;
      // The Agency field group is typically inside a label or div - insert before it
      var labelParent = agencyWrap.closest("label") || agencyWrap.parentElement;
      insertParent = labelParent.parentElement;
      insertBefore = labelParent;
    }
    if (!insertParent) return false;

    // Already exists?
    if (insertParent.querySelector(".be-dr-wrap")) return false;

    var wrap = document.createElement("div");
    wrap.className = "be-dr-wrap";

    var label = document.createElement("label");
    label.className = "be-dr-label";
    label.textContent = "\u{1F4B0} 인보이스 발행사 (Billing Entity)";
    wrap.appendChild(label);

    var select = document.createElement("select");
    select.className = "be-dr-select";
    select.id = "be-dr-select";
    wrap.appendChild(select);

    var info = document.createElement("div");
    info.className = "be-dr-info";
    info.id = "be-dr-info";
    wrap.appendChild(info);

    insertParent.insertBefore(wrap, insertBefore);

    refreshDropdown();

    // Listen for tour code changes
    tcInput.addEventListener("input", function(){ refreshDropdown(); });
    tcInput.addEventListener("change", function(){ refreshDropdown(); });

    select.addEventListener("change", function(){
      // ★ data-locked 속성이 있으면 변경 거부 (개발자 도구로 disabled 풀고 변경하는 경우 방어)
      //   refreshDropdown을 호출해서 일정의 BE로 강제 복원
      if (select.getAttribute('data-locked') === '1') {
        console.warn("[partner-driver-be] BE change blocked — locked by schedule. Forcing refresh.");
        refreshDropdown();
        return;
      }
      window._drBillingEntity = select.value;
      var p = getKnownPartners();
      applyStyle(wrap, select, info, p);
      console.log("[partner-driver-be] BE selected:", select.value);
    });

    return true;
  }

  function applyStyle(wrap, select, info, partners) {
    var cfg = getCfg(select.value, partners);
    select.style.borderColor = cfg.border;
    select.style.color = cfg.color;
    var locked = !select.disabled === false;
    if (isDC(select.value)) {
      info.style.color = "#3b82f6";
      info.innerHTML = "\u2713 DC가 클라이언트에 인보이스를 발행합니다 (자사 운영)" +
        (select.disabled ? ' <span class="be-dr-badge locked">\u{1F512} 자동 (어드민 일정)</span>' : ' <span class="be-dr-badge personal">\u{270F} 개인일정 모드</span>');
    } else {
      info.style.color = "#7c3aed";
      info.innerHTML = "\u26A0 " + select.value + "가 클라이언트에 인보이스를 발행합니다" +
        (select.disabled ? ' <span class="be-dr-badge locked">\u{1F512} 자동 (어드민 일정)</span>' : ' <span class="be-dr-badge personal">\u{270F} 개인일정 모드</span>');
    }
    if (select.disabled) {
      wrap.classList.add("be-locked");
      wrap.classList.remove("be-personal");
    } else {
      wrap.classList.remove("be-locked");
      wrap.classList.add("be-personal");
    }
  }

  function refreshDropdown() {
    bridgeGlobals();
    var select = document.getElementById("be-dr-select");
    var info = document.getElementById("be-dr-info");
    if (!select) return;
    var wrap = select.closest(".be-dr-wrap");
    if (!wrap) return;

    var tcInput = document.getElementById("dr-tourcode");
    var tourCode = tcInput ? String(tcInput.value || "").trim() : "";

    var partners = getKnownPartners();
    // Rebuild options
    while (select.firstChild) select.removeChild(select.firstChild);
    partners.forEach(function(p){
      var opt = document.createElement("option");
      opt.value = p.id; opt.textContent = p.label;
      select.appendChild(opt);
    });

    var mode = getMode();
    var lookup = lookupBEByTourCode(tourCode);
    var detected = lookup.be;
    var matchedInSchedule = lookup.matched;

    if (matchedInSchedule) {
      // ★ Admin schedule에 매칭됨 — 드라이버는 변경 불가 (BE가 비었으면 'DC'로 표시)
      //   "개인일정 추가" 버튼을 클릭해도 무시: 일정에 있으면 일정이 source of truth
      if (!partners.find(function(p){ return p.id === detected; })) {
        var opt = document.createElement("option");
        opt.value = detected; opt.textContent = "\u2754 " + detected;
        select.appendChild(opt);
      }
      select.value = detected;
      select.disabled = true;
      // ★ disabled 외 readonly 속성 + pointer-events 차단 (개발자 도구 우회 한 단계 방어)
      select.setAttribute('data-locked', '1');
      window._drBillingEntity = detected;
    } else if (mode === "personal") {
      // No admin schedule match + personal mode — unlocked
      select.disabled = false;
      select.removeAttribute('data-locked');
      var v = window._drBillingEntity || "DC";
      if (!partners.find(function(p){ return p.id === v; })) {
        var opt2 = document.createElement("option");
        opt2.value = v; opt2.textContent = "\u2754 " + v;
        select.appendChild(opt2);
      }
      select.value = v;
    } else {
      // Auto mode, no admin match — unlocked default
      select.disabled = false;
      select.removeAttribute('data-locked');
      select.value = window._drBillingEntity || "DC";
      window._drBillingEntity = select.value;
    }

    applyStyle(wrap, select, info, partners);
    console.log("[partner-driver-be] refresh: tourCode=" + tourCode + ", mode=" + mode + ", matched=" + matchedInSchedule + ", BE=" + select.value);
  }

  function injectPersonalButton() {
    if (document.getElementById("be-personal-btn")) return;
    // Find "리포트 추가" link/button
    var addLink = null;
    var allLinks = document.querySelectorAll("a, button");
    for (var i = 0; i < allLinks.length; i++) {
      var t = (allLinks[i].textContent || "").trim();
      if (/리포트 추가|Add Daily Report/i.test(t)) { addLink = allLinks[i]; break; }
    }
    if (!addLink) return;

    var btn = document.createElement("button");
    btn.id = "be-personal-btn";
    btn.className = "be-personal-btn";
    btn.type = "button";
    btn.innerHTML = "\u{1F195} 개인일정 추가";
    btn.title = "어드민 일정에 없는 ad-hoc 작업 보고";
    btn.addEventListener("click", function(){
      setMode("personal");
      // Trigger the existing "Add Daily Report" link
      addLink.click();
      // Refresh dropdown to unlocked mode
      setTimeout(refreshDropdown, 200);
      console.log("[partner-driver-be] personal mode activated");
    });
    addLink.parentElement.insertBefore(btn, addLink.nextSibling);
  }

  function injectFetchHook() {
    if (window.__beDrFetchHooked) return;
    window.__beDrFetchHooked = true;
    var origFetch = window.fetch.bind(window);
    window.fetch = function(url, options) {
      var injected = null;
      try {
        if (options && options.body && typeof options.body === "string") {
          var body = JSON.parse(options.body);
          if (body && body.action === "save_report" && body.data && typeof body.data === "object") {
            var be = window._drBillingEntity;
            if (be) {
              body.data.Billing_Entity = be;
              injected = be;
              options = Object.assign({}, options, { body: JSON.stringify(body) });
              console.log("[partner-driver-be] injected Billing_Entity=" + be + " into save_report");
            }
          }
        }
      } catch(e){}
      return origFetch(url, options);
    };
    console.log("[partner-driver-be] fetch hook installed");
  }

  function checkAndBuild() {
    if (document.getElementById("dr-tourcode") && !document.querySelector(".be-dr-wrap")) {
      buildDropdown();
    } else if (document.getElementById("be-dr-select")) {
      // Refresh on any tour code change
      var tcInput = document.getElementById("dr-tourcode");
      var lastTC = window.__beDrLastTC || "";
      var curTC = tcInput ? tcInput.value : "";
      if (lastTC !== curTC) {
        window.__beDrLastTC = curTC;
        refreshDropdown();
      }
    }
    injectPersonalButton();
  }

  window.__beDrDebug = function(){
    bridgeGlobals();
    var p = getKnownPartners();
    var tcInput = document.getElementById("dr-tourcode");
    return {
      version: "v1",
      hookInstalled: !!window.__beDrFetchHooked,
      currentBE: window._drBillingEntity,
      currentMode: getMode(),
      currentTourCode: tcInput ? tcInput.value : null,
      detectedBE: tcInput ? lookupBEByTourCode(tcInput.value) : null,  // { matched, be }
      partners: p.map(function(x){return x.id;}),
      partnerCount: p.length,
      schedulesCount: getSchedules().length,
      subsCount: getSubArr().length
    };
  };

  function init() {
    injectStyle();
    injectFetchHook();
    bridgeGlobals();
    setInterval(bridgeGlobals, 1500);
    checkAndBuild();
    setInterval(checkAndBuild, 600);
    var obs = new MutationObserver(function(){ checkAndBuild(); });
    obs.observe(document.body, { childList: true, subtree: true });
    // ★ 외부에서 접근 가능하도록 노출 (prefill 후 강제 refresh 등)
    window.__partnerDriverBE = {
      refreshDropdown: refreshDropdown,
      setMode: setMode,
      getMode: getMode,
      lookupBEByTourCode: lookupBEByTourCode,
      bridgeGlobals: bridgeGlobals,
      checkAndBuild: checkAndBuild
    };
    console.log("[partner-driver-be] v2 initialized — call window.__beDrDebug() for diagnostics");
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
