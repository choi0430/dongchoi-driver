/* ============================================================================
 * DC Fleet — Partner Dropdown v3 (additive, standalone)
 * Adds a BE dropdown above TourCode in schedule edit modal.
 * Reads from window._subCompanies (admin.html main cache).
 * Has own fetch hook to ensure BillingEntity persists on save_schedule.
 * ============================================================================ */
(function(){
  'use strict';
  if (typeof window === "undefined") return;

  var PARTNER_DEFAULTS = [
    { id: 'DC',                  label: '\u{1F1F0}\u{1F1F7} DC (자사)',          color: '#1e40af', bg: '#dbeafe', border: '#3b82f6' },
    { id: 'EG TRAVEL PTY LTD',   label: '\u{1F68C} EG TRAVEL (파트너)',  color: '#6d28d9', bg: '#ede9fe', border: '#7c3aed' }
  ];

  var isAdminPage = /admin\.html?/.test(location.pathname) || /Admin/.test(document.title || '');
  if (!isAdminPage) return;

  function injectStyle() {
    if (document.getElementById('be-dropdown-style')) return;
    var s = document.createElement('style');
    s.id = 'be-dropdown-style';
    s.textContent = [
      '#sch-modal .be-tab-row { display: none !important; }',
      '#sch-modal .be-dropdown-wrap { margin-bottom: 12px; }',
      '#sch-modal .be-dropdown-wrap .be-dd-label { font-size: 11px; font-weight: 700; color: var(--t3, #6b7280); margin-bottom: 6px; }',
      '#sch-modal .be-dropdown-wrap .be-dd-select { width: 100%; padding: 10px 12px; border-radius: 8px; border: 2px solid #e5e7eb; background: #f9fafb; font-weight: 700; font-size: 13px; cursor: pointer; transition: all .15s; box-sizing: border-box; }',
      '#sch-modal .be-dropdown-wrap .be-dd-info { margin-top: 6px; font-size: 10px; line-height: 1.5; }'
    ].join('\n');
    document.head.appendChild(s);
  }

  function addCompanyToList(out, known, name) {
    if (name == null) return;
    var s = String(name).trim();
    if (!s) return;
    var key = s.toUpperCase();
    if (known.has(key)) return;
    known.add(key);
    out.push({
      id: s,
      label: '\u{1F69A} ' + s,
      color: '#475569', bg: '#f1f5f9', border: '#64748b'
    });
  }

  function getKnownPartners() {
    var out = PARTNER_DEFAULTS.slice();
    var known = new Set(out.map(function(p){ return String(p.id||"").toUpperCase(); }));

    // Source A: window._subCompanies (admin.html main cache)
    if (Array.isArray(window._subCompanies)) {
      window._subCompanies.forEach(function(item) {
        if (typeof item === "string" || typeof item === "number") addCompanyToList(out, known, item);
        else if (item && typeof item === "object") {
          addCompanyToList(out, known, item.Company || item.SubCompany || item.Sub || item.PartnerCompany || item.name || item.Name);
        }
      });
    }
    // Source B: backups
    ["_priceSubCache", "_subCompanyCache", "_priceSub", "_partners", "_partnerCompanies"].forEach(function(varName) {
      var v = window[varName];
      if (Array.isArray(v)) {
        v.forEach(function(item) {
          if (typeof item === "string" || typeof item === "number") addCompanyToList(out, known, item);
          else if (item && typeof item === "object") {
            addCompanyToList(out, known, item.Company || item.SubCompany || item.Sub || item.PartnerCompany || item.name || item.Name);
          }
        });
      }
    });
    // Source C: DB.PRICE_SUB
    if (window.DB && Array.isArray(window.DB.PRICE_SUB)) {
      window.DB.PRICE_SUB.forEach(function(item) {
        if (item && typeof item === "object") addCompanyToList(out, known, item.Company || item.SubCompany || item.Sub);
      });
    }
    return out;
  }

  function loadSchedule(tourId) {
    if (!tourId) return null;
    var cache = (window._schCache && window._schCache.length) ? window._schCache :
                ((window.DB && window.DB.SCH) ? window.DB.SCH : []);
    if (!Array.isArray(cache) || !cache.length) return null;
    return cache.find(function(s){ return String(s.TourID||"").trim() === String(tourId).trim(); }) ||
           cache.find(function(s){ return String(s.TourCode||"").trim() === String(tourId).trim(); }) ||
           null;
  }

  function getCfgFor(beId, partners) {
    var found = partners.find(function(p){ return String(p.id||"").toUpperCase() === String(beId||"").toUpperCase(); });
    return found || partners[0];
  }

  function isDC(be) {
    return !be || String(be).toUpperCase() === "DC" || String(be).toUpperCase() === "DONGCHOI";
  }

  function buildDropdown() {
    var modal = document.getElementById("sch-modal");
    if (!modal) return false;
    if (modal.querySelector(".be-dropdown-wrap")) return false;
    var tcInput = document.getElementById("sm-tourcode");
    if (!tcInput) return false;
    var target = tcInput.parentElement;
    if (!target) return false;

    var partners = getKnownPartners();

    var wrap = document.createElement("div");
    wrap.className = "be-dropdown-wrap";
    var label = document.createElement("div");
    label.className = "be-dd-label";
    label.textContent = "\u{1F4B0} 인보이스 발행사 (Billing Entity)";
    wrap.appendChild(label);

    var select = document.createElement("select");
    select.className = "be-dd-select";
    partners.forEach(function(p){
      var opt = document.createElement("option");
      opt.value = p.id;
      opt.textContent = p.label;
      select.appendChild(opt);
    });
    wrap.appendChild(select);

    var info = document.createElement("div");
    info.className = "be-dd-info";
    wrap.appendChild(info);

    target.parentElement.insertBefore(wrap, target);

    var tourId = window._schEditTourId || "";
    var sch = loadSchedule(tourId);
    var current = (sch && sch.BillingEntity) ? String(sch.BillingEntity).trim() :
                  (window._schEditBillingEntity || "DC");

    if (current && !partners.find(function(p){ return p.id === current; })) {
      var opt = document.createElement("option");
      opt.value = current;
      opt.textContent = "\u2754 " + current;
      select.appendChild(opt);
    }

    select.value = current;
    window._schEditBillingEntity = current;

    function updateStyle() {
      var cfg = getCfgFor(select.value, partners);
      select.style.borderColor = cfg.border;
      select.style.background = cfg.bg;
      select.style.color = cfg.color;
      if (isDC(select.value)) {
        info.style.color = "#3b82f6";
        info.textContent = "\u2713 DC가 클라이언트에 인보이스를 발행합니다 (자사 운영)";
      } else {
        info.style.color = "#7c3aed";
        info.textContent = "\u26A0 " + select.value + "가 클라이언트에 인보이스를 발행합니다 (DC는 서브 계약자)";
      }
    }
    updateStyle();

    select.addEventListener("change", function(){
      window._schEditBillingEntity = select.value;
      updateStyle();
      console.log("[partner-dropdown] BE selected:", select.value);
    });

    return true;
  }

  function refreshDropdown() {
    var modal = document.getElementById("sch-modal");
    if (!modal) return;
    var select = modal.querySelector(".be-dd-select");
    if (!select) return;
    var tourId = window._schEditTourId || "";
    var sch = loadSchedule(tourId);
    var current = (sch && sch.BillingEntity) ? String(sch.BillingEntity).trim() :
                  (window._schEditBillingEntity || "DC");
    var opts = Array.prototype.map.call(select.options, function(o){ return o.value; });
    if (current && opts.indexOf(current) < 0) {
      var opt = document.createElement("option");
      opt.value = current;
      opt.textContent = "\u2754 " + current;
      select.appendChild(opt);
    }
    select.value = current;
    window._schEditBillingEntity = current;
    select.dispatchEvent(new Event("change"));
  }

  // Defensive own fetch hook — ensures BillingEntity is in save_schedule body
  function injectFetchHook() {
    if (window.__beDdFetchHooked) return;
    window.__beDdFetchHooked = true;
    var origFetch = window.fetch.bind(window);
    window.fetch = function(url, options) {
      var injectedBE = null;
      var injectedTour = null;
      try {
        if (options && options.body && typeof options.body === "string") {
          var body = JSON.parse(options.body);
          if (body && body.action === "save_schedule" && body.data && typeof body.data === "object") {
            var be = window._schEditBillingEntity;
            if (be && be !== body.data.BillingEntity) {
              body.data.BillingEntity = be;
              injectedBE = be;
              injectedTour = body.data.TourID || body.data.TourCode;
              options = Object.assign({}, options, { body: JSON.stringify(body) });
              console.log("[partner-dropdown] injected BE=" + be + " into save_schedule");
            }
          }
        }
      } catch(e){ /* not JSON */ }
      var p = origFetch(url, options);
      // After save success, update local cache so re-open shows new value
      if (injectedBE && injectedTour) {
        p.then(function(res){
          var clone = res.clone();
          clone.json().then(function(j){
            if (j && j.ok) {
              if (Array.isArray(window._schCache)) {
                var sch = window._schCache.find(function(s){
                  return String(s.TourID||"").trim() === String(injectedTour).trim() ||
                         String(s.TourCode||"").trim() === String(injectedTour).trim();
                });
                if (sch) { sch.BillingEntity = injectedBE; console.log("[partner-dropdown] cache updated for " + injectedTour); }
              }
              if (window.DB && Array.isArray(window.DB.SCH)) {
                var sch2 = window.DB.SCH.find(function(s){
                  return String(s.TourID||"").trim() === String(injectedTour).trim() ||
                         String(s.TourCode||"").trim() === String(injectedTour).trim();
                });
                if (sch2) sch2.BillingEntity = injectedBE;
              }
            }
          }).catch(function(){});
        }).catch(function(){});
      }
      return p;
    };
    console.log("[partner-dropdown] fetch hook installed");
  }

  function checkAndBuild() {
    var modal = document.getElementById("sch-modal");
    if (!modal) return;
    var cs = getComputedStyle(modal);
    var visible = cs.display !== "none" && cs.visibility !== "hidden";
    if (!visible) { modal.dataset.beDdInit = ""; return; }
    if (!modal.querySelector(".be-dropdown-wrap")) {
      var ok = buildDropdown();
      if (ok) modal.dataset.beDdInit = String(window._schEditTourId || "new");
    } else {
      var lastId = modal.dataset.beDdInit || "";
      var curId = String(window._schEditTourId || "new");
      if (lastId !== curId) {
        modal.dataset.beDdInit = curId;
        refreshDropdown();
      }
    }
  }

  function init() {
    injectStyle();
    injectFetchHook();
    checkAndBuild();
    setInterval(checkAndBuild, 600);
    var obs = new MutationObserver(function(){ checkAndBuild(); });
    obs.observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ["style", "class"] });
    console.log("[partner-dropdown] v3 initialized");
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
