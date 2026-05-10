/* ============================================================================
 * DC Fleet — Partner Dropdown v6 (additive, standalone)
 * Bridges let-declared globals (_subCompanies, _schCache, etc.) to window
 * via dynamically injected <script> tag. Then reads from bridged copies.
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

  // Bridge let-declared globals to window via injected <script>
  function bridgeGlobals() {
    var script = document.createElement("script");
    script.textContent = [
      "(function(){",
      "  try { if (typeof _subCompanies !== \"undefined\") window.__beDdSubCompanies = _subCompanies; } catch(e) {}",
      "  try { if (typeof _schCache !== \"undefined\") window.__beDdSchCache = _schCache; } catch(e) {}",
      "  try { if (typeof _schEditTourId !== \"undefined\") window.__beDdEditTourId = _schEditTourId; } catch(e) {}",
      "  try { if (typeof currentTourId !== \"undefined\") window.__beDdCurrentTourId = currentTourId; } catch(e) {}",
      "  try { if (typeof _schEditBillingEntity !== \"undefined\") window.__beDdEditBE = _schEditBillingEntity; } catch(e) {}",
      "})();"
    ].join("");
    (document.head || document.documentElement).appendChild(script);
    if (script.parentNode) script.parentNode.removeChild(script);
  }

  // Push value back to let-declared global via injected script
  function pushGlobalBE(value) {
    var script = document.createElement("script");
    var safe = String(value || "DC").replace(/"/g, "\\\"");
    script.textContent = [
      "(function(){",
      "  try { if (typeof _schEditBillingEntity !== \"undefined\") _schEditBillingEntity = \"" + safe + "\"; } catch(e) {}",
      "  window._schEditBillingEntity = \"" + safe + "\";",
      "})();"
    ].join("");
    (document.head || document.documentElement).appendChild(script);
    if (script.parentNode) script.parentNode.removeChild(script);
  }

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

  function getSubCompaniesArr() {
    if (Array.isArray(window.__beDdSubCompanies) && window.__beDdSubCompanies.length) return window.__beDdSubCompanies;
    if (Array.isArray(window._subCompanies) && window._subCompanies.length) return window._subCompanies;
    return [];
  }

  function getSchCacheArr() {
    if (Array.isArray(window.__beDdSchCache) && window.__beDdSchCache.length) return window.__beDdSchCache;
    if (Array.isArray(window._schCache) && window._schCache.length) return window._schCache;
    if (window.DB && Array.isArray(window.DB.SCH)) return window.DB.SCH;
    return [];
  }

  function getKnownPartners() {
    var out = PARTNER_DEFAULTS.slice();
    var known = new Set(out.map(function(p){ return String(p.id||"").toUpperCase(); }));
    var sources = [];

    var subArr = getSubCompaniesArr();
    if (subArr.length) sources.push({name: "_subCompanies", count: subArr.length});
    subArr.forEach(function(item){
      var name = tryName(item);
      if (!name) return;
      var key = String(name).trim().toUpperCase();
      if (!key || known.has(key)) return;
      known.add(key);
      out.push({
        id: String(name).trim(),
        label: '\u{1F69A} ' + String(name).trim(),
        color: '#475569', bg: "#f1f5f9", border: "#64748b"
      });
    });

    // Also extract from schedule cache (covers PK NATURE GREEN etc not in M_SUB)
    var schArr = getSchCacheArr();
    schArr.forEach(function(s){
      if (s && s.BillingEntity) {
        var name = String(s.BillingEntity).trim();
        var key = name.toUpperCase();
        if (key && !known.has(key)) {
          known.add(key);
          out.push({
            id: name,
            label: '\u{1F4CD} ' + name,
            color: '#475569', bg: "#f1f5f9", border: "#64748b"
          });
        }
      }
    });

    return { partners: out, sources: sources };
  }

  function getCurrentTourId() {
    if (window.__beDdEditTourId) return String(window.__beDdEditTourId);
    if (window._schEditTourId) return String(window._schEditTourId);
    if (window.__beDdCurrentTourId) return String(window.__beDdCurrentTourId);
    if (window.currentTourId) return String(window.currentTourId);
    var ti = document.getElementById("sm-tourcode");
    if (ti && ti.value) return String(ti.value).trim();
    return "";
  }

  function loadSchedule(tourId) {
    if (!tourId) return null;
    var cache = getSchCacheArr();
    if (!cache.length) return null;
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

  function rebuildSelectOptions(select, partners, currentValue) {
    while (select.firstChild) select.removeChild(select.firstChild);
    partners.forEach(function(p){
      var opt = document.createElement("option");
      opt.value = p.id;
      opt.textContent = p.label;
      select.appendChild(opt);
    });
    if (currentValue && !partners.find(function(p){ return p.id === currentValue; })) {
      var opt = document.createElement("option");
      opt.value = currentValue;
      opt.textContent = "\u2754 " + currentValue;
      select.appendChild(opt);
    }
    select.value = currentValue || "DC";
  }

  function applySelectStyle(select, info, partners) {
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

  function buildDropdown() {
    var modal = document.getElementById("sch-modal");
    if (!modal) return false;
    if (modal.querySelector(".be-dropdown-wrap")) return false;
    var tcInput = document.getElementById("sm-tourcode");
    if (!tcInput) return false;
    var target = tcInput.parentElement;
    if (!target) return false;

    var wrap = document.createElement("div");
    wrap.className = "be-dropdown-wrap";
    var label = document.createElement("div");
    label.className = "be-dd-label";
    label.textContent = "\u{1F4B0} 인보이스 발행사 (Billing Entity)";
    wrap.appendChild(label);

    var select = document.createElement("select");
    select.className = "be-dd-select";
    wrap.appendChild(select);

    var info = document.createElement("div");
    info.className = "be-dd-info";
    wrap.appendChild(info);

    target.parentElement.insertBefore(wrap, target);

    refreshDropdown();

    select.addEventListener("change", function(){
      window._schEditBillingEntity = select.value;
      pushGlobalBE(select.value);
      var p = getKnownPartners();
      applySelectStyle(select, info, p.partners);
      console.log("[partner-dropdown] BE selected:", select.value);
    });

    return true;
  }

  function refreshDropdown() {
    var modal = document.getElementById("sch-modal");
    if (!modal) return;
    var select = modal.querySelector(".be-dd-select");
    var info = modal.querySelector(".be-dd-info");
    if (!select) return;

    bridgeGlobals();

    var tourId = getCurrentTourId();
    var sch = loadSchedule(tourId);
    var current = (sch && sch.BillingEntity) ? String(sch.BillingEntity).trim() :
                  (window.__beDdEditBE || window._schEditBillingEntity || "DC");

    var p = getKnownPartners();
    rebuildSelectOptions(select, p.partners, current);
    window._schEditBillingEntity = current;
    pushGlobalBE(current);
    if (info) applySelectStyle(select, info, p.partners);

    console.log("[partner-dropdown] refresh: tourId=" + tourId + ", BE=" + current + ", options=" + select.options.length + ", subs=" + getSubCompaniesArr().length);
  }

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
            if (be) {
              body.data.BillingEntity = be;
              injectedBE = be;
              injectedTour = body.data.TourID || body.data.TourCode;
              options = Object.assign({}, options, { body: JSON.stringify(body) });
              console.log("[partner-dropdown] injected BE=" + be + " into save_schedule");
            } else {
              console.warn("[partner-dropdown] save_schedule but BE empty");
            }
          }
        }
      } catch(e){}
      var p = origFetch(url, options);
      if (injectedBE && injectedTour) {
        p.then(function(res){
          var clone = res.clone();
          clone.json().then(function(j){
            if (j && j.ok) {
              var cache = getSchCacheArr();
              if (cache.length) {
                var sch = cache.find(function(s){
                  return String(s.TourID||"").trim() === String(injectedTour).trim() ||
                         String(s.TourCode||"").trim() === String(injectedTour).trim();
                });
                if (sch) sch.BillingEntity = injectedBE;
              }
              console.log("[partner-dropdown] cache updated for " + injectedTour);
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
      if (ok) modal.dataset.beDdInit = String(getCurrentTourId() || "new");
    } else {
      var lastId = modal.dataset.beDdInit || "";
      var curId = String(getCurrentTourId() || "new");
      var select = modal.querySelector(".be-dd-select");
      var p = getKnownPartners();
      var optsCount = select ? select.options.length : 0;
      if (lastId !== curId || optsCount < p.partners.length) {
        modal.dataset.beDdInit = curId;
        refreshDropdown();
      }
    }
  }

  window.__beDdDebug = function(){
    bridgeGlobals();
    var p = getKnownPartners();
    var subArr = getSubCompaniesArr();
    return {
      version: "v6",
      hookInstalled: !!window.__beDdFetchHooked,
      currentBE: window._schEditBillingEntity,
      tourId: getCurrentTourId(),
      partners: p.partners.map(function(x){return x.id;}),
      partnerCount: p.partners.length,
      subCount: subArr.length,
      subSample: subArr[0],
      schCacheCount: getSchCacheArr().length
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
    obs.observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ["style", "class"] });
    console.log("[partner-dropdown] v6 initialized — call window.__beDdDebug() for diagnostics");
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
