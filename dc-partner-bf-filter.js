/* ============================================================================
 * DC Fleet — Balance Filter (v1)
 * In bal-detail modal, visually mark DRSUB entries where tour BillingEntity
 * matches the SubCompany — these are informational (the partner invoices client
 * directly, so DC has no debt). Shows separate banner with total.
 * ============================================================================ */
(function(){
  'use strict';
  if (typeof window === "undefined") return;
  var isAdminPage = /admin\.html?/.test(location.pathname) || /Admin/.test(document.title || '');
  if (!isAdminPage) return;

  function bridgeGlobals() {
    var script = document.createElement("script");
    script.textContent = [
      "(function(){",
      "  try { if (typeof _schCache !== \"undefined\") window.__beDdSchCache = _schCache; } catch(e) {}",
      "  try { if (typeof _balSubTxns !== \"undefined\") window.__beDdBalSubTxns = _balSubTxns; } catch(e) {}",
      "})();"
    ].join("");
    (document.head || document.documentElement).appendChild(script);
    if (script.parentNode) script.parentNode.removeChild(script);
  }

  function injectStyle() {
    if (document.getElementById("be-bf-style")) return;
    var s = document.createElement("style");
    s.id = "be-bf-style";
    s.textContent = [
      "[data-bf-info=\"1\"] { opacity: 0.55 !important; background: rgba(254, 243, 199, 0.4) !important; border-left: 3px solid #f59e0b !important; }",
      ".be-bf-banner { margin: 10px 14px; padding: 10px 14px; background: #fef3c7; border: 1.5px solid #f59e0b; border-radius: 8px; font-size: 12px; color: #78350f; font-weight: 700; line-height: 1.6; }",
      ".be-bf-banner b { color: #92400e; font-size: 13px; }",
      ".be-bf-banner .small { display:block; font-weight:500; font-size:11px; color:#92400e; margin-top:4px; }",
      ".be-bf-adjusted-balance { display:block; margin-top:6px; padding:6px 10px; background:#fef9c3; border-radius:6px; font-weight:700; color:#78350f; }"
    ].join("\n");
    document.head.appendChild(s);
  }

  function getTourBEMap() {
    var cache = (window.__beDdSchCache && window.__beDdSchCache.length) ? window.__beDdSchCache : (window._schCache || []);
    var map = {};
    if (Array.isArray(cache)) {
      cache.forEach(function(s){
        if (s && s.TourCode) {
          map[String(s.TourCode).trim()] = String(s.BillingEntity || "").trim().toUpperCase();
        }
      });
    }
    return map;
  }

  function extractTourCodes(text) {
    if (!text) return [];
    var found = [];
    // Pattern A: DRSUB:YYYY-MM-DD_VEHICLE_TOURCODE
    var rxA = /DRSUB:\d{4}-\d{2}-\d{2}_[^_\s]+_(\S+(?:\s\S+)?)/g;
    var m;
    while ((m = rxA.exec(text))) found.push(m[1].trim());
    // Pattern B: TC:XXXXX (used in InterCo descriptions)
    var rxB = /TC:([^\s\[\]]+(?:\s[^\s\[\]]+)*?)(?=\s\[|\s—|\s•|$|\s{2,})/g;
    while ((m = rxB.exec(text))) found.push(m[1].trim());
    return found;
  }

  function getCurrentSubCompany() {
    var modal = document.getElementById("bal-detail-modal");
    if (!modal) return null;
    // Try to find a heading element with the company name
    var titleEls = modal.querySelectorAll("h1, h2, h3, .modal-title, [class*=\"title\"], [class*=\"header\"], [class*=\"name\"]");
    for (var i = 0; i < titleEls.length; i++) {
      var t = (titleEls[i].textContent || "").trim();
      var m = t.match(/^[\u{1F68C}\u{1F69A}\u{1F1F0}\u{1F1F7}\u{1F4CD}\u{1F4B0}\u{2754}\s]*([A-Z][\w\s\u00C0-\u017F\.]+?(?:PTY\s*LTD|P\/L|PTY|LTD|TOURS|TRAVEL)?)\s*거래/u);
      if (m) return m[1].trim();
    }
    // Fallback: scan all text for "거래 내역" header
    var allText = modal.textContent || "";
    var fm = allText.match(/^[\s\u{1F68C}\u{1F69A}\u{1F1F0}\u{1F1F7}\u{1F4CD}\u{1F4B0}\u{2754}]*([A-Z][\w\s\u00C0-\u017F\.]+?)\s*거래\s*내역/u);
    return fm ? fm[1].trim() : null;
  }

  function processBalDetailModal() {
    bridgeGlobals();
    var modal = document.getElementById("bal-detail-modal");
    if (!modal) return;
    var cs = getComputedStyle(modal);
    var visible = cs.display !== "none" && cs.visibility !== "hidden";
    if (!visible) {
      modal.dataset.bfProcessed = "";
      return;
    }
    var key = String(getCurrentSubCompany() || "?") + "|" + (window.__beDdBalSubTxns || []).length;
    if (modal.dataset.bfProcessed === key) return;
    modal.dataset.bfProcessed = key;

    // Remove old artifacts
    modal.querySelectorAll(".be-bf-banner, .be-bf-adjusted-balance").forEach(function(e){ e.remove(); });
    modal.querySelectorAll("[data-bf-info]").forEach(function(e){
      e.removeAttribute("data-bf-info");
    });

    var subCompany = getCurrentSubCompany();
    if (!subCompany) {
      console.log("[bf-filter] could not determine subcompany");
      return;
    }
    var subUpper = subCompany.toUpperCase();

    var tourBE = getTourBEMap();
    var infoTotal = 0;
    var infoCount = 0;
    var seenRows = new Set();

    // Find all candidate row elements (use the most leaf-like element to avoid double counting)
    var allEls = Array.from(modal.querySelectorAll("div, tr, li"));
    // Sort by smallest first (innermost)
    allEls.sort(function(a, b){ return (a.textContent || "").length - (b.textContent || "").length; });

    allEls.forEach(function(el){
      if (seenRows.has(el)) return;
      var text = (el.textContent || "").trim();
      if (text.length < 10 || text.length > 600) return;
      var tcs = extractTourCodes(text);
      if (!tcs.length) return;
      var matched = false;
      for (var i = 0; i < tcs.length; i++) {
        if (tourBE[tcs[i]] && tourBE[tcs[i]] === subUpper) { matched = true; break; }
      }
      if (!matched) return;
      // Skip if a parent already counted
      var p = el.parentElement;
      while (p && p !== modal) {
        if (seenRows.has(p)) return;
        p = p.parentElement;
      }
      el.setAttribute("data-bf-info", "1");
      seenRows.add(el);
      // Mark all descendants as seen too
      el.querySelectorAll("div, tr, li").forEach(function(c){ seenRows.add(c); });
      // Try to extract amount from this element
      var amtMatches = text.match(/\$\s*([\d,]+(?:\.\d+)?)/g);
      if (amtMatches && amtMatches.length) {
        var amtStr = amtMatches[amtMatches.length - 1].replace(/[^\d.]/g, "");
        var amt = parseFloat(amtStr);
        if (!isNaN(amt) && amt > 0) {
          infoTotal += amt;
          infoCount++;
        }
      }
    });

    if (infoCount === 0) {
      console.log("[bf-filter] no informational entries for " + subCompany);
      return;
    }

    // Inject banner near top of modal
    var banner = document.createElement("div");
    banner.className = "be-bf-banner";
    banner.innerHTML = "\u{1F4CC} <b>정보용 (DC 잔액 무관)</b>: " + subCompany + " 가 클라이언트에 직접 인보이스하는 일정의 sub 기록 <b>" + infoCount + "건</b>, 합계 <b>$" + infoTotal.toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2}) + "</b><span class=\"small\">→ 이 금액은 DC 가 " + subCompany + " 에게 빚지는 금액에서 제외됨 (해당 일정은 " + subCompany + " 가 자체 인보이스 발행 후 자체 드라이버 페이 처리)</span>";
    var firstChild = modal.firstElementChild;
    if (firstChild) modal.insertBefore(banner, firstChild);
    else modal.appendChild(banner);

    // Try to patch the main balance display
    // Look for "잔액 $XXX" or "총 발생 (DR) $XXX" pattern in modal header area
    setTimeout(function(){ patchMainBalance(modal, infoTotal); }, 50);

    console.log("[bf-filter] " + subCompany + ": " + infoCount + " informational entries, $" + infoTotal);
  }

  function patchMainBalance(modal, infoTotal) {
    // Find element with "잔액" text and a $ amount
    var allEls = modal.querySelectorAll("*");
    for (var i = 0; i < allEls.length; i++) {
      var el = allEls[i];
      if (el.children.length > 0) continue; // leaf only
      var t = (el.textContent || "").trim();
      if (!/잔\s*액\s*\$[\d,]+/.test(t)) continue;
      // Found a balance label - inject adjusted balance below its container
      if (el.parentElement && !el.parentElement.querySelector(".be-bf-adjusted-balance")) {
        var container = el.parentElement;
        // Walk up to a sensible parent
        while (container.parentElement && container.tagName !== "TR" && container.parentElement.children.length < 3) {
          container = container.parentElement;
        }
        var origMatch = t.match(/\$([\d,]+(?:\.\d+)?)/);
        if (origMatch) {
          var orig = parseFloat(origMatch[1].replace(/,/g, ""));
          var adjusted = orig - infoTotal;
          var adjEl = document.createElement("div");
          adjEl.className = "be-bf-adjusted-balance";
          adjEl.textContent = "조정 잔액 (정보용 제외): $" + adjusted.toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2});
          container.appendChild(adjEl);
          break;
        }
      }
    }
  }

  function init() {
    injectStyle();
    bridgeGlobals();
    setInterval(bridgeGlobals, 1500);
    setInterval(processBalDetailModal, 800);
    var obs = new MutationObserver(function(){ processBalDetailModal(); });
    obs.observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ["style", "class"] });
    console.log("[bf-filter] v1 initialized");
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
