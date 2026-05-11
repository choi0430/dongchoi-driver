/* ============================================================================
 * DC Fleet — Balance Filter (v2 - main page + modal)
 * Marks DRSUB entries where tour BillingEntity matches the SubCompany.
 * Works on bal-detail-modal AND main 거래처 잔액 page.
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
      ".be-bf-adjusted-balance { display:inline-block; margin-left:8px; padding:3px 8px; background:#fef9c3; border:1px solid #f59e0b; border-radius:5px; font-weight:700; color:#78350f; font-size:11px; }",
      ".be-bf-main-badge { display:inline-block; margin-left:6px; padding:2px 6px; background:#fbbf24; color:#78350f; border-radius:4px; font-size:10px; font-weight:700; }"
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
    var rxA = /DRSUB:\d{4}-\d{2}-\d{2}_[^_\s]+_(\S+(?:\s\S+)?)/g;
    var m;
    while ((m = rxA.exec(text))) found.push(m[1].trim());
    var rxB = /TC:([^\s\[\]]+(?:\s[^\s\[\]]+)*?)(?=\s\[|\s\u2014|\s\u2022|$|\s{2,})/g;
    while ((m = rxB.exec(text))) found.push(m[1].trim());
    return found;
  }

  function getCurrentSubCompanyFromModal() {
    var modal = document.getElementById("bal-detail-modal");
    if (!modal) return null;
    var titleEls = modal.querySelectorAll("h1, h2, h3, .modal-title, [class*=\"title\"]");
    for (var i = 0; i < titleEls.length; i++) {
      var t = (titleEls[i].textContent || "").trim();
      var m = t.match(/^[\u{1F68C}\u{1F69A}\u{1F1F0}\u{1F1F7}\u{1F4CD}\u{1F4B0}\u{2754}\s]*([A-Z][\w\s\u00C0-\u017F\.]+?(?:PTY\s*LTD|P\/L|PTY|LTD|TOURS|TRAVEL)?)\s*거래/u);
      if (m) return m[1].trim();
    }
    var allText = modal.textContent || "";
    var fm = allText.match(/^[\s\u{1F68C}\u{1F69A}\u{1F1F0}\u{1F1F7}\u{1F4CD}\u{1F4B0}\u{2754}]*([A-Z][\w\s\u00C0-\u017F\.]+?)\s*거래\s*내역/u);
    return fm ? fm[1].trim() : null;
  }

  function processModal(tourBE) {
    var modal = document.getElementById("bal-detail-modal");
    if (!modal) return;
    var cs = getComputedStyle(modal);
    var visible = cs.display !== "none" && cs.visibility !== "hidden";
    if (!visible) { modal.dataset.bfProcessed = ""; return; }
    var subCompany = getCurrentSubCompanyFromModal();
    if (!subCompany) return;
    var key = subCompany + "|" + (window.__beDdBalSubTxns || []).length;
    if (modal.dataset.bfProcessed === key) return;
    modal.dataset.bfProcessed = key;
    modal.querySelectorAll(".be-bf-banner, .be-bf-adjusted-balance").forEach(function(e){ e.remove(); });
    modal.querySelectorAll("[data-bf-info]").forEach(function(e){ e.removeAttribute("data-bf-info"); });
    var result = markInfoEntries(modal, subCompany, tourBE);
    if (result.infoCount > 0) {
      var banner = document.createElement("div");
      banner.className = "be-bf-banner";
      banner.innerHTML = "\u{1F4CC} <b>정보용 (DC 잔액 무관)</b>: " + subCompany + " 가 클라이언트에 직접 인보이스하는 일정의 sub 기록 <b>" + result.infoCount + "건</b>, 합계 <b>$" + result.infoTotal.toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2}) + "</b><span class=\"small\">→ 이 금액은 DC 가 " + subCompany + " 에게 빚지는 금액에서 제외됨</span>";
      var firstChild = modal.firstElementChild;
      if (firstChild) modal.insertBefore(banner, firstChild);
      else modal.appendChild(banner);
      setTimeout(function(){ patchMainBalance(modal, result.infoTotal); }, 50);
    }
  }

  function markInfoEntries(rootEl, subCompany, tourBE) {
    var subUpper = subCompany.toUpperCase();
    var infoTotal = 0, infoCount = 0;
    var seenRows = new Set();
    var allEls = Array.from(rootEl.querySelectorAll("div, tr, li"));
    allEls.sort(function(a, b){ return (a.textContent || "").length - (b.textContent || "").length; });
    allEls.forEach(function(el){
      if (seenRows.has(el)) return;
      if (el.dataset.bfChecked === subUpper) return;
      var text = (el.textContent || "").trim();
      if (text.length < 10 || text.length > 600) return;
      var tcs = extractTourCodes(text);
      if (!tcs.length) return;
      var matched = false;
      for (var i = 0; i < tcs.length; i++) {
        if (tourBE[tcs[i]] && tourBE[tcs[i]] === subUpper) { matched = true; break; }
      }
      if (!matched) return;
      var p = el.parentElement;
      while (p && p !== rootEl) {
        if (seenRows.has(p)) return;
        p = p.parentElement;
      }
      el.setAttribute("data-bf-info", "1");
      el.dataset.bfChecked = subUpper;
      seenRows.add(el);
      el.querySelectorAll("div, tr, li").forEach(function(c){ seenRows.add(c); });
      var amtMatches = text.match(/\$\s*([\d,]+(?:\.\d+)?)/g);
      if (amtMatches && amtMatches.length) {
        var amtStr = amtMatches[amtMatches.length - 1].replace(/[^\d.]/g, "");
        var amt = parseFloat(amtStr);
        if (!isNaN(amt) && amt > 0) { infoTotal += amt; infoCount++; }
      }
    });
    return {infoTotal: infoTotal, infoCount: infoCount};
  }

  function patchMainBalance(rootEl, infoTotal) {
    var allEls = rootEl.querySelectorAll("*");
    for (var i = 0; i < allEls.length; i++) {
      var el = allEls[i];
      if (el.children.length > 0) continue;
      var t = (el.textContent || "").trim();
      if (!/잔\s*액\s*\$[\d,]+/.test(t)) continue;
      if (el.parentElement && !el.parentElement.querySelector(".be-bf-adjusted-balance")) {
        var container = el.parentElement;
        while (container.parentElement && container.tagName !== "TR" && container.parentElement.children.length < 3) {
          container = container.parentElement;
        }
        var origMatch = t.match(/\$([\d,]+(?:\.\d+)?)/);
        if (origMatch) {
          var orig = parseFloat(origMatch[1].replace(/,/g, ""));
          var adjusted = orig - infoTotal;
          var adjEl = document.createElement("span");
          adjEl.className = "be-bf-adjusted-balance";
          adjEl.textContent = "정보용 제외 $" + adjusted.toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2});
          container.appendChild(adjEl);
          break;
        }
      }
    }
  }

  // ── Main page processor ──
  // Find each company "card/section" and process tour entries within it
  function processMainPage(tourBE) {
    // Heuristic: find elements that look like company section headers
    // Company section = container that has "$" and company name + tour code listings
    var candidates = document.querySelectorAll("div, section");
    candidates.forEach(function(section){
      if (section.closest("#bal-detail-modal")) return; // skip modal
      if (section.dataset.bfMainProcessed) return;
      if (section.children.length < 2) return;
      var text = section.textContent || "";
      if (text.length < 50 || text.length > 5000) return;
      // Must contain company name pattern + tour code pattern
      var nameMatch = text.match(/^[\s\u{1F68C}\u{1F69A}\u{1F1F0}\u{1F1F7}\u{1F4CD}]*([A-Z][\w\s\u00C0-\u017F\.]+?(?:PTY\s*LTD|P\/L|PTY|LTD|TOURS|TRAVEL))/u);
      if (!nameMatch) return;
      // Must contain at least one tour code
      var hasTour = /(?:DRSUB:|TC:|\d{4}[\s-][\w-]+|[A-Z]{2,}-?\d+)/.test(text);
      if (!hasTour) return;
      
      var subCompany = nameMatch[1].trim();
      section.dataset.bfMainProcessed = "1";
      var result = markInfoEntries(section, subCompany, tourBE);
      if (result.infoCount > 0) {
        // Find balance display in section header area
        var balanceEl = null;
        section.querySelectorAll("*").forEach(function(el){
          if (balanceEl) return;
          if (el.children.length > 0) return;
          var t = (el.textContent || "").trim();
          if (/^\$[\d,]+\.?\d*$/.test(t)) {
            balanceEl = el;
          }
        });
        if (balanceEl && balanceEl.parentElement && !balanceEl.parentElement.querySelector(".be-bf-main-badge")) {
          var amt = parseFloat(balanceEl.textContent.replace(/[^\d.]/g, ""));
          var adjusted = amt - result.infoTotal;
          var badge = document.createElement("span");
          badge.className = "be-bf-main-badge";
          badge.textContent = "정보용 -$" + result.infoTotal.toLocaleString() + " = $" + adjusted.toLocaleString();
          badge.title = result.infoCount + "건 정보용 (BE=" + subCompany + ")";
          balanceEl.parentElement.appendChild(badge);
          console.log("[bf-filter main] " + subCompany + ": info=$" + result.infoTotal + " (count=" + result.infoCount + "), adjusted=$" + adjusted);
        }
      }
    });
  }

  function processAll() {
    bridgeGlobals();
    var tourBE = getTourBEMap();
    if (Object.keys(tourBE).length === 0) return;
    processModal(tourBE);
    processMainPage(tourBE);
  }

  function init() {
    injectStyle();
    bridgeGlobals();
    setInterval(bridgeGlobals, 1500);
    setInterval(processAll, 1000);
    var obs = new MutationObserver(function(){ processAll(); });
    obs.observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ["style", "class"] });
    console.log("[bf-filter] v2 initialized");
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
