/* ============================================================================
 * DC Fleet — Driver Maintenance Alert (v1)
 * Shows vehicle maintenance status (overdue / upcoming) in driver app.
 * Bridges M_Vehicles data + detects current vehicle from Pre-Departure.
 * ============================================================================ */
(function(){
  'use strict';
  if (typeof window === "undefined") return;
  var isDriverPage = !/admin\.html?/.test(location.pathname) && !/Admin/.test(document.title || '');
  if (!isDriverPage) return;

  function bridgeGlobals() {
    var script = document.createElement("script");
    script.textContent = [
      "(function(){",
      "  try { if (typeof _vehicles !== \"undefined\") window.__beDdVehicles = _vehicles; } catch(e) {}",
      "  try { if (typeof _activeRegos !== \"undefined\") window.__beDdActiveRegos = _activeRegos; } catch(e) {}",
      "  try { if (typeof _activeShifts !== \"undefined\") window.__beDdActiveShifts = _activeShifts; } catch(e) {}",
      "})();"
    ].join("");
    (document.head || document.documentElement).appendChild(script);
    if (script.parentNode) script.parentNode.removeChild(script);
  }

  function injectStyle() {
    if (document.getElementById("be-maint-style")) return;
    var s = document.createElement("style");
    s.id = "be-maint-style";
    s.textContent = [
      ".be-maint-alert { margin: 10px 14px; padding: 12px 14px; border-radius: 10px; font-size: 13px; line-height: 1.6; font-weight: 600; }",
      ".be-maint-overdue { background: #fee2e2; border: 1.5px solid #dc2626; color: #991b1b; }",
      ".be-maint-soon { background: #fef3c7; border: 1.5px solid #f59e0b; color: #78350f; }",
      ".be-maint-ok { background: #d1fae5; border: 1.5px solid #10b981; color: #065f46; }",
      ".be-maint-alert .be-maint-title { font-weight: 700; font-size: 14px; display: block; margin-bottom: 4px; }",
      ".be-maint-alert .be-maint-detail { font-size: 12px; font-weight: 500; }"
    ].join("\n");
    document.head.appendChild(s);
  }

  function getVehicles() {
    if (Array.isArray(window.__beDdVehicles) && window.__beDdVehicles.length) return window.__beDdVehicles;
    if (Array.isArray(window._vehicles) && window._vehicles.length) return window._vehicles;
    return [];
  }

  function findVehicleByRego(rego) {
    if (!rego) return null;
    var regoU = String(rego).trim().toUpperCase();
    var list = getVehicles();
    return list.find(function(v){ return String(v.Rego||"").trim().toUpperCase() === regoU; }) || null;
  }

  function getCurrentVehicleRego() {
    // 1) Try Pre-Departure form input
    var preRego = document.getElementById("pre-rego") || document.querySelector("[id*=\"pre-rego\"]");
    if (preRego && preRego.value) return String(preRego.value).trim();
    // 2) Try active shift display
    var shift = window.__beDdActiveShifts || window._activeShifts || [];
    if (Array.isArray(shift) && shift.length) {
      var s = shift.find(function(x){ return x && x.Rego && !x.End_Time; });
      if (s) return String(s.Rego).trim();
    }
    return "";
  }

  function buildAlert(vehicle) {
    if (!vehicle || !vehicle.Rego) return null;
    var current = Number(vehicle.Current_KM || 0);
    var next = Number(vehicle.Next_Service_KM || 0);
    if (!next) return null;
    var remain = next - current;
    var wrap = document.createElement("div");
    var status;
    if (remain < 0) {
      status = "overdue";
      wrap.className = "be-maint-alert be-maint-overdue";
      wrap.innerHTML = "<span class=\"be-maint-title\">\u{1F6A8} 정비 초과 (Service OVERDUE)</span>" +
        "<span class=\"be-maint-detail\">차량 " + vehicle.Rego + " — 현재 " + current.toLocaleString() + "km, 다음 정비 " + next.toLocaleString() + "km. <b>초과 " + Math.abs(remain).toLocaleString() + "km</b>. 정비 후 운행 권장합니다.</span>";
    } else if (remain < 1500) {
      status = "soon";
      wrap.className = "be-maint-alert be-maint-soon";
      wrap.innerHTML = "<span class=\"be-maint-title\">\u{1F527} 정비 임박 (Service Soon)</span>" +
        "<span class=\"be-maint-detail\">차량 " + vehicle.Rego + " — 현재 " + current.toLocaleString() + "km, 다음 정비 " + next.toLocaleString() + "km. <b>" + remain.toLocaleString() + "km 남음</b>. 곧 정비 예약 필요.</span>";
    } else {
      return null;
    }
    wrap.dataset.beMaintStatus = status;
    wrap.dataset.beMaintRego = vehicle.Rego;
    return wrap;
  }

  function findInsertionParent() {
    // Try main app container near license expiry alert (top of page)
    var candidates = [
      document.querySelector("main"),
      document.querySelector(".main-content"),
      document.querySelector(".app-content"),
      document.querySelector("#app"),
      document.body
    ];
    return candidates.find(function(c){ return c; }) || document.body;
  }

  function processMaintenance() {
    bridgeGlobals();
    var rego = getCurrentVehicleRego();
    if (!rego) {
      // No current vehicle — clear any existing alert
      document.querySelectorAll(".be-maint-alert").forEach(function(e){ e.remove(); });
      return;
    }
    var vehicle = findVehicleByRego(rego);
    if (!vehicle) return;
    
    var existing = document.querySelector(".be-maint-alert[data-be-maint-rego=\"" + vehicle.Rego + "\"]");
    var newAlert = buildAlert(vehicle);
    
    if (existing && !newAlert) {
      existing.remove();
      return;
    }
    if (existing && newAlert) {
      // Check if status changed
      if (existing.dataset.beMaintStatus !== newAlert.dataset.beMaintStatus) {
        existing.replaceWith(newAlert);
      }
      return;
    }
    if (!existing && newAlert) {
      // Clear any other alerts (different rego), insert new
      document.querySelectorAll(".be-maint-alert").forEach(function(e){ e.remove(); });
      var parent = findInsertionParent();
      // Try to insert near top — before first major content
      var firstChild = parent.firstElementChild;
      if (firstChild) parent.insertBefore(newAlert, firstChild);
      else parent.appendChild(newAlert);
      console.log("[driver-maint] alert inserted: " + vehicle.Rego + " status=" + newAlert.dataset.beMaintStatus);
    }
  }

  window.__beMaintDebug = function(){
    bridgeGlobals();
    var rego = getCurrentVehicleRego();
    var vehicle = findVehicleByRego(rego);
    return {
      version: "v1",
      currentRego: rego,
      vehicleFound: !!vehicle,
      vehicleSample: vehicle,
      vehiclesCount: getVehicles().length,
      activeShifts: (window.__beDdActiveShifts || []).length
    };
  };

  function init() {
    injectStyle();
    bridgeGlobals();
    setInterval(bridgeGlobals, 2000);
    setInterval(processMaintenance, 1500);
    var obs = new MutationObserver(function(){ processMaintenance(); });
    obs.observe(document.body, { childList: true, subtree: true });
    console.log("[driver-maint] v1 initialized");
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
