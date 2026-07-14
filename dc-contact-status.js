// ============================================================================
// DC Fleet — 연락처 탭 드라이버 운행 상태 표시
// 연락처(Contact)에서 드라이버를 누르면, 그 사람이 지금 어떤 차량을
// 운행 중인지(차량번호 · 운행 날짜 · 시작 시간)를 팝업으로 보여줍니다.
//
// 적용 방법: index.html </body> 직전에 아래 한 줄 추가
//   <script src="dc-contact-status.js"></script>
//
// 기존 코드는 전혀 수정하지 않습니다 (런타임 확장 방식).
// ============================================================================
(function () {
  'use strict';

  // ── 1) 운행 상태 모달 DOM 생성 (최초 1회) ──────────────────────────────
  function ensureModal() {
    if (document.getElementById('modal-driver-status')) return;
    var wrap = document.createElement('div');
    wrap.className = 'modal-overlay';        // 앱 기존 모달 스타일 재사용
    wrap.id = 'modal-driver-status';
    wrap.innerHTML =
      '<div class="modal-sheet">' +
        '<div class="modal-handle"></div>' +
        '<div class="modal-title" id="ds-modal-name">👤 —</div>' +
        '<div id="ds-modal-body" style="background:#f0f7ff;border:1.5px solid #bfdbfe;' +
             'border-radius:12px;padding:16px;margin-bottom:16px;"></div>' +
        '<button class="btn btn-secondary" style="width:100%;" ' +
          'onclick="document.getElementById(\'modal-driver-status\').classList.remove(\'open\')">' +
          'OK (확인)</button>' +
      '</div>';
    document.body.appendChild(wrap);
    // 배경(딤) 클릭 시 닫기
    wrap.addEventListener('click', function (e) {
      if (e.target === wrap) wrap.classList.remove('open');
    });
  }

  // ── 2) 모달 열기 — get_active_regos에서 해당 드라이버 조회 ───────────────
  window.showDriverStatusModal = async function (name) {
    ensureModal();
    var nameEl = document.getElementById('ds-modal-name');
    var bodyEl = document.getElementById('ds-modal-body');
    if (nameEl) nameEl.textContent = '👤 ' + name;
    if (bodyEl) {
      bodyEl.innerHTML =
        '<div style="color:#6b7280;font-size:13px;">🔄 운행 상태 확인 중... (Checking...)</div>';
    }
    document.getElementById('modal-driver-status').classList.add('open');

    var info = null;
    try {
      // APPS_SCRIPT_URL 은 index.html의 전역 const — 이후 로드된 스크립트에서 접근 가능
      var res = await fetch(APPS_SCRIPT_URL + '?action=get_active_regos');
      var json = await res.json();
      if (json && json.ok && Array.isArray(json.regos)) {
        info = json.regos.find(function (r) { return r.driver === name; }) || null;
      }
    } catch (e) {
      // 네트워크 오류 — info 없이 진행 (아래에서 '운행 중 아님' 처리)
    }

    // 그새 모달을 닫았으면 무시
    var modal = document.getElementById('modal-driver-status');
    if (!modal || !modal.classList.contains('open')) return;
    if (!bodyEl) return;

    if (info) {
      bodyEl.innerHTML =
        '<div style="font-size:14px;color:#1e3a5f;line-height:2.1;">' +
          '<div>🚌 운행 차량 (Vehicle): <b>' + (info.rego || '—') + '</b></div>' +
          '<div>📅 운행 날짜 (Date): <b>' + (info.date || '—') + '</b></div>' +
          '<div>🕐 시작 시간 (Start Time): <b>' + (info.startTime || '—') + '</b></div>' +
        '</div>';
    } else {
      bodyEl.innerHTML =
        '<div style="font-size:14px;color:#374151;line-height:1.7;">' +
          '🅿️ 현재 운행 중이 아닙니다.<br>' +
          '<span style="font-size:12px;color:#6b7280;">Not currently on a shift.</span>' +
        '</div>';
    }
  };

  // ── 3) renderContactItem 을 감싸서 드라이버 행에 '탭 → 운행 상태' 추가 ────
  function patchRenderContactItem() {
    if (typeof window.renderContactItem !== 'function') return false;
    if (window.renderContactItem.__dcStatusPatched) return true;

    var orig = window.renderContactItem;
    var wrapped = function (c) {
      var html = orig(c);
      // 드라이버(및 driver-guide)만 탭 시 운행 상태 표시 (가이드는 제외)
      if (c && (c.type === 'driver' || c.type === 'driver-guide')) {
        var safeName = String(c.name || '').replace(/'/g, "\\'");
        // 카드 최상위 div 에 커서 + onclick 주입.
        // 전화/문자(<a>) 버튼을 눌렀을 때는 모달이 뜨지 않도록 가드.
        html = html.replace(
          '<div class="contact-item">',
          '<div class="contact-item" style="cursor:pointer;" ' +
            'onclick="if(!event.target.closest(\'a\'))showDriverStatusModal(\'' + safeName + '\')">'
        );
      }
      return html;
    };
    wrapped.__dcStatusPatched = true;
    window.renderContactItem = wrapped;

    // 이미 연락처 화면이 그려져 있으면 새로 렌더
    try {
      if (typeof window.renderContacts === 'function') window.renderContacts();
    } catch (e) {}
    return true;
  }

  // renderContactItem 이 정의될 때까지 잠깐 대기 후 패치
  if (!patchRenderContactItem()) {
    var tries = 0;
    var timer = setInterval(function () {
      tries++;
      if (patchRenderContactItem() || tries > 100) clearInterval(timer);
    }, 100);
  }
})();
