# DC Fleet 컴포넌트 컨벤션 가이드

**목적:** 새로 작성하거나 수정하는 코드는 이 가이드를 따라주세요. 기존 코드는 건드리지 않아도 됩니다 (점진적 정리).

**대상 파일:** `admin.html`, `index.html`

**작성일:** 2026-04-29

---

## 📐 핵심 원칙

1. **기존 코드와 시각적으로 일관**되게 유지
2. **CSS 변수 사용** — 절대 하드코딩 색상 금지 (`#fff` ❌ → `var(--sf)` ✅)
3. **인라인 스타일 최소화** — 가능하면 클래스 사용
4. **한 줄 한 책임** — 모달 생성 / 데이터 처리 / 토스트는 분리

---

## 🎨 CSS 변수 (디자인 토큰)

이미 정의된 변수만 사용하세요. 새 변수 추가는 별도 협의 필요.

### 배경 (Surface)
| 변수 | 용도 |
|---|---|
| `--bg` | 페이지 최상단 배경 |
| `--sf` | 카드/모달 본문 배경 |
| `--sf1` | 약간 어두운 영역 (테이블 줄무늬 등) |
| `--sf2` | 더 어두운 영역 (헤더, 비활성 버튼) |
| `--sf3` | 가장 어두운 영역 |

### 테두리
| 변수 | 용도 |
|---|---|
| `--b1` | 옅은 테두리 (1px 구분선) |
| `--b2` | 표준 테두리 (input, 카드) |
| `--brd` | 강조 테두리 |

### 텍스트
| 변수 | 용도 |
|---|---|
| `--t1` | 본문 (가장 진함) |
| `--t2` | 보조 텍스트 |
| `--t3` | 힌트, 라벨, 작은 글씨 |

### 의미 색상
| 변수 | 용도 |
|---|---|
| `--ac` / `--acd` | Accent (주요 액션) — 보라색 |
| `--gn` / `--gnd` | 성공 / 수입 — 초록 |
| `--rd` / `--rdd` | 에러 / 경고 — 빨강 |
| `--or` / `--ord` | 주의 — 주황 |
| `--yl` / `--yld` | 노란색 (강조) |
| `--pu` / `--pud` | 보라 (특수 액션) |

### 폰트
| 변수 | 용도 |
|---|---|
| `--fn` | 숫자 전용 (mono) |
| `--fm` | 본문 |

---

## 🔘 버튼

### 표준 클래스 (그대로 사용)

```html
<!-- 주요 액션 (저장, 발행 등) -->
<button class="btn bac">저장</button>

<!-- 보조 액션 (취소, 닫기 등) -->
<button class="btn brd">취소</button>

<!-- 호버 강조 (toggle 등) -->
<button class="btn bgh">필터</button>

<!-- 크기 변형 -->
<button class="btn bac sm">작은 버튼</button>
<button class="btn bac xs">아주 작은 버튼</button>

<!-- 활성 상태 (toggle) -->
<button class="btn bac sm on">선택됨</button>
```

### 클래스 조합표

| 클래스 | 의미 |
|---|---|
| `bac` | Background Accent (보라색 배경, 흰 글씨) |
| `brd` | Border (테두리만, 배경 투명) |
| `bgh` | Background Hover (호버 시 강조) |
| `sm` | Small |
| `xs` | Extra Small |
| `on` | 활성 상태 |

### ⚠️ 하지 말 것

```html
<!-- ❌ 인라인 스타일로 색깔 직접 지정 -->
<button style="background:#7c3aed;color:#fff;padding:8px 16px;">저장</button>

<!-- ✅ 표준 클래스 사용 -->
<button class="btn bac">저장</button>
```

```html
<!-- ❌ 하드코딩 색상 -->
<button class="btn" style="background:#10b981;">완료</button>

<!-- ✅ CSS 변수 사용 (필요한 경우만) -->
<button class="btn" style="background:var(--gn);color:#fff;">완료</button>
```

### 특수 케이스 (의미 색상 강조 시)

지급, 삭제, 취소 같이 의미가 강한 버튼은 인라인 스타일 허용:

```html
<!-- 지급 처리 (초록) -->
<button class="btn bac sm" style="background:var(--gn);color:#fff;">💰 지급 처리</button>

<!-- 삭제 (빨강) -->
<button class="btn sm" style="background:var(--rd);color:#fff;">🗑 삭제</button>
```

---

## 🪟 모달

### 전역 모달 (정보 표시)

이미 만들어진 헬퍼 사용:

```javascript
// 표시
showGenModal({
  title: '인보이스 상세',
  body: htmlContent,  // HTML 문자열
  width: 'min(90vw,500px)'  // 선택
});

// 닫기
closeGenModal();
```

### 커스텀 모달 (입력 폼 등)

직접 만드는 경우 다음 패턴 사용:

```javascript
function openMyModal() {
  const overlay = document.createElement('div');
  overlay.id = 'my-modal-overlay';  // 고유 ID
  overlay.style.cssText = `
    position:fixed;inset:0;z-index:10000;
    background:rgba(0,0,0,.45);
    display:flex;align-items:center;justify-content:center;
  `;
  overlay.innerHTML = `
    <div style="background:var(--sf);border-radius:16px;padding:20px;width:92%;max-width:380px;box-shadow:0 8px 32px rgba(0,0,0,.25);">
      <div style="font-size:15px;font-weight:800;color:var(--t1);margin-bottom:12px;">
        🎯 모달 제목
      </div>
      <!-- 본문 -->
      <div style="display:flex;gap:8px;margin-top:14px;">
        <button id="my-cancel-btn" style="flex:1;padding:10px;border-radius:10px;border:1px solid var(--b2);background:var(--sf2);color:var(--t2);font-size:13px;font-weight:700;cursor:pointer;">취소</button>
        <button id="my-save-btn" class="btn bac" style="flex:1;">저장</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);

  // 바깥 클릭으로 닫기
  overlay.addEventListener('click', e => { if(e.target === overlay) overlay.remove(); });
  document.getElementById('my-cancel-btn').onclick = () => overlay.remove();
  document.getElementById('my-save-btn').onclick = async () => {
    // 저장 로직
    overlay.remove();
  };
}
```

**참고 예시:** `markSubInvPaid()` 함수의 모달 (admin.html line 12613 부근)

### 모달 일관성 체크리스트

- [ ] z-index: `10000` (위로)
- [ ] 배경: `rgba(0,0,0,.45)` (반투명)
- [ ] 본문 배경: `var(--sf)` (다크모드 대응)
- [ ] 둥근 모서리: `border-radius:16px`
- [ ] 패딩: `padding:20px`
- [ ] 최대 너비: `max-width:380px` ~ `max-width:600px`
- [ ] 그림자: `box-shadow:0 8px 32px rgba(0,0,0,.25)`
- [ ] 바깥 클릭으로 닫기 가능

---

## 🍞 토스트 (알림)

```javascript
toast('메시지', 'type');
```

### 타입

| 타입 | 색깔 | 용도 |
|---|---|---|
| `'ok'` | 초록 | 성공 |
| `'er'` | 빨강 | 에러 |
| `'in'` | 파랑 | 진행 중 (info) |
| `'warn'` | 노랑 | 경고 |

### 예시

```javascript
toast('💾 저장 중...', 'in');
toast('✅ 저장 완료', 'ok');
toast('❌ 저장 실패: 네트워크 오류', 'er');
toast('⚠️ 금액을 확인하세요', 'warn');
```

### ⚠️ 하지 말 것

```javascript
// ❌ alert 사용 (사용자 흐름 차단)
alert('저장되었습니다');

// ✅ toast 사용 (자동 사라짐)
toast('✅ 저장 완료', 'ok');
```

```javascript
// ❌ confirm 사용 (사용자 흐름 차단)
if (confirm('정말 삭제하시겠습니까?')) { ... }

// ✅ 모달 사용 (예시: deleteBalTxn 같은 패턴 참고)
showGenModal({
  title: '삭제 확인',
  body: '...정말 삭제?',
  ...
});
```

---

## 📝 입력 필드

### 표준 input 스타일

```html
<input type="text" 
  style="width:100%;padding:9px 10px;border:1px solid var(--b2);border-radius:8px;font-size:13px;box-sizing:border-box;"
  placeholder="입력하세요">
```

### 라벨

```html
<label style="font-size:11px;font-weight:600;color:var(--t2);display:block;margin-bottom:3px;">
  지급 금액 ($)
</label>
```

### 일반적 패턴

```html
<div style="margin-bottom:10px;">
  <label style="font-size:11px;font-weight:600;color:var(--t2);display:block;margin-bottom:3px;">
    필드 라벨
  </label>
  <input type="text" style="width:100%;padding:9px 10px;border:1px solid var(--b2);border-radius:8px;font-size:13px;box-sizing:border-box;">
</div>
```

---

## 🃏 카드 (crd 클래스)

```html
<div class="crd" style="padding:8px 10px;margin-bottom:6px;">
  <!-- 카드 내용 -->
</div>
```

`crd` 클래스가 자동으로 다음 스타일 적용:
- 배경 `var(--sf)`
- 테두리 `1px solid var(--b1)`
- 둥근 모서리

### 클릭 가능한 카드

```html
<div class="crd" style="padding:8px 10px;margin-bottom:6px;cursor:pointer;" 
     onclick="handleClick()">
  <!-- 카드 내용 -->
</div>
```

---

## 🏷️ 배지 (Badge)

상태 표시용 작은 라벨:

```html
<!-- 일반 상태 -->
<span style="font-size:9px;padding:1px 6px;border-radius:4px;font-weight:700;background:var(--gn);color:#fff;">PAID</span>

<!-- 미지급 -->
<span style="font-size:9px;padding:1px 6px;border-radius:4px;font-weight:700;background:var(--or);color:#fff;">미지급</span>

<!-- 부분지급 -->
<span style="font-size:9px;padding:1px 6px;border-radius:4px;font-weight:700;background:#f59e0b;color:#fff;">부분 $X/$Y</span>
```

배지 색상 매핑:

| 의미 | 배경 |
|---|---|
| 성공/완료 | `var(--gn)` |
| 에러/긴급 | `var(--rd)` |
| 미완료 | `var(--or)` |
| 부분/진행 중 | `#f59e0b` |
| 정보 | `var(--ac)` |
| 비활성 | `var(--t3)` |

---

## 💰 금액 표시

숫자는 항상 호주 로케일로 포맷:

```javascript
// ✅ 표준
const amt = 3310;
const formatted = '$' + amt.toLocaleString('en-AU', {minimumFractionDigits:2, maximumFractionDigits:2});
// → "$3,310.00"

// ❌ 하지 말 것
const formatted = '$' + amt;  // "$3310" — 천단위 콤마 없음
```

### 폰트는 mono 사용

```html
<span style="font-family:var(--fn);">$3,310.00</span>
```

### 색상 (수입/지출)

```html
<!-- 수입 (DR, 받을 돈) -->
<span style="color:var(--ac);font-family:var(--fn);">$3,310.00</span>

<!-- 지출 (CR, 지급) -->
<span style="color:var(--gn);font-family:var(--fn);">$3,310.00</span>

<!-- 잔액 음수 -->
<span style="color:var(--rd);font-family:var(--fn);">$-100.00</span>
```

---

## 📅 날짜 표시

### 표시용 (사용자에게 보임)

```javascript
fmtDate('2026-04-25')        // → "25/04/2026"
fmtDateTime('2026-04-25T10:30')  // → "25/04/2026 10:30"
```

### 내부 비교/정렬용

```javascript
parseToISO('25/04/2026')  // → "2026-04-25"
```

### ⚠️ 직접 비교 금지

```javascript
// ❌ 위험 — 형식에 따라 다른 결과
if (r.Date === '2026-04-25') ...

// ✅ 안전 — 항상 ISO로 정규화
if (parseToISO(r.Date) === '2026-04-25') ...
```

---

## 🔗 투어코드

투어코드는 시트마다 다른 필드명으로 저장될 수 있어요. 항상 헬퍼 사용:

```javascript
// ✅ DR 행 (Daily_Report, Schedule)
const tc = pickTourCode(r);

// ✅ 인보이스 Items 안의 항목
const tc = pickItemTourCode(it);
```

**절대 직접 접근 금지:**

```javascript
// ❌ 한 필드만 봄 — 다른 변형은 놓침
const tc = r.Tour_Code;

// ✅ 모든 변형 처리
const tc = pickTourCode(r);
```

---

## 🎯 함수 네이밍

### 권장 컨벤션

| 패턴 | 용도 | 예시 |
|---|---|---|
| `_xxx` | 내부 헬퍼 (파일 외부 노출 안 함) | `_buildSubSettlementHTML` |
| `xxx` | 공개 함수 | `markSubInvPaid` |
| `renderXxx` | UI 렌더링 | `renderDashboard` |
| `loadXxx` | 데이터 로드 | `loadAllFromSheets` |
| `saveXxx` | 데이터 저장 | `saveDREdit` |
| `openXxx` / `closeXxx` | 모달 열기/닫기 | `openBalTxnModal` |
| `onXxxChange` | 이벤트 핸들러 | `onSettleTourCodeSearch` |

### window 노출

전역에서 호출 가능해야 하면 명시적으로:

```javascript
async function markSubInvPaid(invNum){
  // ...
}
window.markSubInvPaid = markSubInvPaid;
```

---

## 📦 GAS 호출

### 표준 패턴

```javascript
const res = await fetch(APPS_SCRIPT_URL, {
  method: 'POST',
  headers: { 'Content-Type': 'text/plain' },
  body: JSON.stringify({
    action: 'action_name',
    data: { /* ... */ },
    _user: getAdminUser()
  })
});
const json = await res.json();
if (json.ok) {
  // 성공
} else {
  console.error('Failed:', json.error);
  toast('❌ 실패: ' + (json.error || 'unknown'), 'er');
}
```

### ⚠️ 캐시 동기화

GAS 호출 성공 후 **로컬 캐시도 갱신**해야 합니다:

```javascript
// 예: SUB_Txn 추가 후
if (json.ok) {
  _balSubTxns.push({...txnData, _rowIndex: json.row || json.rowIndex});
  if (typeof renderBalSubList === 'function') renderBalSubList();
}
```

---

## ✅ 새 기능 추가 체크리스트

새 모달/페이지/기능을 추가할 때 확인:

- [ ] 인라인 색상 하드코딩 없음 (`#fff` ❌ → `var(--sf)`)
- [ ] 버튼은 `btn bac` / `btn brd` / `btn bgh` 클래스 사용
- [ ] alert / confirm 안 씀 → toast / 모달 사용
- [ ] 금액은 `toLocaleString('en-AU', {minimumFractionDigits:2})` 포맷
- [ ] 날짜 비교는 `parseToISO()` 거쳐서
- [ ] 투어코드는 `pickTourCode()` / `pickItemTourCode()` 사용
- [ ] GAS 호출 후 로컬 캐시 갱신
- [ ] 에러 처리 — try/catch + toast
- [ ] 함수 네이밍 컨벤션 (`_internal` vs `public`)
- [ ] 새 함수가 전역 호출 필요하면 `window.x = x`

---

## 🚧 향후 정리할 것 (메모)

이 가이드는 **신규 코드용**이고 기존 코드는 그대로 둡니다. 다만 시간이 나면 점진적으로 다음을 정리하면 좋습니다 (위험도 낮은 순):

1. **하드코딩 색상 → CSS 변수** (검색해서 일괄 치환 가능, 위험도 낮음)
2. **인라인 버튼 스타일 → btn 클래스** (페이지 단위로 진행, 중간 위험도)
3. **alert/confirm 잔재 제거** (toast로 교체, 낮은 위험도)
4. **모달 패턴 통일** (가장 큰 작업, 신중히)

각각 별도 PR로 진행하고 충분히 테스트하세요.

---

**문서 끝.** 질문이나 새 컨벤션 추가 제안은 Branden과 상의.
