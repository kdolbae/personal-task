# personal-task

내가 벌여놓은 일 전부의 **단일 색인**. 코드 창고가 아니라 목차다.

- 새로 시작할 때: `./new.sh` → 항목 생성 → 아래 표에 한 줄 추가
- "내가 뭐 하고 있었지?" → 이 파일만 보면 된다
- 규칙은 [CONVENTIONS.md](CONVENTIONS.md)

## 관제탑 (웹 페이지)

[`index.html`](index.html) — 배포·코드·대화 링크로 바로 들어가는 상황판.
Vercel(personal-task-flame)에 올라가 있다. 프로젝트를 고칠 때는 `index.html` 안의 `PROJECTS` 배열만 고치면 된다.

마지막 정리: 2026-09-26 (저장소 전수 점검)

---

## 저장소 구분 — 어디에 무엇을 넣나

새 기능을 어디에 넣을지 헷갈리면 이 질문 순서대로 고른다.

| 질문 | 저장소 |
|---|---|
| 광고비가 나가는 코드인가? (네이버·메타 API, 입찰, 키워드, 상황판 생성) | **nanomaster-ads** |
| 사이트 방문자나 관리자가 브라우저에서 보는 화면인가? (랜딩, /admin, 사이트 자동화) | **nanomaster-web** |
| 엔엠솔루션 회사 사이트인가? | **nmsolution** (광고는 nanomaster-ads의 `nmsolution` 브랜드) |
| 계약·시공·정산·협력업체 관리인가? (유피스를 대체하는 것) | **manageapp** |
| 입주자가 폰에서 쓰는 화면인가? | **jipdarie** (관리 기능은 manageapp API로) |
| 회중 명단인가? | **mjcon** (다른 것과 절대 섞지 않는다) |
| 목차·메모·하루짜리 실험인가? | **personal-task** |

**새 저장소는 만들지 않는다.** 꼭 필요하면 이 표부터 고친 뒤 만든다. 9월에 그렇게 안 해서 12개까지 늘었다.

### 나노마스터 (회사)

| 저장소 | 역할 | 배포 | 배포 브랜치 |
|---|---|---|---|
| [nanomaster-web](https://github.com/kdolbae/nanomaster-web) 🔒 | 회사 사이트 + 관리화면(/admin) + 사이트에서 도는 자동화(유피스 기사·단지 동기화, 단지 수집, 리뷰 수집) | Cloudflare Workers | `claude/gallant-meitner-zvk5ze` (기본 브랜치 `master`가 아님 — 주의) |
| [nanomaster-ads](https://github.com/kdolbae/nanomaster-ads) 🔒 | 광고 실행 코드 전부. 네이버 검색광고·메타 광고 API, 누적 데이터(`data/`), 상황판 HTML 생성(`site/`), DO→메타 맞춤타겟 변환(`scripts/meta_audience/`) | 없음 (Actions 매일 06:35) | `main` |
| [nmsolution](https://github.com/kdolbae/nmsolution) | 엔엠솔루션 회사 사이트 | Vercel | `main` |

상황판 흐름: nanomaster-ads가 `site/index.html`을 만들어 커밋 → nanomaster-web `/admin/ads`가 GitHub API로 읽어 관리자에게만 보여줌. 중간 복사본은 없다.

### 집대리 (사업)

| 저장소 | 역할 | 배포 | 브랜치 |
|---|---|---|---|
| [manageapp](https://github.com/kdolbae/manageapp) 🔒 | 집대리 관리 시스템 (Next.js + Supabase). 협력업체·계약·수납·정산·플랫폼. 유피스(UPIS) 분석 문서 `uffice/`도 여기 | 없음 | `main` (2026-09-26 앱 코드 합침) |
| [jipdarie](https://github.com/kdolbae/jipdarie) 🔒 | 집대리 소비자 앱 (React + Vite + Capacitor). 입주 준비·하자·견적 | 없음 | `claude/friendly-mayer-3gwqw1` → `main`으로 바꿀 것 |

### 개인

| 저장소 | 역할 | 배포 |
|---|---|---|
| [mjcon](https://github.com/kdolbae/mjcon) 🔒 | 만정회중 전도인 관리 (Next.js + Supabase) | Vercel |
| [personal-task](https://github.com/kdolbae/personal-task) 🔒 | 이 색인 | Vercel |

### 보관 (더 이상 손대지 않음)

| 저장소 | 왜 | 어디로 갔나 |
|---|---|---|
| nanomaster-ads-site | 상황판 복사본. 9/18 비공개, 9/26 게시 단계 제거 | nanomaster-ads `site/` |
| advertising- | 메타 맞춤타겟 변환기 1개 | nanomaster-ads `scripts/meta_audience/` |
| nanomaster | README 한 줄 + 카페24 백업 셸. 카페24는 이전 완료 | (필요 없음) |
| manageapp_sb | 유피스 버튼 동기화. 9/9 이후 중단 | nanomaster-web 크론 + manageapp |
| Gridworks | 빈 저장소 | — |

---

## 남은 정리 (사람이 GitHub에서 눌러야 하는 것)

- [ ] 비공개 전환: manageapp, personal-task, nanomaster (지금 public)
- [ ] 보관 처리: nanomaster-ads-site, advertising-, nanomaster, manageapp_sb, Gridworks → Settings → Archive
- [ ] nanomaster 저장소의 `claude/intelligent-clarke-600315` 브랜치 README에 카페24 비밀번호가 적혀 있다. 보관 전에 그 비밀번호를 바꿀 것
- [ ] nanomaster-ads Secrets에서 `SITE_DEPLOY_KEY` 삭제 (더 이상 안 씀)
- [ ] 기본 브랜치: jipdarie → `main`, personal-task → `main`
- [ ] nanomaster-web 기본 브랜치를 `master`에서 gallant로 — 단, 바꾸는 순간 gallant에만 있는 `uffice-sync.yml` 크론(20분마다)이 돌기 시작한다. 그걸 원할 때 바꿀 것. 지금 master에서 도는 `리뷰 자동수집` 크론은 매일 실패 중
- [ ] 머지된 브랜치 삭제: nanomaster-web 17개, nanomaster-ads 13개 (`gh pr list --state merged` 의 head 브랜치)
- [ ] jipdarie의 `claude/upis-contract-management-info-ewsq1o` 브랜치 삭제 (manageapp으로 옮김)
- [ ] manageapp_sb의 `claude/evangelist-management-program-0mma3o` 브랜치 삭제 (mjcon `docs/plan.md`가 상위본)
- [ ] 열린 PR 정리: nanomaster-web 13개(체인), jipdarie 5개(체인), advertising- 2개(#2는 옮겼으니 닫기, #1은 nanomaster-web과 중복), nanomaster-ads #2(카페24, 불필요)

---

## 상태 표시

| | 뜻 |
|---|---|
| 💡 아이디어 | 머릿속에만 있음. 코드 없음 |
| 🧪 실험 | 되는지 보는 중. 버려도 아깝지 않음 |
| 🟢 개발중 | 정식으로 만들고 있음 |
| 🚀 배포됨 | 실제로 돌아가는 중 |
| ⚠️ 손볼 것 | 돌아가지만 문제가 있음 |
| 🧊 중단 | 멈춤. 나중에 볼 수도 |
| 📦 보관 | 끝났거나 접음 |
