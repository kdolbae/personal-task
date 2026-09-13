# personal-task

내가 벌여놓은 일 전부의 **단일 색인**. 코드 창고가 아니라 목차다.

- 새로 시작할 때: `./new.sh` → 항목 생성 → 아래 표에 한 줄 추가
- "내가 뭐 하고 있었지?" → 이 파일만 보면 된다
- 규칙은 [CONVENTIONS.md](CONVENTIONS.md)

마지막 정리: 2026-09-13

---

## 지금 살아있는 것

| 프로젝트 | 상태 | 무엇 | 스택 | 배포 | 저장소 | 최근 |
|---|---|---|---|---|---|---|
| [나노마스터 광고 도구](projects/nanomaster-ads.md) | 🟢 개발중 | 네이버 검색광고 자동 운영 도구 + 상황판 | HTML | — | [nanomaster-ads](https://github.com/kdolbae/nanomaster-ads) 🔒 | 09-13 |
| [나노마스터 광고 상황판](projects/nanomaster-ads-site.md) | 🚀 배포됨 | 위 도구의 결과를 보여주는 자동 갱신 대시보드 | HTML | [nanomaster-ads-auvk](https://nanomaster-ads-auvk.vercel.app) | [nanomaster-ads-site](https://github.com/kdolbae/nanomaster-ads-site) | 09-12 |
| [mjcon](projects/mjcon.md) | 🚀 배포됨 | *(설명 없음 — 채울 것)* | TypeScript | [mjcon](https://mjcon.vercel.app) | [mjcon](https://github.com/kdolbae/mjcon) 🔒 | 09-12 |
| [개인 관리 프로그램](projects/manageapp-sb.md) | ⚠️ 배포됨 | 개인 관리 프로그램 | TypeScript + Supabase | [manageapp-sb](https://manageapp-sb.vercel.app) | [manageapp_sb](https://github.com/kdolbae/manageapp_sb) 🔒 | 09-12 |
| [나노마스터 웹](projects/nanomaster-web.md) | 🚀 배포됨 | 회사 웹사이트 | CSS 중심 | [nanomaster-web](https://nanomaster-web.vercel.app) | [nanomaster-web](https://github.com/kdolbae/nanomaster-web) 🔒 | 09-11 |

## 실험 중

| 실험 | 시작 | 질문 | 결론 |
|---|---|---|---|
| *(아직 없음)* | | | |

## 정리 대상 / 보관

| 항목 | 왜 | 저장소 |
|---|---|---|
| [nanomaster](archive/nanomaster.md) | 사실상 빈 저장소 (7KB). 이름만 선점된 상태 | [nanomaster](https://github.com/kdolbae/nanomaster) |
| [jipdarie](archive/jipdarie.md) | 완전히 빈 저장소. 8월에 만들고 방치 | [jipdarie](https://github.com/kdolbae/jipdarie) 🔒 |

---

## 지금 눈에 띄는 문제

이 색인을 만들면서 발견한 것들. 고치면 지우면 된다.

- [ ] **`manageapp_sb`의 기본 브랜치가 `codex/uffice-sync`다.** 피처 브랜치가 기본 브랜치로 설정돼 있다. 의도한 게 아니라면 `main`으로 되돌릴 것. 열린 이슈도 1개 있다.
- [ ] **나노마스터 저장소가 4개로 쪼개져 있다** — `nanomaster`(빈 것), `nanomaster-web`, `nanomaster-ads`, `nanomaster-ads-site`. 도구/상황판 쌍은 한 저장소로 합칠 수 있는지 볼 것.
- [ ] **기본 브랜치 이름이 제각각이다** — `nanomaster-web`만 `master`, 나머지는 `main`.
- [ ] **`nanomaster-web`이 29MB다.** CSS가 주 언어인데 이 크기면 이미지·폰트가 저장소에 그대로 들어가 있을 가능성이 높다.
- [ ] **빈 저장소 2개**(`nanomaster`, `jipdarie`)를 쓸지 지울지 결정할 것.

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
