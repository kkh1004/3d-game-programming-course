# 저장소 안내 — 3d-game-programming-course

> 연성대학교 게임콘텐츠과 · 김강호 교수 운영 도구 모음
> 이 파일은 Claude Code가 세션마다 자동으로 읽는다.

## ⚠️ 가장 중요한 사실

**이 저장소는 하나의 프로젝트가 아니라, 서로 아무 코드도 공유하지 않는 독립 앱 4개가 한 폴더에 나란히 들어있는 형태다.**

`index.html` ~ `index4.html` 은 각각 완전히 다른 앱이며, 이름만 비슷할 뿐 이어지는 번호가 아니다.

### 작업 절대 규칙

1. **한 번에 하나의 앱만 건드린다.** 사용자가 "시간표"를 고쳐달라고 하면 `index4.html` 만 수정한다. 다른 `index*.html` 은 열지도 말고 손대지도 않는다.
2. **공통 리팩터링·일괄 수정 금지.** 4개 파일에 비슷한 코드가 있어도 "공통 모듈로 빼자"는 제안을 먼저 하지 않는다. 각 앱은 독립 배포 단위이고, 하나가 깨지면 실제 수업 운영이 막힌다.
3. **작업 전 해당 앱의 상세 문서를 먼저 읽는다.** (`docs/` 아래, 아래 표의 마지막 열)
4. **파일 이름·위치를 바꾸지 않는다.** GitHub Pages로 실제 서비스 중이라 주소가 깨진다. (구조 개편은 사용자가 따로 요청할 때만)

## 프로젝트 지도

| 파일 | 앱 이름 | 한 줄 설명 | 규모 | 백엔드 | 상세 문서 |
|---|---|---|---|---|---|
| `index.html` | **공지 메일 작성기** | 공지 유형·대상 고르면 메일 초안 생성 → Gmail로 바로 발송 | 1,138줄 | Google OAuth + Gmail API | [docs/01-공지메일작성기.md](docs/01-공지메일작성기.md) |
| `index2.html` | **과목 소개페이지** | 3D게임프로그래밍 홍보용 정적 랜딩 페이지 | 557줄 | 없음 | [docs/02-과목소개페이지.md](docs/02-과목소개페이지.md) |
| `index3.html` | **프로젝트팀 관리 시스템** | 명단 업로드 → 팀 배정 → 팀원 상호평가 → 결과 집계 | 2,321줄 | Apps Script `Code.gs` | [docs/03-프로젝트팀관리.md](docs/03-프로젝트팀관리.md) |
| `index4.html` | **시간표 생성기** | 과목·강의실 등록 → 드래그로 시간표 배치 → 교수별 시수 집계 | 3,193줄 | Apps Script `Timetable.gs` | [docs/04-시간표생성기.md](docs/04-시간표생성기.md) |

## 부속 파일

| 경로 | 소속 | 설명 |
|---|---|---|
| `apps-script/Code.gs` | index3 전용 | 팀 관리 공유 데이터 백엔드 (285줄) |
| `apps-script/Timetable.gs` | index4 전용 | 시간표 공유 데이터 백엔드 (216줄) |
| `img/favicon.png` | 공용 | 파비콘 |
| `img/roster-format-guide.svg` | index3 전용 | 명단 엑셀 형식 안내 이미지 |
| `3D게임프로그래밍_과목-컨텍스트.md` | index2 전용 | 소개페이지 작업 이력·확정 사항 (2026-08-19 기준) |

**두 `.gs` 파일은 서로 별개의 Apps Script 프로젝트로 배포되어 있다.** 데이터 시트도 따로다. 절대 합치지 말 것.

## 공통 기술 성격

- 4개 모두 **빌드 도구 없는 단일 HTML 파일**. npm·번들러·프레임워크 없음. 파일 하나에 HTML+CSS+JS가 전부 들어있다.
- 외부 의존성은 CDN 직접 로드뿐 (SheetJS, Google Identity).
- JS 스타일은 `var` + `function` 기반의 보수적인 ES5 위주. **기존 파일의 스타일을 그대로 따를 것.** 갑자기 `const`/화살표함수/클래스로 바꾸지 않는다.

## 배포

- **GitHub Pages** 로 자동 배포. `master` 에 push하면 몇 분 뒤 반영된다.
- 주소: `https://kkh1004.github.io/3d-game-programming-course/index4.html` 형식
- 구글 OAuth 승인된 자바스크립트 원본: `https://kkh1004.github.io`
- `.gs` 파일은 **저장소에 push해도 반영되지 않는다.** script.google.com에서 직접 붙여넣고 "배포 관리 > 편집 > 버전: 새 버전"으로 재배포해야 한다.

## 커밋 메시지 관례

기존 이력은 영어 명령문 + 버전 표기 형식이다. 그대로 따를 것.

```
Set the operating weeks per grade (v1.13)
Cap a session at 4 hours and say what to change when it will not fit (v1.12)
```

버전이 붙는 건 `index4.html` 뿐이다 (`APP_VERSION` 상수). 나머지는 버전 표기 없음.

## 미해결 사항

- `apps-script/*.gs` 두 파일에 기본 비밀번호 `1004`가 평문으로 있고 저장소가 공개 상태다. 사용자가 정리를 요청하면 처리할 것.
