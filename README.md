# interx-deck

INTERX 업무 및 기안에 필요한 PPT 제작 Claude Code 플러그인 입니다.

## Install

**Windows** (PowerShell)
```powershell
npx --allow-git=all muramasaDev/interx-deck-skill
```

**macOS** (터미널)
```sh
npx --allow-git=all muramasaDev/interx-deck-skill
```

- Node.js 18 이상, 이 저장소 읽기 권한(GitHub 로그인) 필요
- Python, 패키지, 글꼴은 설치 중 자동 설치
- 설치 뒤 Claude Code 다시 시작
- 업데이트: 같은 명령 (새 판이 나오면 Claude Code 시작 때 알림)
- 점검: 명령 끝에 `doctor` / 제거: 명령 끝에 `uninstall` / 특정 판: 명령 끝에 `#v1.5.1`

## Use

Claude Code 에서 말로 요청

| 명령 | 용도 | 요청 예 |
|---|---|---|
| `/interx-deck:report` | 보고 자료, 고객 보고, 주간보고 | "○○ 결과 보고 덱 만들어 줘" |
| `/interx-deck:design` | 설계 문서, 구성도 | "수집 흐름 구성도 그려 줘" |
| `/interx-deck:okr` | 월간 Individual OKR | "이번 달 OKR 써 줘" (지난달 덱 함께 전달) |
| `/interx-deck:using` | 어느 명령을 쓸지 모를 때 | |

- 덱마다 작성자, 파일 이름, 판 번호를 먼저 확인
- 결과: 작업 폴더에 pptx 및 PDF (`제목_작성자_v버전_YYMMDD`)
- 수정: "3장 표를 아래로" 처럼 말로 요청하거나 PowerPoint 에서 직접 수정

## What it does

- 회사 양식 배치: 표지, 목차, 본문, Thank you 및 패널, 카드, 표, 차트, 일정, 구성도, 수식을 자동 배치
- 편집 가능한 결과물: 모든 상자, 선, 표, 차트를 PowerPoint 에서 바로 수정
- 관문 검사: 겹침, 넘침, 글꼴, 제목 규칙, 금지 표현, AI 투 문장을 저장 전에 검사
- 사람이 쓴 덱 검토: 검토 에이전트가 장마다 결론, 숫자, 용어, 시제를 대조해 고칠 곳을 찾고 다시 검토
- 렌더 확인: 장마다 그림으로 확인하고 PDF 생성
- OKR: 지난달 덱, git 기록, Jira, 캘린더에서 실적 근거를 모아 당월 실적 및 익월 계획 작성

## Notice

- 회사 DRM: PowerPoint 로 한 번 연 파일은 Claude 가 다시 읽지 못함. 받을 사람에게는 PDF 전달
- macOS 에서 PowerPoint 가 없으면 LibreOffice 로 렌더 (줄 바꿈이 조금 다를 수 있음)
- 고객사 이름, 설비 이름, 회사 데이터는 공개 범위가 정해지기 전에 외부 서비스로 보내지 않음
- INTERX 사내 전용. 함께 배포하는 글꼴 Noto Sans KR 은 SIL OFL 1.1

## Release

| 판 | 날짜 | 핵심 내용 |
|---|---|---|
| v1.5.1 | 2026-11-13 | 설치 화면 카타나를 세로로 세워 모빌처럼 회전, 칼날을 얇게 |
| v1.5.0 | 2026-11-13 | 사람이 쓴 덱 수준 검사 (문장 검사, 검토 에이전트의 되풀이 검토), 편집 가능한 산점도 및 히스토그램, 파일 이름 규칙 `제목_작성자_v버전_YYMMDD`, 설치 점검 `doctor`, 새 판 알림 |
| v1.4.1 | 2026-11-13 | 설치 화면 디자인 변경, 판 번호 태그 및 특정 판 설치 |
| v1.4.0 | 2026-11-13 | macOS 지원, 강조색을 남색으로 변경 (주황은 목차 번호만), 논문식 수식 (PowerPoint 수식 개체), AI 티 줄이기 (구어체, 같은 이름표 되풀이, 표 위주 장, 화살표 남용 검사) |
| v1.3.3 | 2026-11-10 | 설치 화면 디자인 |
| v1.3.2 | 2026-11-10 | 파일 이름 규칙 검사, pptx 및 PDF 함께 생성 |
| v1.3.1 | 2026-11-10 | 설치 명령 한 줄, Python 자동 설치 |
| v1.3.0 | 2026-11-10 | OKR 덱 자동 작성 (지난달 덱, git, Jira, 캘린더 근거), 쓰임새별 보고 구성, 글꼴 자동 설치 |
| v1.2.0 | 2026-11-10 | npx 설치 및 제거 |
| v1.1.0 | 2026-11-10 | 개인 설정 분리 (작성자, OKR 정보) |
| v1.0.0 | 2026-11-10 | 첫 판: 보고, 설계, OKR 갈래 및 관문 검사 |

판별 자세한 변경은 [CHANGELOG.md](CHANGELOG.md)
