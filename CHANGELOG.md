# 변경 기록

## 1.3.3 (2026-11-10)
- 설치 화면: 가로 그라데이션 로고, 정보 카드, 단계 표시

## 1.3.2 (2026-11-10)
- 파일 이름 규칙 `YYYYMMDD_파일명_작성자_v버전` (`d.filename(파일명, 버전)`), 관문이 검사
- 결과물 pptx 와 PDF 가 같은 폴더에 같은 이름으로 (렌더 그림만 %TEMP%)

## 1.3.1 (2026-11-10)
- 설치 한 줄: `npx --allow-git=all github:muramasaDev/interx-deck-skill` (npm 12 는 git 설치가 기본으로 꺼져 있어 옵션 필요) (토큰, .npmrc 불필요, git 로그인 사용)
- 설치 명령이 Python(winget, 사용자 범위)과 Python 패키지를 없으면 설치

## 1.3.0 (2026-11-10)
- OKR: 설정 파일 없이 지난달 덱에서 작성자, 계획, KR 행을 읽음 (okr_scan.py), 자료 한 번에 받기, Jira, 캘린더 근거
- OKR: KR 행 1~5 자동 증감, 화면 맞춤, 티켓 번호, 실적 미완 낱말, 달성현황 16자 검사, `--style own`
- report: 쓰임새 5종 (사내보고, 회의, 고객, 파트너, 문서 배포), `Deck(forbid=[정규식])` 금지 문구 검사
- check.py 주의: 마지막 줄에 낱말 하나, 표 장 35% 넘음, 같은 구도 3장 연속
- 설치 명령이 Noto Sans KR 을 사용자 글꼴로 설치, 관문이 PDF 도 만듦 (`render.ps1 -Pdf`, DRM 이 암호화하지 않게 다시 씀)
- render.ps1 이 사용자가 열어 둔 PowerPoint 창을 닫지 않음

## 1.2.0 (2026-11-10)
- npm 배포: `npx @muramasadev/interx-deck` 로 설치와 새 판 받기, `uninstall` 로 지우기 (GitHub Packages)
- 설치 명령은 관리자 원본 작업 폴더(git)와 손으로 둔 폴더를 덮지 않음, Python 패키지와 Noto Sans KR 이 없으면 알려 줌
- 개인 설정 위치와 엔진 찾기가 CLAUDE_CONFIG_DIR 을 따름
- 설치 화면: INTERX DECK 픽셀 배너, 점검 결과 표시
- 주석 정리, README 를 GitHub 용으로 (사용법.md 합침), render.ps1 하나로

## 1.1.0 (2026-11-10)
- 배포용 정리: 개인 정보(작성자, OKR 프로필, 근거 저장소)를 플러그인 밖 개인 설정(~/.claude/interx-deck/profile)으로 분리, 처음 쓸 때 묻고 저장
- 과제 고유 규칙은 개인 설정 projects/<과제>/README.md 로, 공통 스킬은 어느 과제에나 맞는 규칙만
- 경로를 ${CLAUDE_PLUGIN_ROOT} 로 (마켓플레이스 설치본에서도 동작), 출발 파일이 엔진 위치를 찾음
- 사내 git 저장소 마켓플레이스(interx) 추가

## 1.0.0 (2026-11-10)
- 갈래 스킬 넷(using, report, design, okr)으로 나눔, 보고 자료 규칙과 OKR 덱 스킬을 합침, 수식 상자 formula
