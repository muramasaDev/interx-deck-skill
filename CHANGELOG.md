# 변경 기록

## 1.5.0 (2026-11-13)
- 사람이 쓴 덱: 문장 검사 `check_text.py` (AI 투 수식어, 서술체, 느낌표는 실패, 의인화, 명사 넷 잇기, 긴 줄, 같은 시작 되풀이는 주의), 관문이 `text.md` 를 만들고 검토 지시문(`references/review_prompt.md`)으로 '그대로 사용 가능' 판정까지 되풀이
- 파일 이름 규칙 `제목_작성자_v버전_YYMMDD` (띄어쓰기는 밑줄)
- `Deck(allow=[정규식])`: 화면 문구처럼 원문 그대로 옮긴 글을 검사에서 뺌
- 편집 가능한 산점도 `scatter`, 히스토그램 `hist`. 강조 막대 하나일 때 나머지는 회색
- `regress.py`: 다른 과제 덱을 새 엔진으로 다시 빌드해 검사 (결과물은 임시 폴더)
- 설치 점검 `doctor`, Claude Code 시작 때 새 판 알림 (SessionStart 훅)
- 개인 설정 `template_source`: 회사 양식 원본이 더 새로우면 알림
- 예시 덱(example_build.py)을 한 과제 이야기로 다시 씀: 요약 줄은 결론, 장 사이 숫자와 이름 일치 (검토 에이전트로 되풀이 검토)
- 표지 제목이 한 줄을 넘으면 멈춤, 간트 세로선 '기준일', 히스토그램 건수 단위 지정, 수식 첨자 줄 높이
- 산점도 축 이름도 금지 표현과 문장 검사 대상, `allow` 는 정규식마다 따로 적용

## 1.4.1 (2026-11-13)
- 설치 화면: 카타나를 화면 높이와 폭에 맞춰 크게, 물결 무늬와 피, 멈춘 뒤 화면 폭까지 커짐
- 설치 화면: 만든 사람 이름을 큰 강철 글자로, 베인 선과 피
- 배포마다 판 번호 태그 (v1.4.1), 특정 판 설치 `#v1.4.1`

## 1.4.0 (2026-11-13)
- Mac 지원: 렌더 render.py (Windows 는 render.ps1, Mac 은 PowerPoint for Mac 또는 LibreOffice), 설치 명령이 python3, Homebrew, ~/Library/Fonts 처리
- 색: 강조색 ACCENT 남색(002060). 번호 원, 이름표, 핵심 수치, 지금 단계, 핵심 행, 주 흐름. 대제목 번호는 글과 같은 회색, 캡처 네모는 빨강, 주황은 목차 번호만
- 이름표 묶음: 연회색 테두리와 왼쪽 위 네모 이름표, 이름표가 명사가 아니면 경고
- 수식: `formula` 가 PowerPoint 수식 개체(Cambria Math)로 그림. 분수, 아래첨자, 기울인 한 글자 기호
- 글쓰기: 구어 부사(늘, 딱, 그만큼 등), '새 ~', 곱하기 'x' 를 금지 표현 검사에 추가, 읽기 쉬운 문장 네 가지 규칙
- 설치 명령 줄임: `npx --allow-git=all muramasaDev/interx-deck-skill`. 설치 화면은 도는 카타나 (村正)
- 관문, 렌더 하위 실행을 `python -P` 로 (사용자 범위 패키지도 보임), 출력은 UTF-8
- 주석을 짧은 개조식으로 정리
- AI 티 줄이기: 같은 이름표 3장 되풀이 잡음, 메모 줄 '※', 패널 사이 화살표 기본 없음 (`arrows=True` 로 켬), 표 장 절반 넘으면 check 실패, `mpl_style()` 과 matplotlib 기본 색 그림 경고, 차트 색에서 주황 뺌

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
