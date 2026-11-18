# 변경 기록

판 번호 규칙 (대.중간.작은): 대는 회사 양식 교체나 기존 빌더 전면 재작성, 중간은 기존 빌더 코드를 고쳐야만 통과하는 변경(함수 이름이나 인자 변경 및 삭제, 관문 실패 규칙 추가), 작은은 그 밖의 전부(결함 수정, 자산, 주의 검사, 새 도구, 문서, 모양 조정, 영향 없는 새 부품)

## 1.7.6 (2026-11-18)
- **글꼴을 Pretendard 로 고정** (사용자 결정). `FONT`, 템플릿(`interx_template.pptx` 글꼴 이름), 렌더 글꼴 확인, 설치기(Pretendard 보통과 굵게 OTF 설치, Windows 등록 값 또는 Mac 글꼴 폴더), OKR 덱, 문서를 바꿈. `assets/fonts` 는 Pretendard Regular, Bold OTF 와 SIL OFL 1.1 원문(OFL-Pretendard.txt), Noto Sans KR 파일 제거 (Noto 10.4MB 제거, Pretendard 3.2MB 추가, 순감 약 7MB). 엔진의 글자 폭 표는 Pretendard 굵은 글자 실측(대문자 0.672, 소문자 0.540, 숫자 0.617, 한글 0.864 em)보다 크거나 같아 그대로 안전 (한글은 14% 여유)
- **목차 단차 정리**: 두 열일 때 모든 절을 같은 높이의 격자에 놓아 왼쪽과 오른쪽 열의 번호, 이름, 구분선이 같은 높이에 섬 (한 열일 때는 절마다 자기 줄 수만큼, 예전과 같음). 한 절은 최대 4줄 (3줄과 `외 N개`). 칸이 모자라면 줄 간격을 줄이고, 그래도 넘으면 한 열을 두 열로, 다시 줄 수를 줄이고, 그래도 넘으면 주의
- **구성도 읽힘**: `DocGrid.group` 과 `sub` 에 기능 색 `tone` (`TONES`: collect, send, manage, store, report, 머리띠와 옅은 바탕을 같은 계열로), `header` 만 있고 `area=False` 면 하위 묶음을 감싸지 않는다는 주의, 상자 `flow="source|sink|process"` 와 `role_box(flow=)` 를 적으면 저장 때 선이 빠진 상자(처리 상자에 이어진 선 2개 미만)를 알려 줌, `legend(kind="square")` 기능 색 범례, `area(fill=)`, `big_arrow` 는 그대로
- 설계 문서 기본기: `references/아키텍처_설계_순서.md` (사용자 흐름, 시퀀스, 배치, 흐름 방식, 내부 구조와 선택지), `references/구성도_읽힘_규칙.md`, 구성도 출발점 `scripts/example_diagram.py` (배치도, 기능과 배치 대응표, 처리 순서). design 스킬 0-3절
- 문장 검사 주의: 같은 장에 같은 그림 제목이 두 번 (구성도 장 `figure=` 와 `mermaid(title=)` 중복)
- **`check.py` 글꼴 검사 복구**: 정규식에 `\b` 대신 백스페이스 문자가 들어 있어 1.3.3 부터 어떤 글꼴도 걸리지 않았음. 이제 기준 글꼴 아닌 이름을 실패로 잡음
- 호환: 빌더 코드는 고치지 않아도 됨. 이전 판으로 만든 덱은 글꼴 이름이 Noto Sans KR 이라 빌더를 다시 돌리면 Pretendard 로 바뀜

## 1.7.5 (2026-11-18)
- 덱 용량 95% 감소: 템플릿(`assets/interx_template.pptx`)이 원본 양식의 임베드 글꼴 9개(Pretendard 7개, 맑은 고딕 2개, 약 22MB)를 담고 있었음. 글꼴 이름을 모두 Noto Sans KR 로 바꾼 뒤에는 쓰이지 않고 이름표만 잘못 붙어 모든 덱의 용량만 키움. 6장 덱 23.4MB 에서 1.2MB, 템플릿 31.6MB 에서 9.5MB (설치할 때 받는 용량도 줄어듦). 같은 덱을 이전과 이후 템플릿으로 만들어 렌더한 그림이 화소 단위로 같음을 확인. 회사 양식 디자인은 그대로
- `make_template.py` 가 양식 원본에서 템플릿을 만들 때 임베드 글꼴을 자동으로 뺌, `--drop-fonts 입력 출력` 으로 이미 만든 템플릿이나 덱에서도 뺄 수 있음
- 이전 판으로 만든 덱은 그대로이며, 다시 저장하면 작아짐 (덱은 템플릿에서 새로 만들어지므로 빌더를 다시 돌릴 때)

## 1.7.4 (2026-11-18)
- 관문이 PDF 를 자동으로 만들지 않음 (사용자가 직접 저장). 렌더는 바뀐 장만 다시 그리는 경로를 항상 씀
- 설계 문서 덱 출발점 `scripts/example_doc.py` (기능 상세, 화면 상세 예제, 관문 통과). `func_page`, `screen_page` 에 `source=` (패널 아래 출처 한 줄, 문서 덱은 한 장에만)
- 구성도 부품 `ixdiag.py`: `headline` (가운데 맞춤 헤드라인의 키워드 색 강조), `backdrop` (캡처 뒤 연한 쉐브론 배경)

## 1.7.3 (2026-11-18)
- xlsx 기본 서식: 머리를 연한 회색(D9D9D9) 바탕에 검정 굵은 글씨, 격자를 가는 회색(A6A6A6)으로 (공공기관 정의서식, 이전 남색 머리). 정의서 종류별 권장 열 구성 `skills/sheet/references/정의서_열_구성.md`
- 구성도 단계형 장 부품 `ixdiag.py`: `step_header` (남색 쉐브론 머리띠, 열 위치를 돌려줌), `summary_bar` (둥근 요약바), `speech_bubble` (말풍선)
- 문장 검사 주의: 같은 장에서 패널 제목과 표 제목이 같은 말
- 검수 반영: sheet 하위 표를 위한 `ref` 종류 (상위 ID 를 되풀이하고 첫 두 열 조합이 유일, sheet_check 도 같은 규칙), 문서 덱은 장마다 출처는 요구하지 않되 덱 어디에도 없으면 실패, 아이콘은 그림 장 판정에서 제외 (큰 그림만), 통과 기록은 덱이 없고 30일 넘게 안 쓴 것만 정리, 굵은 열의 행 높이 여유, 영문 뒤 한글 검사에서 조사 시작이면 정상, 블록 화살표와 약어 풀이도 겹침 검사 등록, 옛 이름 연결은 옛 설치가 있던 PC 에만, xlsx_view 색 값과 --rows 검증, `big_arrow` 가 색을 받아 `block_arrow` 가 재사용, Pillow 14 에서 없어질 getdata 대체
- 토큰: using 스킬의 유지보수 절(템플릿, 판 올린 뒤 확인, 글만 고치기, Google Slides)을 `references/유지보수.md` 로 분리 (덱 만들 때 읽는 글 약 10% 감소). 검토 에이전트 횟수를 덱에 맞춤 (사내 1회, 고객 보고와 납품 최대 2회, 20장 넘는 문서 덱은 대표 장만)

## 1.7.2 (2026-11-18)
- 이름: interx-deck 에서 ixdeck 으로 (플러그인 이름 및 호출 `ixdeck:using` 등, 설치 폴더 `skills/ixdeck`, 개인 설정 폴더 `~/.claude/ixdeck`는 첫 실행 때 옛 폴더에서 옮김). 옛 설치 경로 `skills/interx-deck` 는 새 폴더를 가리키는 연결로 잠시 유지 (다른 프로젝트 덱 빌더가 옛 경로로 엔진을 찾음). `example_build.py` 의 엔진 찾기는 두 이름 모두 봄. 저장소 이름은 그대로
- 새 갈래 `sheet` (xlsx 정의서): `ixsheet.py` (한 표 한 시트, 머리행 필터 및 틀 고정, 병합 및 빈 셀 금지, 숫자 오른쪽 및 코드 가운데 정렬, 상태 색, 줄 수로 계산한 행 높이, A4 가로 인쇄, 문서 정보 및 개정 이력 시트, 추적 매트릭스), `sheet_check.py` (규칙 검사), `xlsx_view.py` (Edge 로 시트별 그림, DRM 때문에 Excel 을 못 써서 근사 그림), `selfcheck_sheet.py`
- 설계 문서 부품 `ixdoc.py`: `Deck.func_page`, `Deck.screen_page`, `Col.field_table`, `limit_table`, `error_table`, `element_table`, `ref_table`, `flow`. 표는 행마다 줄 수에 맞춰 높이를 따로 정함 (`table(row_hs=)`). `selfcheck_doc.py`
- 통과 기록 `.gate.json` 을 덱 옆이 아니라 개인 설정 폴더 `gate/` 에 둠 (납품 폴더에 부산물이 섞이지 않음). 덱 옆의 옛 기록은 다음 관문 때 치움. 경로 처리는 `ixpaths.py`
- 구성도 문법 주의: 연결선 색은 의미 셋(데이터, 제어, 공통 버스)까지 허용 (빨강 끊김 별도), 넷부터 주의
- 표 캡션: 머리에 `단위` 열이 있는 표는 단위 필요 주의를 내지 않음
- 설치: Python 패키지에 openpyxl 추가. 옛 이름 자리 `skills/interx-deck` 는 `scripts` 연결만 둔 폴더 (plugin.json 이 없어 두 폴더가 같은 플러그인 이름으로 겹치지 않음), 연결을 못 만들면 옛 폴더를 되돌림. 옛 개인 설정 폴더는 새 폴더에 없는 것만 합침 (시작 훅이 새 폴더를 먼저 만들어도 안전). 제거는 옛 이름 자리도 지움. 설치 시험에 옛 판에서 올리는 경우 추가. 엔진 스크립트는 연결 경로로 불려도 실제 위치(realpath)에서 assets 를 찾음
- 문서 덱(`purpose="문서"`) 검사: 장마다 출처 줄 요구 없음, 처리 흐름이나 캡처가 있는 장은 표만 있는 장으로 세지 않음, 같은 출처 문구 반복은 주의. 오류 표의 서술체 화면 문구는 검사 예외 (`d.allow_text`)
- sheet 입력 검증: 이름, 머리, 단위, 그룹 라벨은 비거나 `=` 로 시작하면 거부, 사전 행의 여분 및 빠진 키, 날짜 범위, 제어 문자, 행 높이 409pt 상한, 시트 이름 대소문자 중복. sheet_check 는 수식 셀과 409.5pt 초과 행 오류, 큰 표 속도 개선
- 통과 기록에 덱 경로를 적고 없어진 덱의 기록은 치움
- 구성도 레퍼런스 문법 `skills/design/references/구성도_레퍼런스_문법.md` 와 부품 `ixdiag.py`: `role_box` (역할 색 칩), `block_arrow` (그라데이션 블록 화살표), `pill` (프로토콜 알약), `abbr_legend` (약어 풀이). `selfcheck_diag.py`
- 문장 검사 주의 추가: 표지 제목에 제품이나 과제 이름 없음, 제목 및 짧은 줄의 영문 뒤 한글 붙여 씀, 같은 출처 문구가 여러 장에 반복

## 1.7.1 (2026-11-16)
- 회사 방법론 A1 (전략 수립 절차서, 공정 분야) 요약 `skills/design/references/A1_방법론.md`: 7단계 절차, 5개 역할, X1 ~ X5 및 Yn 변수, DAS 칸, 품질 지수, Level 1 ~ 5, 과제 우선순위 가중치. design 스킬 0절, using 스킬 0-1절
- 새 부품: `Col.das_table` (데이터 수집 명세서), `Col.quality_table` (품질 지수 자동 계산, 95% 기준), `Col.priority_table` 및 `Col.priority_quad` (과제 우선순위 평가표와 4분면), `Col.levels` (수준 단계 띠)
- 문장 검사 주의: 방법론과 다른 용어 (목표변수, 타깃변수, 종속변수, 품질 점수, 성숙도 레벨, 설정 변수, 세팅 변수)
- 고침: 한글 줄 바꿈이 'Acq. Time', '[표 1. 제목]' 처럼 짧은 영문 약어나 번호 뒤 점에서 줄을 나누던 것

## 1.7.0 (2026-11-16)
- **기존 빌더 수정 필요 (중간 판 이유)**: 표와 그림 제목 필수. 표는 위 왼쪽 `[표 번호. 제목]`, 위 오른쪽 `(단위: ...)` (`Col.table`, `kpi_table`, `compare_table`, `collect_table` 의 `title=`, `unit=`, 단위 없는 맨 숫자 칸 둘 이상이면 단위 필수). 그림은 아래 가운데 `[그림 번호. 제목]` (`Col.chart`, `scatter`, `hist`, `image`, `gantt`, `mermaid` 의 `title=`, 구성도 장은 `d.page(..., figure=)`). 번호는 저장 때 덱 전체 연번. 빠지면 저장이 멈춤. `Col.chart(title=)` 는 위 파란 제목에서 아래 그림 제목으로 바뀜
- 두 장 연속 같은 패널 폭과 같은 블록 순서면 주의 (출력만)
- 구성도 문법: 세 계층 머리(원천, EdgePC, 서버)를 같은 네모 띠와 7.26pt 로 통일하고 색만 계층별로. 저장 때 주의(출력만): 연결선 색 셋 이상, 두 색 구성도 아이콘과 옛 색 아이콘 섞임. `legend(kind="line")` 점선 항목 `(색, 이름, "dash")`. 예제 구성도는 선 색 둘(남색 주 흐름, 회색)과 선 모양 범례
- Tabler 아이콘 7개 안쪽 청록 채움, firewall 은 Phosphor wall 로
- 그림 회귀 시험: `regress.py --baseline` 으로 기준 그림 저장, 이후 장마다 비교해 바뀐 장을 나란히 붙인 그림으로. 회귀 모드에서는 엄격 실패도 저장해 비교(엄격 실패로 표시)
- 관문: 지난 통과 때와 엔진 판이 다르면 그림이 바뀐 장을 알림, 사람 견본 수치 비교 단계(주의만)
- 사람 견본 라이브러리 `samples.py` (`assets/samples.json`, 배치 수치만), 사람 판별 시험 `blind.py`, 배포 뒤 새 폴더 설치 시험 `install_test.js`
- 고침: 설치본에서 `doctor` 의 엔진 시험이 package.json 을 찾다 실패하던 것 (설치본에는 없음)
- SKILL: 빌더를 쓰기 전 장 목록(절, 소제목, 요약 줄, 근거 자료) 확인, 구성도 문법 절

## 1.6.1 (2026-11-16)
- 구성도 아이콘 123개 추가: 남색 선과 청록 바탕 두 색 세트(Phosphor duotone, Tabler, MIT). 서버, PC, 노트북, 모바일, 사용자, 관리자, 클라우드, DB, 라우터, 방화벽, API, 게이트웨이, Edge, PLC, 센서 등
- 기술 로고 34개 추가(Simple Icons, CC0, 브랜드 색): docker, kubernetes, python, grafana, apachekafka, redis, nginx, mqtt, nodered, keycloak, pytorch, onnx, nvidia, siemens 등
- `scripts/build_icons.py`: 이름 목록으로 아이콘과 로고 PNG 및 `_catalog.png` 생성, 라이선스는 `third_party/icons`

## 1.6.0 (2026-11-16)
- 패널이 내용에 맞춰 줄어듦: `report_panels(fit=True)` 기본, 같은 줄 패널은 가장 긴 칸의 내용 높이로 같게 줄고 남는 장 공간은 비움. 빈 곳 실패 규칙 삭제(본문 높이의 45% 미만이면 주의만), 옛 높이는 `fit=False`, check.py 아래 여백 판정은 패널 장에서 주의로. `Col.kpi_table(last=)`
- 검사: 차트나 표가 있는 본문 장의 출처 줄 필수, 자리표시자(TODO, TBD, xxx, lorem, [insert, '텍스트를 입력', '예시 값', '요약 한 줄') 금지, 사내보고 및 제안 덱의 첫 본문 장 '요약', 발표 시간 대비 장 수 주의
- 쓰임새와 발표 시간: `Deck(purpose=, minutes=)`
- 발표자 노트: `d.page(..., notes=)`, 노트도 금지 표현 및 문장 검사 (서술체 제외)
- 도장 `stamp(s, "예시")`: 장 오른쪽 위 기운 빨간 테두리 (예시, 잠정, 검토용, 개념도 등)
- 저장이 숫자와 단위 사이 공백을 줄 바꿈 없는 공백으로 바꿈 (95.7 %, 24 시간)
- 관문: 장별 소제목 및 요약 줄 목록 `titles.md`, 렌더 실패 때 사유 출력
- 렌더: Noto Sans KR 이 없으면 종료 코드 3 으로 멈춤
- im-not-ai (MIT) 연동: `third_party/im-not-ai/` 규칙 파일, 어휘 검사(바꿀 어휘 주의), 검토 지시문에 A ~ J 분류, 갱신 `scripts/update_imnotai.py`
- 검토 지시문: 논지 흐름(titles.md), 장별 내용 및 디자인 점수와 덱 전체 일관성 점수, '그대로 사용 가능' 판정 기준 (결함 0건 및 모든 점수 4 이상)
- 새 장 유형: `Col.kpi_table`(머리행 고정, 출처 메모 필수), `Col.collect_table`(집계 방법 자동, 반자동, 수동), 평가항목 대응표 `d.criteria` 및 `d.page(covers=)`(저장 때 Thank you 앞에 자동 장, 답 없는 항목은 저장 실패), 일정 및 산출물 `d.deliverables`(gantt 이름과 대조)
- 차트: `y_title`, 단위 글자(y_title 또는 number_format) 없으면 저장 실패, 계열 4개 넘으면 주의
- 영업 덱 검사: 회사 소개 장 1장 이하, 첫 본문 장 금지, 본문 6 ~ 15장 주의, 영업 글쓰기(Raskin 순서). 문서 덱 밀도: 긴 줄 110자, 표 장 비율 70%
- 블록 글자 한도(번호 카드 16 및 40자, 이름표 줄 46자, 핵심 수치 이름 12자, 목록 행 30자, 규칙 글 30자) 넘으면 저장 실패
- 검토 지시문: 승리 주제(요약 장 핵심 문장이 절 요약 줄로 이어지는가)
- 관문: 바뀐 장만 렌더(장 XML 해시 `hashes.json`, 장 수나 틀이 바뀌면 전체, PDF 는 항상 전체), 통과 기록 `<이름>.gate.json` 과 `gate.py --verify 덱.pptx`
- `mermaid(s, 글, x, y, w, h)`: flowchart(graph TD, TB, LR) 와 sequenceDiagram 을 편집 가능한 도형과 붙은 연결선으로 (지원 밖 문법은 ValueError)
- `scripts/edit_deck.py`: 기존 덱의 지정 장 글만 바꿔 사본 저장, 나머지 파트는 바이트 그대로, 해시 대조 보고
- check.py: 90도 돌린 연결선의 바깥 상자를 돌린 채로 계산 (아래 여백 오탐, 장 밖 도형 판정에도)
- 배치 4, 검토 반영 (코드): `allow` 정규식이 숫자 및 단위 사이 NBSP 글에도 맞음, 블록 글자 한도를 칸 폭에 비례, 관문 장 순서를 presentation.xml 장 목록 순서로, 틀 해시에 관계 파일, media, presentation.xml, 렌더 스크립트 포함, 판 번호는 plugin.json 먼저, `gate.py --deck 덱.pptx`(build 없이 기존 덱에 검사, 렌더, 통과 기록), `edit_deck.py` 빈 옛 글 오류 및 끝에 `--deck` 안내
- 배치 4, 검토 반영 (검사): 출처 줄은 줄 머리 또는 괄호 안의 출처, 데이터 출처, 자료, Source 다음 ':' ('기준:' 은 출처 아님), 차트 단위에 통화 기호($, ₩, €, ¥), 일정 및 산출물 대조에서 마일스톤 제외, 날짜 표기 3종 이상 혼용 주의, 패널 머리와 소제목이 같은 말이면 저장이 출력, im-not-ai 갱신은 어휘 형식 확인 후에만 덮고 사전은 필드 빠진 항목을 건너뜀, 렌더 글꼴 확인에 Windows 등록 값과 fc-list 추가
- 배치 4, mermaid: 한 줄 또는 한 열 체인의 건너뛰거나 되돌아가는 선을 체인 옆 우회선으로(붙은 연결선, 바깥 층), 못 그리는 선은 이유를 밝힌 ValueError, 노드 글의 따옴표 제거와 `<br>` 줄 바꿈
- 배치 4, 엔진 디자인: `num_cards(numbered=False)`, `Col.compare_table`(바뀐 행의 개선 칸만 강조), `Col.collect_table(first=)`, 간트 막대 글자색(끝남 진한 회색, 진행과 예정 흰색), `rule_list` 이름표를 `tab_box` 와 같은 왼쪽 위 남색 네모로, 날짜 표기 YYYY.MM.DD 하나
- 예제 덱: 요약 장 먼저(사내보고), 지표 정의를 결과 앞에, 현행 및 개선 한 표, 용어와 날짜 통일, 구성도에 선 색 범례

## 1.5.1 (2026-11-13)
- 설치 화면: 카타나가 세로로 서서 모빌처럼 세로축으로 돌다 멈춤, 칼날을 얇고 날카롭게 (폭 3 픽셀)

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
