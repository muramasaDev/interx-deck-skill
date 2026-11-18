---
name: design
description: INTERX 회사 양식 설계 문서와 구성도 장 (시스템 구성도, 아키텍처도, 데이터 흐름도, 처리 순서 번호가 붙은 흐름, AAS 구조도, INTERX 설계 문서 양식, 손으로 그린 것처럼 도형과 붙은 연결선), 그리고 다른 양식 문서(사내 PM03 전사 문서: 요구사항정의서, 기능설계서, 표준데이터셋정의서, 데이터수집설계서 등)에 넣을 본문 장(parts_only). 공통 규칙은 ixdeck:using 을 먼저 읽는다. "구성도", "아키텍처도", "흐름도", "PM03 문서" 요청.
---

# ixdeck:design: 설계 문서와 구성도

공통 규칙은 `ixdeck:using`. 엔진 `${CLAUDE_PLUGIN_ROOT}/scripts/ixdeck.py`, 출발 파일 `scripts/example_build.py` 의 구성도 장.

## 0. 회사 방법론 먼저
- 설계 문서, 제안, 수행계획서는 `references/A1_방법론.md` (INTERX 제조 AI 및 데이터 분석 방법론 A1, 전략 수립 절차서) 를 먼저 읽고 장 순서와 용어를 맞춤
- 장 순서: 데이터 수집 방안 → 분석 방안 → 활용 계획 → 추진 계획 (방법론 4 ~ 7단계)
- 용어: 변수는 X1 작업변수, X2 환경변수, X3 셋팅변수, X4 실행변수1, X5 실행변수2, Yn 결과(목적)변수. 수준은 Level 1 ~ 5 (모니터링, 제어, 최적화, 자율운영). 품질은 품질 지수 (95% 이상). 문장 검사가 다른 말(목표변수, 품질 점수, 성숙도 레벨, 설정 변수)을 주의로 냄
- 부품

| 부품 | 쓰임 |
|---|---|
| `Col.das_table([(구분, Data, Data Type, Range, Acq. Time, Unit, Stage, Connected)])` | 데이터 수집 명세서 (DAS). 구분은 X1 ~ X5 또는 Yn, 빈 칸이면 오류 |
| `Col.quality_table({완전성: %, 유일성: %, 유효성: %, 일관성: %, 정확성: %})` | 품질 지수 자동 계산 (완전성 60%, 나머지 각 10%), 기준 95% 충족 또는 미달 표시 |
| `Col.priority_table([(과제, 기여도, 효과 크기, 시급성, 투자, 효과 발생, 관리)], cut=3.5)` | 과제 우선순위 평가 (점수 1 ~ 5, 중요도 0.4/0.3/0.3, 용이성 0.4/0.4/0.2), 구분 QW, SI, NH, LN |
| `Col.priority_quad(같은 rows)` | 4분면 (세로 중요도, 가로 용이성), 칸마다 과제 이름 |
| `Col.levels(지금 단계 0 ~ 3)` | 수준 단계 띠 (L1 ~ 2 모니터링, L3 제어, L4 최적화, L5 자율운영) |

- 자율 공장 목표 수치(생산성 200% 등)는 근거 자료가 없는 회사 목표라 인용하면 '회사 목표 수치' 로 밝힘

## 0-2. 기능 상세 및 화면 상세 장 (설계서, 표 중심)
- 기능 하나는 `d.func_page(소제목, 요약, flow=[...], limits=[...], io=[...], errors=[...], refs={...})`, 화면 하나는 `d.screen_page(소제목, 요약, image=, marks=, elements=, rules=, refs=)` (엔진 `ixdoc.py`)
- 있는 구역만 패널이 되고 패널 너비는 내용 양으로 정함. `func_page` 는 행이 적으면 행을 넉넉하게, 많으면 촘촘하게 자동 조정 (`screen_page` 는 고정). 오류 표의 화면 문구 중 서술체 검사에 걸리는 것은 자동으로 검사 예외, 그 밖의 화면 문구는 `d.allow_text("문구")`
- 한계 값은 글 목록이 아니라 `limit_table` (값 열은 남색 굵은 글씨), 입출력은 `field_table`, 오류는 `error_table` (코드는 3자리 숫자 또는 대문자 식별자), 화면 요소는 `element_table` (번호가 캡처의 번호 원과 같아야 하고 어긋나면 오류), 관련 번호는 `ref_table`, 처리 순서는 `flow`
- 표 칸이 줄 바꿈되면 그 행만 높아짐. 칸이 비면 오류 (없으면 `-`)
- 같은 틀이 연달아 나오지 않게: 구역이 다른 기능끼리 묶고, 흐름이 핵심인 기능은 `chevrons` 를 쓴 장으로 따로 만듦
- 구성도를 새로 그릴 때는 `references/구성도_레퍼런스_문법.md` 를 먼저 읽음 (역할 색 칩 `role_box`, 블록 화살표 `block_arrow`, 프로토콜 알약 `pill`, 약어 풀이 `abbr_legend`)

## 1. 구성도 장 (`d.page(..., doc=True)`)
```python
g = DocGrid(s, rows=5)                                       # 줄 5개, 도식은 오른쪽 끝까지
g.source("source", [("plc", "설비 1"), ("plc", "설비 2")], title="설비 PLC")   # 묶음보다 먼저
e = g.group("edge", "EdgePC", kind="edge", sub="설비 옆", header="INTERX.DA", cols=2, badge="badge_da")
cache, core = e.box(0, 0, "수집 계획 캐시"), e.box(3, 0, "수집 코어")         # (줄, 칸)
p = g.group("platform", "중앙 서버", rows=(0, 1), header="AAS Server", badge="badge_aas_server")
lake = g.group("platform", "연계 대상", rows=(3, 4), area=False).box(3, 0, "DataLake", kind="server", icon_="database")
link(s, core, lake, color=ACCENT, text="HTTP", step=(2, "전송"), cond="조건 문구")
```
- 열: `right=None`(기본, 플랫폼 6.71 ~ 12.56), `right="source"`(오른쪽 Data Source 열 11.65). 왼쪽 원천 열이 없으면 EdgePC 가 0.75 부터
- 같은 줄 번호의 상자는 묶음이 달라도 같은 높이(가로선이 곧게). 같은 묶음, 같은 칸 상자는 첫 선 전에 같은 폭
- 묶음 안 작은 묶음 `e.sub(칸, 줄0, 줄1, header=)`, 묶음 아래 띠 `g.bar`, 여러 대 `g.stack(..., n=1, dots=False)` (`DocGrid(bottom=6.3)`)
- **줄 규칙**: 묶음을 가로지르는 선의 줄은 그 사이 칸을 비운다. 곧은 선과 꺾인 선이 같은 변을 다투면 곧은 선을 먼저 긋는다
- **설명은 설명하는 자리 옆** (오른쪽 설명 열 없음): 흐름 조건은 `link(..., cond=)`, 상자 수치는 `note_by(s, 상자, "진한 회색 머리 글", ["메모 줄"])` (빨강은 실패 경로 선 색이라 쓰지 않음), 설정값은 상자 이름 괄호, 요약 줄과 같은 말은 쓰지 않음
- 강조 `focus(상자)`, 단계 범위 `group.zone(...)`, 끊긴 선 `link(..., color=RED, dash="dash", cut=True)`, 범례 `legend(...)`(색 종류만, 번호 범례 금지, `kind="line"` 이면 선 모양이고 `[(색, 이름, "dash")]` 로 점선. 선 범례는 한 줄로 `legend` 를 가로로 나란히)
- 선이 안 이어지는 쌍은 상자 옆 설명 글로 (예: 데이터 저장소에서 웹 화면으로 가는 선이 상자를 피해 못 가면 웹 화면 `note_by` 에 '수집 값은 데이터 저장소에서 조회')
- 같은 바탕 여러 장(평소 흐름, 끊겼을 때): 바탕을 함수 하나로 만들고 장마다 강조 색과 설명만 바꿈
- 남색 묶음도 `sub` 를 제목 띠 둘째 줄에. 칸이 비면 `cols=1`
- 기본 균형: 설계문서형 내용 중심 6.58, 오른 여백 0.77. 도식 아래는 꼬리말 선 위 0.30 in 까지

## 1-2. 구성도 문법 (공개 아키텍처 그림 공통 관찰)
- 영역: 왼쪽 현장(설비, EdgePC), 오른쪽 서버. 세 계층 머리는 같은 네모 띠와 같은 글자 크기, 색만 다름(설비 먹색, EdgePC 옅은 회청, 서버 남색)
- 선 색은 한 장에 의미 셋까지 (끊김 및 실패 경로 빨강은 따로): 데이터 흐름 남색 실선, 제어 흐름, 굵은 공통 버스. 종류 구분은 선 모양(실선, `dash`, `sysDot`)과 범례로. 넷부터 저장 때 주의
- 아이콘은 한 장에 한 세트: 두 색 구성도 세트(server_rack, pc, user, firewall, router, edge_box, plc_unit 등) 또는 옛 색 아이콘. 섞이면 저장 때 주의
- 아이콘 아래 이름(`icon(..., name=)`, `g.source`), 읽는 순서는 번호 원(`step=`, `step_by`)
- 그림 제목: `d.page(..., doc=True, figure="구성도 제목")` 이 저장 때 도식 아래 가운데에 `[그림 번호. 제목]` (도식 아래 0.3 in 비움, `DocGrid(bottom=)` 를 줄임)

## 2. 선과 번호 (자동)
- 선은 `link(s, a, b, color, text=선 이름, step=(번호, 설명), cond=조건)`. 마주 보는 변, 상자를 뚫지 않음, 곧은 선 > ㄱ 자 > ㄷ 자. 주 흐름부터 긋는다
- 여러 원천을 한 상자로 `fan(s, 상자, g.icons, ...)`, 상자에 붙는 번호 `step_by(s, 상자, n, 설명)`, 변을 꼭 정할 때만 `connect(...)`
- 번호와 선 이름은 저장 직전에: 선 바로 옆 > 선에서 0.6 in 안(테두리 안) > 그래도 없으면 테두리를 걸침
- 선 옆 번호 원은 0.15 in 7.26pt. 흐름 색: 기본 GRAY 1.5pt, 주 ACCENT(남색). 그 밖의 흐름은 GRAY 의 선 모양으로 (색 셋까지). 주황은 쓰지 않음. 아래 번호 범례 줄은 만들지 않는다

## 2-2. Mermaid 로 간단한 흐름과 순서 그리기
```python
mermaid(s, """graph TD
  A[수집 시작] --> B[PLC 읽기]
  B -- 읽기 실패 --> C[재시도]
  B --> D[저장]""", 1.0, 2.0, 11.3, 4.8, title="수집 처리 흐름")   # (s, 글, x, y, 폭, 높이) 인치, 아래 그림 제목
mermaid(s, """sequenceDiagram
  participant EG as INTERX.DA
  participant SV as AAS 서버
  EG->>SV: 값 전송
  SV-->>EG: 저장 완료""", 1.0, 2.0, 11.3, 4.8, title="값 전송 순서")
```
- 결과는 편집 가능한 상자와 붙은 연결선(`box`, `link`)이라 저장 검사와 check.py 를 그대로 통과해야 함. 반환값은 `{id: 도형}`
- flowchart: `graph TD`, `graph TB`, `graph LR`, 노드 `A[이름]`, 선 `A --> B`, `A -- 글 --> B`, 한 줄 연쇄 `A --> B --> C`. 층 배치(가장 긴 경로)이고 바깥 노드가 주어진 영역 끝에 닿음
- sequenceDiagram: `participant A as 이름`, `A->>B: 글`(실선), `A-->>B: 글`(점선 응답). 참가자 상자, 세로 점선, 메시지마다 양쪽 활성 막대 사이 연결선과 글
- 노드 글의 따옴표는 빼고 `<br/>` 와 `<br>` 는 줄 바꿈(참가자 이름도 같음, 메시지는 공백). 줄이 늘면 노드 높이가 같이 커짐
- 한 줄 또는 한 열 체인에서 건너뛰거나 되돌아가는 선은 체인 옆으로 우회(TD 는 오른쪽, LR 은 아래, 같은 쪽 변끼리 바깥으로 꺾인 붙은 연결선, 겹치는 선은 바깥 층). 체인이 아닌 그림의 그런 선이 상자를 피하지 못하면 어느 선인지 밝혀 ValueError
- 그 밖 문법(결정 마름모 `{}`, `---`, `-.->`, subgraph, loop, alt, Note, 자기 자신으로 가는 선)과 영역에 안 들어가는 크기는 ValueError. 복잡한 구성도는 DocGrid 로
- 선 위치를 손으로 고쳐야 하면 `d.save` 뒤 PowerPoint 에서 도형을 움직이면 연결선이 따라감

## 3. AAS 과제용 (AAS 를 다루지 않으면 건너뜀)
- 템플릿 14장 AAS 그림 `d.aas_shell(s, x, y, ["Identification", "HierarchicalStructures", "DataSource"], scale=0.55, focus_="DataSource")`, 펼친 내용은 `zoom_wedge` 와 `Col.chip_tree`
- 점선 다발(용어 표와 칩 잇기) `bundle(...)`, 배지 `badge_aas`, `badge_da`, `badge_ish` 등

## 4. 다른 양식 문서에 넣는 본문 장 (사내 PM03 전사 문서 등)
- `Deck(author, parts_only=True)`: 표지, 목차, Thank you 를 만들지 않고 본문 장만. 장은 `d.report_slide("1. 시스템 개요", "1.1 시스템 목적", lead=...)` 로 직접
- 관문을 통과한 장만 그 양식 장의 본문 자리로 한 배율로 줄여 옮긴다(과제 폴더의 옮기기 스크립트)
- 본문 바닥은 그 양식의 꼬리말 선 위 0.30 in (예: PM03 양식은 선 7.814 in 라 본문 아래 7.51 in)
- 양식 원본은 python-pptx 로 읽기만(회사 DRM)
- 화면 명세의 캡처 네모를 손으로 그리지 않는다. 그림을 놓은 뒤 `marks_on(s, pic, [(x0, y0, x1, y1, 번호), ...])`
- 표 칸이나 글상자를 빌더가 직접 채우면 `ko_breaks(글, 칸 폭, 글자 크기)` 로 줄을 나눠 넣는다(마침표 뒤 새 줄, 쉼표 뒤 끊기)
- 개정 이력의 주요 이슈는 핵심 한 문장(한 줄)
- 과제마다 다른 것(문서 번호, 판 규칙, 양식 원본 위치, IP 와 Lot 번호 가림 규칙)은 개인 설정 `projects/<과제>/README.md` 에 둔다
