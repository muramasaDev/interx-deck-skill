---
name: design
description: INTERX 회사 양식 설계 문서와 구성도 장 (시스템 구성도, 아키텍처도, 데이터 흐름도, 처리 순서 번호가 붙은 흐름, AAS 구조도, INTERX 설계 문서 양식, 손으로 그린 것처럼 도형과 붙은 연결선), 그리고 다른 양식 문서(사내 PM03 전사 문서: 요구사항정의서, 기능설계서, 표준데이터셋정의서, 데이터수집설계서 등)에 넣을 본문 장(parts_only). 공통 규칙은 interx-deck:using 을 먼저 읽는다. "구성도", "아키텍처도", "흐름도", "PM03 문서" 요청.
---

# interx-deck:design: 설계 문서와 구성도

공통 규칙은 `interx-deck:using`. 엔진 `${CLAUDE_PLUGIN_ROOT}/scripts/ixdeck.py`, 출발 파일 `scripts/example_build.py` 의 구성도 장.

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
- **설명은 설명하는 자리 옆** (오른쪽 설명 열 없음): 흐름 조건은 `link(..., cond=)`, 상자 수치는 `note_by(s, 상자, "빨간 한 줄", ["메모 줄"])`, 설정값은 상자 이름 괄호, 요약 줄과 같은 말은 쓰지 않음
- 강조 `focus(상자)`, 단계 범위 `group.zone(...)`, 끊긴 선 `link(..., color=RED, dash="dash", cut=True)`, 범례 `legend(...)`(색 종류만, 번호 범례 금지)
- 같은 바탕 여러 장(평소 흐름, 끊겼을 때): 바탕을 함수 하나로 만들고 장마다 강조 색과 설명만 바꿈
- 남색 묶음도 `sub` 를 제목 띠 둘째 줄에. 칸이 비면 `cols=1`
- 기본 균형: 설계문서형 내용 중심 6.58, 오른 여백 0.77. 도식 아래는 꼬리말 선 위 0.30 in 까지

## 2. 선과 번호 (자동)
- 선은 `link(s, a, b, color, text=선 이름, step=(번호, 설명), cond=조건)`. 마주 보는 변, 상자를 뚫지 않음, 곧은 선 > ㄱ 자 > ㄷ 자. 주 흐름부터 긋는다
- 여러 원천을 한 상자로 `fan(s, 상자, g.icons, ...)`, 상자에 붙는 번호 `step_by(s, 상자, n, 설명)`, 변을 꼭 정할 때만 `connect(...)`
- 번호와 선 이름은 저장 직전에: 선 바로 옆 > 선에서 0.6 in 안(테두리 안) > 그래도 없으면 테두리를 걸침
- 선 옆 번호 원은 0.15 in 7.26pt. 흐름 색: 기본 GRAY 1.5pt, 주 ACCENT(남색), 둘째 TEAL, 의미 연결 GREEN 점선. 주황은 쓰지 않음. 아래 번호 범례 줄은 만들지 않는다

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
