"""AI 학습데이터 명세서 출발점 (spec_build.py 로 복사해 내용만 교체).
- 자리표시 데이터만 있음 (기관, 설비, 수치는 형식을 보이기 위한 값). 실제 값은 과제 자료에서만 가져오고 모르는 값은 비워 확인 목록으로 둔다
- 양식 HTML 은 사용자가 받은 파일을 쓴다 (--template 경로 또는 INTERX_SPEC_TEMPLATE). 이 저장소에는 양식 파일이 없음
사용: python -P example_spec.py --template 기관별작성양식.html --out 폴더   ->   python -P spec_verify.py 만든파일.html"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ixspec import Spec, grid, label, var  # noqa: E402,F401

ap = argparse.ArgumentParser()
ap.add_argument("--template", default=os.environ.get("INTERX_SPEC_TEMPLATE"))
ap.add_argument("--out", default=".")
a = ap.parse_args()
if not a.template:
    sys.exit("양식 HTML 경로가 필요함: --template 또는 INTERX_SPEC_TEMPLATE")

sp = Spec(org="A사", region="경북 포항", template=a.template)
m = sp.model("표면 결함 검출 모델", company="가나다제철(주) 포항공장", company_region="경북 포항",
             overview=["냉연 강판 표면의 결함을 카메라 이미지로 찾아 불량 여부를 판정하는 모델",
                       "- 입력: 표면 이미지, 출력: 결함 종류와 위치"],
             effect=["- 검사 누락률 감소", "- 육안 검사 공수 절감"])
t = m.target("철강", "", "냉간압연·후처리", "품질검사")
p = m.step(t, "표면 검사", work_mode="연속 생산")
eq = m.equipment(p, "표면 검사 라인", maker="자리표시")
cam = m.device(p, eq, "검사 카메라", "카메라", item="강판 표면 이미지", props={"resolution": "1600 x 256", "fps": "30"})
plc = m.device(p, "", "공정 제어 PLC", "PLC·제어기", item="라인 속도", props={"interval": "1초"})
img = m.dataset("image", "강판 표면 이미지", "2024.04 ~ 2025.03", "현장", 12568, 1650, [cam], fmt="png",
                kv={"target": "냉연 강판 상면", "camera": "일반 카메라 (흑백)", "resolution": "1600 x 256", "label": "없음"})
ts = m.dataset("timeseries", "라인 속도 데이터", "2024.04 ~ 2025.03", "현장", 3000000, 420, [plc], fmt="csv",
               kv={"interval": "1초"},
               rows=[var("Line_Speed", "라인 속도", "m/min", plc, "숫자", "모델 입력"), var("Time", "기록 시각", "", plc, "날짜·시각", "미사용")],
               raw_grid=grid(["Time", "Line_Speed"], [["2024-04-01 00:00:00", "812.3"], ["2024-04-01 00:00:01", "812.1"], ["2024-04-01 00:00:02", "811.9"]]))
m.prep_step([img], "밝기 정규화", "조명 변화에 따른 밝기 차이를 보정")
m.prep_result(img, 12568, 1650)
m.io("입력", img, "표면 이미지")
m.io("입력", ts, "Line_Speed")
m.io("출력", img, "결함 종류, 결함 위치")
m.split(80, 10, 10)
m.ai("U-Net", "지도학습", "이미지 > 영역 분할 (시맨틱)", framework="PyTorch 2.2", gpu="NVIDIA RTX 4090", gpu_count=2)
m.metric("Dice 계수", target="0.88 이상")
m.param("epoch", "100")
path, problems = sp.save(a.template, a.out)
for lv, where, msg in problems:
    print(lv, where, msg)
print(path)
