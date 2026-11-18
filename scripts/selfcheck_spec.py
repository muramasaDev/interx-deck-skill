"""ixspec 자체 시험: 가짜 양식(form-data 블록과 선택 목록 상수만)으로 채우기, 점검, 저장, 이름 규칙을 확인.
실제 양식이 있으면(INTERX_SPEC_TEMPLATE) 브라우저 저장본 비교까지 돌림. 사용: python -P selfcheck_spec.py"""
import json
import os
import re
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from ixspec import BLOCK, Spec, _num, grid, load_enums, rich, var  # noqa: E402

STUB = """<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>
<script type="application/json" id="form-data">{}</script>
<script>
const FORM_VERSION = '2.0';
const APP_MODE = 'writer';
const INDUSTRY = {"철강": {"": ["열간압연 (판재/봉형강/특수강 등)", "냉간압연·후처리"]}, "반도체": {"01_전공정_1": ["식각"]}};
const REGIONS = ["경북 포항", "경기 안산"];
const AI_FUNCTIONS = ['생산제어','품질검사'];
const AI_TASKS={"이미지": [["영역 분할 (시맨틱)", "Semantic segmentation"]]};
</script></body></html>"""


def build(lead=False):
    sp = Spec(org="A사", region="경북 포항", lead=lead, template=tpl,
              project={"title": "시험 과제", "number": "EX-2026-0001", "leadOrg": "주관사", "author": "시험"} if lead else None)
    m = sp.model("표면 결함 검출 모델", company="가나다제철(주)", company_region="경북 포항", overview=["문단 <1>", "- 항목"], effect="- 개선",
                 org="A사")
    t = m.target("철강", "", "냉간압연·후처리", "품질검사")
    p = m.step(t, "표면 검사")
    s = m.device(p, "", "카메라", "카메라")
    d = m.dataset("timeseries", "속도", "2024.04 ~ 2025.03", "현장", "12,568", 420.0, [s], kv={"interval": "1초"},
                  rows=[var("v", "속도", "m/min", s, "숫자", "모델 입력")], raw_grid=grid(["a", "b"], [[1, 2]]))
    m.prep_step([d], "정규화")
    m.prep_result(d, 12568, 420)
    m.io("입력", d, "v")
    m.io("출력", d, "불량 여부")
    m.split(80, 10, 10)
    m.ai("U-Net", "지도학습", "이미지 > 영역 분할 (시맨틱)")
    m.metric("Dice")
    return sp, m


tmp = tempfile.mkdtemp()
tpl = os.path.join(tmp, "양식.html")
open(tpl, "w", encoding="utf8").write(STUB)

e = load_enums(tpl)
assert e["REGIONS"] == ["경북 포항", "경기 안산"] and e["AI_FUNCTIONS"] == ["생산제어", "품질검사"] and "이미지 > 영역 분할 (시맨틱)" in e["AI_TASK_ALL"], e
assert "WORK_MODES" in e["_fallback"], "읽지 못한 목록은 내장 목록으로 두고 알림"

# 숫자: 쉼표는 제거, 단위 문자는 ValueError
assert _num("12,568", True) == "12568" and _num(420.0) == "420" and _num("") == ""
for bad in ("1.5GB", "abc"):
    try:
        _num(bad)
        raise AssertionError(bad)
    except ValueError:
        pass
assert rich(["문단 <1>", "- 항목"]) == "<div>문단 &lt;1&gt;</div><ul><li>항목</li></ul>"

# 정상: 오류 0, 저장하면 form-data 블록만 바뀜, 이름 규칙
sp, m = build()
path, problems = sp.save(tpl, tmp)
assert not [p for p in problems if p[0] == "오류"], problems
out = open(path, encoding="utf8").read()
assert os.path.basename(path) == "AI학습데이터명세서_작성_A사.html"
assert len(BLOCK.findall(out)) == 1 and out.replace(BLOCK.search(out).group(0), "") == STUB.replace(BLOCK.search(STUB).group(0), "")
JS = re.compile(r'<script type="application/json" id="form-data">([\s\S]*?)</script>')   # BLOCK 은 여는 태그와 닫는 태그만 잡음
data = json.loads(JS.search(out).group(1))
assert data["role"] == "writer" and data["writerOrg"] == "A사" and data["models"][0]["id"] == 1 and data["models"][0]["datasets"][0]["count"] == "12568"
raw = JS.search(out).group(1)
assert "\\u003cdiv>" in raw and "<div>" not in raw   # 글 안의 < 는 < 로 이스케이프 (</script> 로 블록이 끊기지 않게)

# 목록 밖 업종은 custom 으로 저장되고 주의
sp2, m2 = build()
m2.target("전자부품", "", "도금", "이상감지")
assert m2.d["problem"]["targets"][1].get("custom") is True
assert any("업종 '전자부품'" in x[2] for x in sp2.check() if x[0] == "주의")

# 오류 잡기: 없는 D#, 없는 학습 방식, 필수 빈 값, 5개 초과, 가운뎃점
sp3, m3 = build()
m3.prep_step(["D9"], "없는 데이터")
m3.io("입력", "D8", "x")
m3.d["ai"]["task"] = "이미지 > 없는 방식"
m3.d["problem"]["company"] = ""
m3.d["problem"]["effect"] = rich("검사 · 비용")
msgs = " | ".join(f"{a} {c}" for a, b, c in sp3.check())
for key in ("존재하는 대상 데이터", "입출력의 데이터 'D8'", "학습 방식 '이미지 > 없는 방식'", "대상 기업 이 비었음", "가운뎃점"):
    assert key in msgs, (key, msgs)
try:
    sp3.save(tpl, tmp)
    raise AssertionError("오류가 있으면 저장하지 않음")
except ValueError as ex:
    assert "점검 오류" in str(ex)

# 리뷰에서 나온 경우: 전각 숫자, 쉼표 소수점, NaN, 중복 모델, 허용 밖 태그, 이미지 줄바꿈, 줄 종류 불일치
for bad in ("１２５６８", "1,5"):
    try:
        _num(bad, True)
        raise AssertionError(bad)
    except ValueError:
        pass
assert rich("T<Tmax 이면 경보<br>끝 & 끝") == "T&lt;Tmax 이면 경보<br>끝 &amp; 끝"
sp4, m4 = build()
m4.metric("RMSE", result=float("nan"))
try:
    sp4.save(tpl, tmp)
    raise AssertionError("NaN")
except ValueError as ex:
    assert "NaN" in str(ex), ex
sp5, m5 = build()
m5.d["datasets"][0]["rows"].append({"cls": "라벨"})
m5.d["problem"]["images"].append({"src": "data:image/jpeg;base64,AAAA\n", "name": "a", "origW": 1, "origH": 1, "caption": ""})
sp5.models.append(sp5.models[0])   # 같은 기관과 모델명
msgs = " | ".join(c for _, _, c in sp5.check())
for key in ("표 줄 종류가 안 맞음", "허용 형식", "두 번"):
    assert key in msgs, (key, msgs)

# 취합용: 과제 정보 필수, 파일 이름
spl, _ = build(lead=True)
assert spl.file_name() == "AI학습데이터명세서_취합_EX-2026-0001_주관사.html"
spl.project["number"] = ""
assert any("과제번호" in c for _, _, c in spl.check())

# 실제 양식이 있으면 브라우저에서 저장본 비교 (오래 걸려 선택)
ENV = {**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"}
real = os.environ.get("INTERX_SPEC_TEMPLATE")
if real and os.path.exists(real):
    r = subprocess.run([sys.executable, "-P", os.path.join(HERE, "example_spec.py"), "--template", real, "--out", tmp], capture_output=True, env=ENV)
    made = [l for l in r.stdout.decode("utf8", "replace").splitlines() if l.endswith(".html")]
    assert made, r.stderr.decode("utf8", "replace")[-300:]
    v = subprocess.run([sys.executable, "-P", os.path.join(HERE, "spec_verify.py"), made[-1]], capture_output=True, env=ENV)
    assert v.returncode == 0, v.stdout.decode("utf8", "replace")[-400:]
    print("실제 양식 브라우저 비교 통과")
if not real:
    print("(브라우저 비교는 INTERX_SPEC_TEMPLATE 가 있을 때만 돌림)")
print("selfcheck_spec OK")
