"""AI 학습데이터 명세서 HTML 양식(작성용, 취합용) 채우기.
양식 파일의 form-data 블록만 바꿔 넣고, 양식 자신의 선택 목록으로 미리 점검한다 (양식에는 저장 때 검사가 없어 빠진 값이 그대로 제출됨).
사용: from ixspec import Spec, image   (skills/spec/SKILL.md, scripts/example_spec.py)
 sp = Spec(org="기관명", region="경기 안산")          # 취합용: Spec(lead=True, project={...})
 m = sp.model("모델명", company=..., company_region=..., overview=..., effect=...)
 t = m.target("철강", "", "열간압연 (판재/봉형강/특수강 등)", "품질검사")     # 돌려주는 값은 ID (T1)
 p = m.step(t, "상세 공정명"); s = m.device(p, "", "장치명", "센서"); d = m.dataset("timeseries", ...)
 sp.save(template_html, out_dir)  -> (파일 경로, 점검 결과)
"""
import ast
import base64
import html as _html
import io
import json
import os
import re

FORM_VERSION = "2.0"
BLOCK = re.compile(r'(<script type="application/json" id="form-data">)[\s\S]*?(</script>)')
NUM_INT, NUM_DEC = re.compile(r"[0-9]+"), re.compile(r"[0-9]+\.?[0-9]*|\.[0-9]+")   # 양식은 ASCII 숫자만 받음 (전각 숫자는 지워짐)
THOUSANDS = re.compile(r"[0-9]{1,3}(,[0-9]{3})+(\.[0-9]+)?")
PREFIX = {"target": "T", "step": "P", "equip": "E", "device": "S", "dataset": "D", "prep": "R"}
ID_RE = re.compile(r"^[TPESDR]\d+$")
BAD_CHARS = ("·", "—", "–")   # 사내 글쓰기 규칙: 자유 글에 가운뎃점과 긴 줄표 금지 (선택 목록 값은 원문 그대로라 예외)
# 양식을 읽지 못할 때만 쓰는 내장 목록 (양식 2.0 기준)
FALLBACK = {
    "AI_FUNCTIONS": ["생산제어", "품질검사", "설비예지보전", "이상감지", "에너지사용", "물류관리", "안전관리"],
    "WORK_MODES": ["연속 생산", "배치·Lot 단위", "개별 제품 단위", "사이클 반복"],
    "SPLIT_BY": ["작업 번호 기록 (Lot·배치·제품 번호)", "시작·종료 시각 기록", "설비 신호로 구분", "구분 정보 없음"],
    "DEVICE_KINDS": ["센서", "카메라", "검사기", "계측기", "PLC·제어기", "시스템(MES·ERP 등)", "직접 입력"],
    "LEARN_TYPES": ["지도학습", "비지도학습", "준지도학습", "강화학습"], "DATA_SOURCES": ["현장", "실험실", "혼합"],
    "USE": ["모델 입력", "예측 대상", "미사용"], "VALTYPE": ["숫자", "문자", "날짜·시각", "구분값"],
}
DS_KEYS = ["timeseries", "tabular", "image", "video", "text", "3d", "other"]
DS_REQ = {"timeseries": ["interval"], "tabular": ["rowUnit"], "image": ["target", "camera", "label"], "video": ["target", "label"],
          "text": ["docType"], "3d": ["target", "shape", "label"], "other": ["unitDesc"]}
DS_LABELED = ("image", "video", "3d")
KV_CHOICES = {"camera": ["일반 카메라 (컬러)", "일반 카메라 (흑백)", "열화상", "X-ray", "현미경", "기타"], "form": ["자유롭게 작성", "정해진 양식에 작성", "혼합"],
              "privacy": ["포함 안 됨", "포함 - 가림 처리함", "포함 - 처리 안 함"], "shape": ["점 데이터 (포인트클라우드)", "단면 영상 (CT 등)", "형상 모델 (메시)", "기타"]}
DEVICE_PROP_KEYS = {"센서": ["range", "accuracy", "unit", "interval"], "카메라": ["resolution", "fps", "fov", "mode"],
                    "검사기": ["method", "detect", "speed", "output"], "계측기": ["range", "accuracy", "unit", "interval"],
                    "PLC·제어기": ["signals", "protocol", "interval", "controller"],
                    "시스템(MES·ERP 등)": ["system", "items", "interval", "link"], "직접 입력": ["note"]}


# ---------------- 양식에서 선택 목록 읽기
def _literal(js, name):
    """const NAME = [...] 또는 {...} 의 글자 조각 (따옴표 안 괄호는 무시)"""
    m = re.search(r"const %s\s*=\s*" % name, js)
    if not m or js[m.end()] not in "[{":
        return None
    i, op = m.end(), js[m.end()]
    cl, depth, q, j = {"[": "]", "{": "}"}[op], 0, None, m.end()
    while j < len(js):
        c = js[j]
        if q:
            if c == "\\":
                j += 2
                continue
            if c == q:
                q = None
        elif c in "'\"`":
            q = c
        elif c == op:
            depth += 1
        elif c == cl:
            depth -= 1
            if depth == 0:
                return js[i:j + 1]
        j += 1
    return None


def load_enums(template):
    """양식 HTML 에서 선택 목록을 읽음. 못 읽은 것은 내장 목록으로 두고 enums['_fallback'] 에 이름을 적음"""
    js = open(template, encoding="utf8").read()
    e, miss = {}, []
    for name in list(FALLBACK) + ["INDUSTRY", "REGIONS", "AI_TASKS"]:
        lit = _literal(js, name)
        for parse in (json.loads, ast.literal_eval):   # JSON 모양이면 json, 작은따옴표 목록이면 파이썬 문법으로
            try:
                e[name] = parse(lit)
                break
            except Exception:
                continue
        else:
            e[name] = FALLBACK.get(name)
            miss.append(name)
    e["AI_TASK_ALL"] = {f"{g} > {ko}" for g, lst in (e.get("AI_TASKS") or {}).items() for ko, _ in lst}
    m = re.search(r"const FORM_VERSION = '([^']*)'", js)
    e["version"] = m.group(1) if m else None
    e["mode"] = (re.search(r"const APP_MODE = '(\w+)'", js) or [None, None])[1]
    e["_fallback"] = miss
    return e


# ---------------- 값 다듬기
def _num(v, integer=False, what="숫자"):
    """숫자 칸은 문자열이며 쉼표와 단위 문자가 없어야 함 (있으면 양식이 값을 지움)"""
    if v is None or v == "":
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    if THOUSANDS.fullmatch(s):
        s = s.replace(",", "")
    if not (NUM_INT if integer else NUM_DEC).fullmatch(s):
        raise ValueError(f"{what} '{v}' 은 {'정수' if integer else '숫자'}만 쓸 수 있음 (단위 문자 금지, 용량은 MB 로 환산)")
    return s


def rich(x):
    """서식 있는 글: 문자열 또는 줄 목록. '- ' 로 시작하는 줄은 글머리표. 이미 허용 태그(b, ul, ol, li, div, br)가 있으면 그대로"""
    if x is None:
        return ""
    if isinstance(x, str) and re.search(r"</?(b|ul|ol|li|br|div|p)\b", x, re.I):   # 허용 태그 글: 태그가 아닌 < 와 & 는 글자로
        x = re.sub(r"<(?!/?(b|ul|ol|li|br|div|p)\b)", "&lt;", x, flags=re.I)
        return re.sub(r"&(?!(amp|lt|gt|quot|#[0-9]+);)", "&amp;", x)
    items, out, bul = (x if isinstance(x, (list, tuple)) else str(x).split("\n")), [], []

    def flush():
        if bul:
            out.append("<ul>" + "".join(f"<li>{b}</li>" for b in bul) + "</ul>")
            bul.clear()
    for t in items:
        t = str(t).strip()
        if not t:
            continue
        if t.startswith("- "):
            bul.append(_html.escape(t[2:]))
        else:
            flush()
            out.append("<div>" + _html.escape(t) + "</div>")
    flush()
    return "".join(out)


def image(path, caption="", max_side=1280, quality=85):
    """이미지 파일 -> 양식 이미지 객체 (긴 변 1280px, JPEG 로 줄임. 양식 화면이 넣을 때와 같은 방식)"""
    from PIL import Image, ImageOps
    im = ImageOps.exif_transpose(Image.open(path))   # 휴대폰 사진의 방향 정보를 반영 (브라우저가 보는 방향과 같게)
    w, h = im.size
    if im.mode in ("RGBA", "LA", "P"):
        bg = Image.new("RGB", im.size, (255, 255, 255))
        im = im.convert("RGBA")
        bg.paste(im, mask=im.split()[-1])
        im = bg
    im = im.convert("RGB")
    im.thumbnail((max_side, max_side))
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=quality)
    return {"src": "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode(), "name": os.path.basename(path),
            "origW": w, "origH": h, "caption": caption}


def _clean_name(x):
    return re.sub(r"[\\/:*?\"<>|\s]+", "_", str(x or "").strip())[:40]


def var(name, desc="", unit="", device="", vtype="", use=""):
    """변수 목록 한 줄 (시계열 변수, 테이블형 항목). device 는 수집 장치 ID (S#), use 는 모델 입력, 예측 대상, 미사용"""
    return {"name": name, "desc": desc, "unit": unit, "device": device, "vtype": vtype, "use": use}


def label(cls, definition="", count=""):
    """라벨 종류 한 줄 (이미지, 영상, 3D). count 는 건수 (숫자)"""
    return {"cls": cls, "def": definition, "count": _num(count, what="라벨 건수")}


def grid(cols, rows):
    """원시 데이터 예시 표: 머리글 한 줄과 3~5행 (엑셀에서 복사한 모양)"""
    n = len(cols)
    cell = lambda v: "" if v is None or (isinstance(v, float) and v != v) else str(v)
    if any(len(r) > n for r in rows):
        raise ValueError(f"grid 행에 열 수({n})보다 많은 칸이 있음")
    return {"cols": [cell(c) for c in cols], "rows": [[cell(v) for v in list(r) + [""] * (n - len(r))] for r in rows]}


# ---------------- 모델
class Model:
    def __init__(self, spec, name, company, company_region, overview, effect, process_note="", images=(), org="", org_region=""):
        self.spec = spec
        self.d = {"id": 0, "name": name, "org": org, "orgRegion": org_region,
                  "problem": {"company": company, "companyRegion": company_region, "overview": rich(overview), "images": list(images),
                              "targets": [], "processNote": rich(process_note), "effect": rich(effect)},
                  "site": {"steps": [], "equipment": [], "devices": [], "images": []}, "datasets": [],
                  "prep": {"steps": [], "episodes": [], "results": {}, "example": {"grid": grid(["컬럼 1", "컬럼 2", "컬럼 3", "컬럼 4"], [["", "", "", ""]] * 2), "images": []}},
                  "ready": {"io": [], "ioInit": True, "joinBy": "", "split": {"train": "", "val": "", "test": ""}},
                  "ai": {"algo": "", "learnType": "", "learnTypeText": "", "task": "", "taskText": "", "metrics": [], "params": [],
                         "framework": "", "gpu": "", "gpuCount": "", "hw": "", "devEnv": "", "fileFormat": ""},
                  "etc": {"note": ""}}

    def _id(self, kind, lst):
        return PREFIX[kind] + str(1 + max([int(x["id"][1:]) for x in lst] or [0]))

    def target(self, industry, sub, process, ai_function):
        """대상 업종 및 공정 (최대 5). 목록 밖 값은 직접 입력으로 저장 (custom). 돌려주는 값: T#"""
        lst = self.d["problem"]["targets"]
        t = {"id": self._id("target", lst), "industry": industry, "subIndustry": sub, "process": process, "aiFunction": ai_function}
        ind = (self.spec.enums or {}).get("INDUSTRY") or {}
        if ind and (industry not in ind or sub not in ind.get(industry, {}) or process not in ind.get(industry, {}).get(sub, [])):
            t["custom"] = True
        lst.append(t)
        return t["id"]

    def step(self, target, name, prev="", next="", work_mode="", cycle_time="", split_by="", split_key="", desc=""):
        """상세 공정. 돌려주는 값: P#"""
        lst = self.d["site"]["steps"]
        x = {"id": self._id("step", lst), "target": target, "name": name, "prev": prev, "next": next, "workMode": work_mode,
             "cycleTime": cycle_time, "splitBy": split_by, "splitKey": split_key, "desc": desc}
        lst.append(x)
        return x["id"]

    def equipment(self, step, name, code="", maker="", spec="", desc=""):
        """설비 (해당 시). 돌려주는 값: E#"""
        lst = self.d["site"]["equipment"]
        x = {"id": self._id("equip", lst), "step": step, "name": name, "code": code, "maker": maker, "spec": spec, "desc": desc}
        lst.append(x)
        return x["id"]

    def device(self, step, equip, name, kind, kind_text="", item="", loc="", product="", maker="", props=None, desc=""):
        """수집 장치. equip 은 E# 또는 ''(설비에 설치 안 됨, MES 등). props 는 종류별 특성 (센서: range, accuracy, unit, interval 등). 돌려주는 값: S#"""
        lst = self.d["site"]["devices"]
        x = {"id": self._id("device", lst), "step": step, "equip": equip, "name": name, "kind": kind, "kindText": kind_text, "item": item,
             "loc": loc, "product": product, "maker": maker, "props": {k: v for k, v in (props or {}).items() if v not in ("", None)},
             "desc": desc}
        lst.append(x)
        return x["id"]

    def dataset(self, type, name, period, source, count, size_mb, devices, kv=None, rows=(), fmt="", desc="", source_text="",
                raw_grid=None, raw_images=(), raw_text=""):
        """원시 데이터. type: timeseries, tabular, image, video, text, 3d, other. count 는 건수(정수), size_mb 는 MB.
        kv 는 유형별 상세 (시계열 interval, 테이블형 rowUnit, 이미지 target, camera, label 등). rows 는 var() 또는 label() 목록. 돌려주는 값: D#"""
        lst = self.d["datasets"]
        x = {"id": self._id("dataset", lst), "type": type, "name": name, "period": period, "source": source, "sourceText": source_text,
             "count": _num(count, True, f"{name} 건수"), "sizeMB": _num(size_mb, what=f"{name} 용량"), "format": fmt, "devices": list(devices),
             "desc": desc, "kv": dict(kv or {}), "rows": list(rows),
             "raw": {"grid": raw_grid or grid(["컬럼 1", "컬럼 2", "컬럼 3", "컬럼 4"], [["", "", "", ""]] * 2), "images": list(raw_images), "text": raw_text}}
        lst.append(x)
        return x["id"]

    def prep_step(self, targets, method, detail=""):
        """전처리 단계 (추가한 순서가 처리 순서). 돌려주는 값: R#"""
        lst = self.d["prep"]["steps"]
        x = {"id": self._id("prep", lst), "targets": list(targets), "method": method, "detail": detail}
        lst.append(x)
        return x["id"]

    def episode(self, targets, range_):
        """학습 사용 구간 (안정화 전, 작업 종료 후처럼 학습에서 뺀 구간이 있을 때)"""
        self.d["prep"]["episodes"].append({"targets": list(targets), "range": range_})

    def prep_result(self, ds, count, size_mb):
        """전처리 후 건수와 용량(MB). 전처리 단계가 있는 데이터만"""
        self.d["prep"]["results"][ds] = {"count": _num(count, True, f"{ds} 전처리 후 건수"), "sizeMB": _num(size_mb, what=f"{ds} 전처리 후 용량")}

    def prep_example(self, grid_=None, images=()):
        self.d["prep"]["example"] = {"grid": grid_ or self.d["prep"]["example"]["grid"], "images": list(images)}

    def io(self, role, ds, item, shape="", ds_name=""):
        """모델 입력 및 출력 한 줄. role 은 입력 또는 출력, ds 는 D# (원시 데이터에 없는 값이면 ds='' 와 ds_name)"""
        self.d["ready"]["io"].append({"role": role, "ds": ds, "dsName": ds_name, "item": item, "shape": shape})

    def join_by(self, text):
        self.d["ready"]["joinBy"] = text

    def split(self, train, val="", test=""):
        """학습, 검증, 테스트 비율 (%)"""
        self.d["ready"]["split"] = {"train": _num(train, what="학습 비율"), "val": _num(val, what="검증 비율"), "test": _num(test, what="테스트 비율")}

    def ai(self, algo, learn_type, task, learn_type_text="", task_text="", framework="", gpu="", gpu_count="", hw="", dev_env="", file_format=""):
        """알고리즘, 학습 유형, 학습 방식(AI Task: '이미지 > 영역 분할 (시맨틱)' 처럼 '유형 > 이름' 또는 '기타')과 개발 환경"""
        self.d["ai"].update({"algo": algo, "learnType": learn_type, "learnTypeText": learn_type_text, "task": task, "taskText": task_text,
                             "framework": framework, "gpu": gpu, "gpuCount": _num(gpu_count, True, "GPU 개수"), "hw": hw, "devEnv": dev_env,
                             "fileFormat": file_format})

    def metric(self, name, target="", result="", note=""):
        self.d["ai"]["metrics"].append({"name": name, "target": target, "result": result, "note": note})

    def param(self, name, value="", note=""):
        self.d["ai"]["params"].append({"name": name, "value": value, "note": note})

    def etc(self, note):
        self.d["etc"]["note"] = rich(note)


# ---------------- 전체
class Spec:
    def __init__(self, org="", region="", lead=False, project=None, template=None):
        """lead=False 작성용(기관 하나), lead=True 취합용(project: title, number, leadOrg, author, contact, date)"""
        self.org, self.region, self.lead = org, region, lead
        self.project = dict({"title": "", "number": "", "leadOrg": "", "author": "", "contact": "", "date": ""}, **(project or {}))
        self.models, self.enums = [], load_enums(template) if template else None

    def model(self, name, company, company_region, overview, effect, process_note="", images=(), org="", org_region=""):
        """AI 모델 하나. 취합용에서는 담당 기관 org 와 소재 지역 org_region 을 모델마다 줌 (작성용은 Spec 의 기관이 들어감)"""
        m = Model(self, name, company, company_region, overview, effect, process_note, images, org, org_region)
        self.models.append(m)
        return m

    def state(self):
        out = []
        for i, m in enumerate(self.models, 1):
            d = json.loads(json.dumps(m.d))
            d["id"] = i
            d["org"] = d["org"] or self.org
            d["orgRegion"] = self.region if not self.lead else d["orgRegion"]
            out.append(d)
        return {"formVersion": FORM_VERSION, "savedAt": None, "modifiedAt": None, "role": "lead" if self.lead else "writer",
                "writerOrg": "" if self.lead else self.org, "writerRegion": "" if self.lead else self.region,
                "project": self.project if self.lead else {}, "models": out}

    def file_name(self):
        c = _clean_name
        if self.lead:
            return "_".join(x for x in ("AI학습데이터명세서_취합", c(self.project["number"]), c(self.project["leadOrg"])) if x) + ".html"
        return "_".join(x for x in ("AI학습데이터명세서_작성", c(self.org), c(self.project.get("number", ""))) if x) + ".html"

    def check(self):
        return check_state(self.state(), self.enums or {}, self.lead)

    def save(self, template, out_dir, allow_errors=False):
        """양식 파일에 채워 저장. 점검에서 오류가 있으면 저장하지 않음 (allow_errors 로 막을 수 있음). (경로, 점검 결과) 반환"""
        self.enums = self.enums or load_enums(template)
        problems = self.check()
        errs = [p for p in problems if p[0] == "오류"]
        if errs and not allow_errors:
            raise ValueError("명세서 점검 오류 %d건, 저장하지 않음:\n" % len(errs) + "\n".join(f"  {a} {b}: {c}" for a, b, c in errs))
        src = open(template, encoding="utf8").read()
        if len(BLOCK.findall(src)) != 1:
            raise ValueError("양식에서 form-data 블록(<script type=\"application/json\" id=\"form-data\">)을 하나만 찾아야 함")
        try:
            js = json.dumps(self.state(), ensure_ascii=False, separators=(",", ":"), allow_nan=False).replace("<", "\\u003c")
        except ValueError as ex:
            raise ValueError("값에 NaN 이나 무한대가 있음 (양식이 문서 전체를 못 읽음): " + str(ex)) from None
        os.makedirs(out_dir, exist_ok=True)
        path = os.path.join(out_dir, self.file_name())
        with open(path, "w", encoding="utf8", newline="") as f:
            f.write(BLOCK.sub(lambda m: m.group(1) + js + m.group(2), src, count=1))
        return path, problems


# ---------------- 점검 (양식 스크립트에는 없는 검사. 3.2 별표와 추가 단추 조건 기준)
def _plain(v):
    return re.sub(r"<[^>]+>", "", str(v or "")).strip()


def check_state(st, e, lead=False):
    """(수준, 위치, 내용) 목록. 수준 '오류' 는 제출 전 고쳐야 하고 '주의' 는 확인"""
    out = []
    add = lambda lv, where, msg: out.append((lv, where, msg))
    reg, ind = e.get("REGIONS") or [], e.get("INDUSTRY") or {}
    fb = e.get("_fallback")
    for name in fb or ():
        if name in ("INDUSTRY", "REGIONS", "AI_TASKS"):
            add("오류", "양식", f"양식에서 {name} 목록을 읽지 못해 그 항목을 점검할 수 없음: 양식 판이 다름")
        else:
            add("주의", "양식", f"양식에서 읽지 못한 목록 {name}: 내장 목록(양식 2.0)으로 점검함")
    if e.get("mode") and e["mode"] != ("lead" if lead else "writer"):
        add("오류", "양식", f"양식은 '{e['mode']}' 용인데 {'취합용(lead=True)' if lead else '작성용'} 으로 채움: 저장 때 role 과 기관 값이 바뀜")
    if e.get("version") and e["version"] != FORM_VERSION:
        add("주의", "양식", f"양식 판 {e['version']} 은 이 도구가 아는 {FORM_VERSION} 과 다름: 선택 목록은 양식에서 읽었으나 키 구조를 확인")
    if not lead:
        for k, lab in (("writerOrg", "기관명"), ("writerRegion", "작성 기관 소재 지역")):
            if not st.get(k):
                add("오류", k, f"{lab} 이 비었음")
        if reg and st.get("writerRegion") and st["writerRegion"] not in reg:
            add("오류", "writerRegion", f"'{st['writerRegion']}' 은 양식의 지역 목록에 없음")
    else:
        for k, lab in (("title", "과제명"), ("number", "과제번호"), ("leadOrg", "주관기관"), ("author", "작성자")):
            if not st["project"].get(k):
                add("오류", f"project.{k}", f"{lab} 이 비었음 (취합용 필수)")
    if not st["models"]:
        add("오류", "models", "AI 모델이 하나도 없음")
    seen_models = set()
    for mi, m in enumerate(st["models"], 1):
        w = f"모델{mi}({m.get('name', '')[:12]})"
        key_ = (m.get("org") or st.get("writerOrg", ""), m.get("name", ""))
        if key_ in seen_models:
            add("오류", w, "같은 기관과 같은 모델명이 두 번 (취합 때 앞의 것이 새 것으로 교체되어 사라짐)")
        seen_models.add(key_)
        if not lead and m.get("org") and m["org"] != st.get("writerOrg"):
            add("주의", w, f"담당 기관 '{m['org']}' 이 기관명과 다름: 작성용은 저장 때 기관명으로 덮어씀")
        pr, site = m["problem"], m["site"]
        for lab, v in (("모델명", m["name"]), ("대상 기업", pr["company"]), ("대상 기업 소재 지역", pr["companyRegion"]), ("문제 개요", _plain(pr["overview"])),
                       ("기대 효과", _plain(pr["effect"]))):
            if not v:
                add("오류", w, f"{lab} 이 비었음")
        if reg and pr["companyRegion"] and pr["companyRegion"] not in reg:
            add("오류", w, f"대상 기업 소재 지역 '{pr['companyRegion']}' 은 목록에 없음")
        if lead and not m.get("org"):
            add("오류", w, "담당 기관(org) 이 비었음")
        tg = pr["targets"]
        if not 1 <= len(tg) <= 5:
            add("오류", w, f"대상 업종 및 공정은 1~5개 (지금 {len(tg)}개)")
        keys = set()
        for t in tg:
            if not (t["industry"] and t["process"] and t["aiFunction"]):
                add("오류", f"{w} {t['id']}", "업종, 공정, AI 기능이 모두 있어야 함")
            if t["aiFunction"] and e.get("AI_FUNCTIONS") and t["aiFunction"] not in e["AI_FUNCTIONS"]:
                add("오류", f"{w} {t['id']}", f"AI 기능 '{t['aiFunction']}' 은 목록 {e['AI_FUNCTIONS']} 중 하나")
            if ind:
                if t["industry"] not in ind:
                    add("주의", f"{w} {t['id']}", f"업종 '{t['industry']}' 은 양식 목록에 없음: 직접 입력(custom)으로 저장됨, 분류가 맞는지 확인 (목록: {', '.join(ind)})")
                elif t["subIndustry"] not in ind[t["industry"]]:
                    add("오류", f"{w} {t['id']}", f"세부 구분 '{t['subIndustry']}' 이 업종 '{t['industry']}' 에 없음 (세부 구분 없는 업종은 '')")
                elif t["process"] not in ind[t["industry"]][t["subIndustry"]]:
                    add("주의", f"{w} {t['id']}", f"공정 '{t['process']}' 은 목록에 없음: 직접 입력(custom)으로 저장됨")
            k = (t["industry"], t["subIndustry"], t["process"], t["aiFunction"])
            if k in keys:
                add("오류", f"{w} {t['id']}", "같은 업종, 세부 구분, 공정, AI 기능 조합이 중복")
            keys.add(k)
        tids = {t["id"] for t in tg}
        steps = {s["id"]: s for s in site["steps"]}
        eqs = {x["id"]: x for x in site["equipment"]}
        devs = {x["id"]: x for x in site["devices"]}
        dss = {x["id"]: x for x in m["datasets"]}
        for lst, lab in ((tg, "T"), (site["steps"], "P"), (site["equipment"], "E"), (site["devices"], "S"), (m["datasets"], "D"), (m["prep"]["steps"], "R")):
            seen = set()
            for x in lst:
                if not ID_RE.match(x["id"]) or not x["id"].startswith(lab) or x["id"] in seen:
                    add("오류", w, f"ID '{x['id']}' 형식 또는 중복 문제 ({lab}+정수, 유일)")
                seen.add(x["id"])
        if not steps:
            add("오류", w, "상세 공정이 하나도 없음")
        names = set()
        for s in steps.values():
            if s["target"] not in tids:
                add("오류", f"{w} {s['id']}", f"대상 공정 '{s['target']}' 이 대상 업종 및 공정(T#)에 없음")
            if (s["target"], s["name"].replace(" ", "")) in names:
                add("오류", f"{w} {s['id']}", "같은 대상 공정에서 상세 공정명이 중복")
            names.add((s["target"], s["name"].replace(" ", "")))
            for key, lst in (("workMode", e.get("WORK_MODES")), ("splitBy", e.get("SPLIT_BY"))):
                if s[key] and lst and s[key] not in lst:
                    add("오류", f"{w} {s['id']}", f"{key} '{s[key]}' 은 목록 {lst} 중 하나")
            if s["splitKey"] and not (e.get("SPLIT_BY") and s["splitBy"] == e["SPLIT_BY"][0]):
                add("주의", f"{w} {s['id']}", "번호 이름(splitKey)은 작업 구분 방법이 '작업 번호 기록' 일 때만 의미가 있음")
        for x in eqs.values():
            if x["step"] not in steps or not x["name"]:
                add("오류", f"{w} {x['id']}", "설비는 상세 공정(P#)과 설비명이 있어야 함")
        if not devs:
            add("오류", w, "수집 장치가 하나도 없음")
        for x in devs.values():
            if x["step"] not in steps or not x["name"] or not x["kind"]:
                add("오류", f"{w} {x['id']}", "수집 장치는 상세 공정(P#), 장치명, 종류가 있어야 함")
            if x["equip"] and x["equip"] not in eqs:
                add("오류", f"{w} {x['id']}", f"설치 설비 '{x['equip']}' 이 설비(E#)에 없음 (설비에 설치 안 됐으면 '')")
            if e.get("DEVICE_KINDS") and x["kind"] and x["kind"] not in e["DEVICE_KINDS"]:
                add("오류", f"{w} {x['id']}", f"종류 '{x['kind']}' 은 목록 {e['DEVICE_KINDS']} 중 하나")
            if x["kind"] == "직접 입력" and not x["kindText"]:
                add("오류", f"{w} {x['id']}", "종류가 직접 입력이면 종류 입력(kind_text)이 필요")
            bad = [k for k in x["props"] if k not in DEVICE_PROP_KEYS.get(x["kind"], [])]
            if bad:
                add("주의", f"{w} {x['id']}", f"종류 '{x['kind']}' 에 없는 특성 키 {bad}: 화면에 안 보임 (가능: {DEVICE_PROP_KEYS.get(x['kind'])})")
        if not dss:
            add("오류", w, "원시 데이터가 하나도 없음")
        for d in dss.values():
            dw = f"{w} {d['id']}({d['name'][:10]})"
            if d["type"] not in DS_KEYS:
                add("오류", dw, f"데이터 유형 '{d['type']}' 은 {DS_KEYS} 중 하나 (아니면 양식이 데이터를 지움)")
                continue
            for lab, v in (("데이터명", d["name"]), ("수집 기간", d["period"]), ("데이터 출처", d["source"]), ("건수", d["count"]), ("용량(MB)", d["sizeMB"])):
                if not v:
                    add("오류", dw, f"{lab} 이 비었음")
            src_ok = (e.get("DATA_SOURCES") or []) + ["기타"]
            if d["source"] and d["source"] not in src_ok:
                add("오류", dw, f"데이터 출처 '{d['source']}' 은 {src_ok} 중 하나")
            if d["source"] == "기타" and not d["sourceText"]:
                add("오류", dw, "출처가 기타이면 직접 입력(source_text)이 필요")
            if not d["devices"]:
                add("오류", dw, "수집 장치를 하나 이상 고름")
            for s in d["devices"]:
                if s not in devs:
                    add("오류", dw, f"수집 장치 '{s}' 가 없음 (양식이 조용히 지움)")
            for k in DS_REQ[d["type"]]:
                if not d["kv"].get(k):
                    add("오류", dw, f"{d['type']} 유형의 필수 상세 '{k}' 가 비었음")
            if d["type"] in DS_LABELED and d["kv"].get("label") not in ("있음", "없음"):
                add("오류", dw, "라벨은 '있음' 또는 '없음'")
            if d["type"] in DS_LABELED and d["kv"].get("label") == "있음" and not (d["kv"].get("labelDesc") and any(r.get("cls") for r in d["rows"])):
                add("주의", dw, "라벨이 있으면 라벨 설명(labelDesc)과 라벨 종류 표(label())를 채움")
            want = "name" if d["type"] in ("timeseries", "tabular") else "cls" if d["type"] in DS_LABELED else None
            for r in d["rows"]:
                if want is None or (want == "name" and "cls" in r) or (want == "cls" and "name" in r):
                    add("오류", dw, f"{d['type']} 유형의 표 줄 종류가 안 맞음 (시계열, 테이블형은 var(), 이미지, 영상, 3D 는 label(), 그 외는 표 없음): 화면에 안 보임")
                    break
            for k, ok in KV_CHOICES.items():
                if d["kv"].get(k) and d["kv"][k] not in ok:
                    add("오류", dw, f"{k} '{d['kv'][k]}' 은 {ok} 중 하나")
            for r in d["rows"]:
                if not (r.get("name") or r.get("cls")):
                    add("오류", dw, "변수 목록 또는 라벨 종류에 이름이 빈 줄이 있음")
                if r.get("device") and r["device"] not in devs:
                    add("오류", dw, f"변수 '{r.get('name')}' 의 수집 장치 '{r['device']}' 가 없음")
                if r.get("use") and e.get("USE") and r["use"] not in e["USE"]:
                    add("오류", dw, f"사용 '{r['use']}' 은 {e['USE']} 중 하나")
                if r.get("vtype") and e.get("VALTYPE") and r["vtype"] not in e["VALTYPE"]:
                    add("오류", dw, f"값 종류 '{r['vtype']}' 은 {e['VALTYPE']} 중 하나")
            raw = d["raw"]
            has_grid = any(any(c.strip() for c in row) for row in raw["grid"]["rows"])
            if d["type"] in ("timeseries", "tabular") and not has_grid:
                add("주의", dw, "원시 데이터 예시(머리글 포함 3~5행)가 비었음")
            if d["type"] in ("image", "video") and not raw["images"]:
                add("주의", dw, "대표 이미지가 없음 (이미지는 라벨 유형별 1장 이상, 영상은 대표 장면 2~3장)")
            if d["type"] == "text" and not raw["text"]:
                add("주의", dw, "원문 일부(raw_text)가 비었음")
            if d["type"] in ("3d", "other") and not (raw["images"] or raw["text"]):
                add("주의", dw, "이미지 또는 텍스트 예시가 한 건도 없음")
            if (d["type"] in ("timeseries", "tabular")) and not d["rows"]:
                add("주의", dw, "변수 목록(항목 목록)이 비었음")
        pr_ = m["prep"]
        for s in pr_["steps"]:
            if not s["method"] or not s["targets"] or any(t not in dss for t in s["targets"]):
                add("오류", f"{w} {s['id']}", "전처리 단계는 처리 방법과 존재하는 대상 데이터(D#)가 필요 (없는 D# 는 양식이 단계째 지움)")
        for ep in pr_["episodes"]:
            if not ep["range"] or not ep["targets"] or any(t not in dss for t in ep["targets"]):
                add("오류", w, "학습 사용 구간은 사용 구간과 존재하는 대상 데이터(D#)가 필요")
        used = {t for s in pr_["steps"] for t in s["targets"]}
        for k in pr_["results"]:
            if k not in dss:
                add("오류", w, f"전처리 결과 '{k}' 의 데이터가 없음")
            elif k not in used:
                add("주의", w, f"{k} 는 전처리 단계가 없는데 결과가 있음")
        for k in used:
            if k in dss and k not in pr_["results"]:
                add("주의", w, f"{k} 전처리 후 건수와 용량이 없음 (prep_result)")
        rd = m["ready"]
        if not rd["io"]:
            add("오류", w, "모델 입력 및 출력이 하나도 없음 (io)")
        for r in rd["io"]:
            if r["role"] not in ("입력", "출력") or not r["item"]:
                add("오류", w, "입출력은 구분(입력, 출력)과 항목이 필요")
            if r["ds"] and r["ds"] not in dss:
                add("오류", w, f"입출력의 데이터 '{r['ds']}' 가 없음 (양식이 그 줄을 지움)")
            if not r["ds"] and not r["dsName"]:
                add("오류", w, "원시 데이터에 없는 입출력은 ds_name 이 필요")
        if rd["io"] and not any(r["role"] == "출력" for r in rd["io"]):
            add("주의", w, "출력(예측 대상) 줄이 없음")
        sp = rd["split"]
        if not sp["train"] or not sp["test"]:
            add("오류", w, "학습 비율과 테스트 비율이 필요")
        elif abs(sum(float(sp[k] or 0) for k in ("train", "val", "test")) - 100) > 1e-6:
            add("오류", w, f"학습, 검증, 테스트 비율 합이 100 이 아님 ({sp['train']}+{sp['val'] or 0}+{sp['test']})")
        ai = m["ai"]
        for lab, v in (("알고리즘", ai["algo"]), ("학습 유형", ai["learnType"]), ("학습 방식(AI Task)", ai["task"])):
            if not v:
                add("오류", w, f"{lab} 이 비었음")
        lt = (e.get("LEARN_TYPES") or []) + ["기타"]
        if ai["learnType"] and e.get("LEARN_TYPES") and ai["learnType"] not in lt:
            add("오류", w, f"학습 유형 '{ai['learnType']}' 은 {lt} 중 하나")
        if ai["learnType"] == "기타" and not ai["learnTypeText"]:
            add("오류", w, "학습 유형이 기타이면 직접 입력이 필요")
        if ai["task"] == "기타":
            if not ai["taskText"]:
                add("오류", w, "학습 방식이 기타이면 직접 입력(task_text)이 필요")
        elif ai["task"] and e.get("AI_TASK_ALL") and ai["task"] not in e["AI_TASK_ALL"]:
            add("오류", w, f"학습 방식 '{ai['task']}' 이 목록에 없음: '유형 > 이름' 형식 (예: 이미지 > 영역 분할 (시맨틱)) 또는 '기타'")
        if not [x for x in ai["metrics"] if x.get("name")]:
            add("오류", w, "성능 지표가 하나도 없음 (metric)")
        # 자유 글에 사내 글쓰기 규칙 (가운뎃점, 긴 줄표)
        for where, v in _free_texts(m):
            if any(c in v for c in BAD_CHARS):
                add("주의", f"{w} {where}", f"글에 가운뎃점이나 긴 줄표가 있음: {v[:30]}...")
        for lst in (pr["images"], site["images"], pr_["example"]["images"], *[d["raw"]["images"] for d in dss.values()]):
            for im in lst:
                if not re.fullmatch(r"data:image/(jpeg|png|gif|webp);base64,[A-Za-z0-9+/=]+", im.get("src", "")):
                    add("오류", w, "이미지 src 가 허용 형식(data:image/jpeg|png|gif|webp;base64,...)이 아님 (양식이 조용히 지움)")
                if len(im.get("src", "")) > 1_500_000:
                    add("주의", w, f"이미지 하나가 {len(im['src']) // 1000}KB: 줄여서 넣기 (image() 는 긴 변 1280px)")
    return out


def _free_texts(m):
    p, s = m["problem"], m["site"]
    yield "문제 개요", _plain(p["overview"])
    yield "기대 효과", _plain(p["effect"])
    for lab, v in (("공정 정보", p["processNote"]), ("기타", m["etc"]["note"])):
        yield lab, _plain(v)
    for st in s["steps"]:
        yield st["id"], " ".join((st["name"], st["desc"], st["prev"], st["next"]))
    for d in m["datasets"]:
        yield d["id"], " ".join((d["name"], d["desc"], *[str(v) for v in d["kv"].values()]))
        for r in d["rows"]:
            yield d["id"], " ".join(str(v) for k, v in r.items() if k in ("name", "desc", "def"))
    for r in m["prep"]["steps"]:
        yield r["id"], r["method"] + " " + r["detail"]
    for r in m["ready"]["io"]:
        yield "입출력", r["item"] + " " + r["shape"]
