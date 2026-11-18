"""새 정의서 출발점 (build_sheet.py 로 복사, 내용만 교체)
- 자리표시 데이터만 있음 (시스템, 항목, 수치 모두 형식을 보이기 위한 값)
사용: python -P example_sheet.py 출력.xlsx   (엔진 sheet_check 가 통과해야 저장됨)"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.realpath(__file__)))  # 복사했다면 이 줄을 엔진 scripts 폴더 경로로 바꾼다
from ixsheet import Book, Field  # noqa: E402

LONG = "샘플 설명 문장이다. 이 열은 자동 줄 바꿈이 되고 행 높이는 글자 수와 열 너비로 계산해 지정된다. 여러 값은 한 셀에 넣지 않는다."


def build_example(path):
    b = Book(title="API 정의서 (예시)",
             info={"대상 시스템": "샘플 시스템", "제품 판": "v1.0.0", "문서 판": "v0.1.0", "기준일": "2026-11-17",
                   "범위": "항목 A 부터 항목 F 까지", "근거 자료": "샘플 자료 1"},
             history=[("v0.1.0", "2026-11-17", "신규", "최초 작성")])
    state = ["완료", "진행", "계획", "미정"]
    api = [(f"API-{i:03d}", f"항목 {chr(64 + i)}", ["GET", "POST", "PUT", "DELETE"][i % 4], 10 * i, 12.5 * i,
            LONG if i % 3 == 0 else "샘플 설명", f"2026-11-{i + 9:02d}", state[i % 4]) for i in range(1, 7)]
    b.table("API 목록", [Field("API ID", "id"), Field("이름", "text", width=24), Field("Method", "code"),
                        Field("호출 주기", "num", unit="초"), Field("평균 응답 시간", "num", decimals=1, unit="ms"),
                        Field("설명", "text", width=50), Field("기준일", "date"), Field("상태", "status")],
            api, group={"요청": ("Method", "호출 주기"), "응답": ("평균 응답 시간", "설명")})
    fn = [(f"FN-{i:03d}", f"기능 {chr(64 + i)}", "샘플 설명", ["상", "중", "하"][i % 3], state[i % 4]) for i in range(1, 5)]
    b.table("기능 목록", [Field("기능 ID", "id"), Field("기능명", "text", width=24), Field("설명", "text", width=50),
                        Field("우선순위", "code"), Field("상태", "status")], fn)
    b.table("처리 현황", [Field("항목 ID", "id"), Field("항목명", "text", width=24),
                        Field("처리 건수", "num"), Field("오류 건수", "num")],
            [(f"IT-{i:03d}", f"항목 {chr(64 + i)}", 1000 * i * i, 3 * i) for i in range(1, 5)], unit="건")
    b.matrix("요구사항 추적", [r[0] for r in fn], ["REQ-001", "REQ-002", "REQ-003"],
             [("FN-001", "REQ-001"), ("FN-002", "REQ-001"), ("FN-002", "REQ-002"), ("FN-004", "REQ-003")],
             corner="기능 ID")
    b.save(path)


if __name__ == "__main__":
    build_example(sys.argv[1] if len(sys.argv) > 1 else "example_sheet.xlsx")
