"""
시트에 있는거 파이썬으로 옮겨옴 (vba 너무 여러개라 정신없어서 여기서 만들고 이식할 얘정)
- date_output 시트의 룩업 테이블(A1:E13)은 하드코딩
- 사용자 입력(Y, M)만 외부에서 받음
"""

# ------------------------------------------------------------------
# date_output 시트 룩업 테이블 하드코딩 (A2:E13, base_date = "Y " + M 매칭)
#   base_date(C열) -> (type_12 T0 raw(B열), type_3 T0 raw(E열))
# ------------------------------------------------------------------
LOOKUP_TABLE = {
    "Y 1":  ("Y-1 Q3", "Y-1 9"),
    "Y 2":  ("Y-1 Q3", "Y-1 9"),
    "Y 3":  ("Y-1 Q3", "Y-1 9"),
    "Y 4":  ("Y-1 Q4", "Y-1 12"),
    "Y 5":  ("Y-1 Q4", "Y-1 12"),
    "Y 6":  ("Y-1 Q4", "Y-1 12"),
    "Y 7":  ("Y Q1",   "Y-1 15"),
    "Y 8":  ("Y Q1",   "Y-1 15"),
    "Y 9":  ("Y Q1",   "Y-1 15"),
    "Y 10": ("Y Q2",   "Y 6"),
    "Y 11": ("Y Q2",   "Y 6"),
    "Y 12": ("Y Q2",   "Y 6"),
}


def q_to_num(q: str) -> str:
    """QToNum: Q1~Q4 -> 3/6/9/12, 그 외는 그대로"""
    mapping = {"Q1": "3", "Q2": "6", "Q3": "9", "Q4": "12"}
    return mapping.get(q, q)


def parse_value(raw: str, base_year: int):
    """ParseValue: 'Y 6' / 'Y-1 Q3' / 'Y+1 3' 같은 문자열을 (year, val) 로 분해"""
    if not raw:
        return base_year, "0"

    parts = raw.split(" ")
    if len(parts) < 2:
        return base_year, "0"

    prefix, val = parts[0], parts[1]

    if prefix == "Y":
        year = base_year
    elif prefix == "Y-1":
        year = base_year - 1
    elif prefix == "Y+1":
        year = base_year + 1
    else:
        year = base_year

    return year, val


def prev_session(in_year: int, in_val: int):
    """PrevSession: 3/6/9/12/15 결산 체계에서 직전 회차 계산"""
    mapping = {
        3:  (in_year - 1, "12"),
        6:  (in_year,     "3"),
        9:  (in_year,     "6"),
        12: (in_year,     "9"),
        15: (in_year,     "12"),
    }
    return mapping.get(in_val, (in_year, str(in_val)))


def fill_date_info(input_year: int, input_month: int):
    """
    VBA Sub FillDateInfo 와 동일한 계산. 헷갈리면 시트 참고하기
    반환값: dict (B22, B23, C22, C23, E22, E23, F22, F23)
    """
    if not (1 <= input_month <= 12):
        raise ValueError("M 값은 1~12 사이여야 합니다.")

    lookup_key = f"Y {input_month}"
    type12_t0_raw, type3_t0_raw = LOOKUP_TABLE[lookup_key]

    # T0 계산
    type12_t0_year, type12_t0_val = parse_value(type12_t0_raw, input_year)
    type3_t0_year, type3_t0_val = parse_value(type3_t0_raw, input_year)

    type12_t0_val = q_to_num(type12_t0_val)
    type3_t0_val = q_to_num(type3_t0_val)

    # T-1 계산
    type12_t1_year, type12_t1_val = prev_session(type12_t0_year, int(type12_t0_val))
    type3_t1_year, type3_t1_val = prev_session(type3_t0_year, int(type3_t0_val))

    return {
        "type12_T0_year": type12_t0_year,
        "type12_T0_val": int(type12_t0_val),
        "type12_T1_year": type12_t1_year,
        "type12_T1_val": int(type12_t1_val),
        "type3_T0_year": type3_t0_year,
        "type3_T0_val": int(type3_t0_val),
        "type3_T1_year": type3_t1_year,
        "type3_T1_val": int(type3_t1_val),
    }


if __name__ == "__main__":
    # ---- 여기에 입력 ----
    input_year = int(input("Y (연도)를 입력하세요: ").strip())
    input_month = int(input("M (월, 1~12)을 입력하세요: ").strip())
    # ---------------------------

    result = fill_date_info(input_year, input_month)

    print("\n=== date_output 결과 ===")
    print(f"B22 (type_12 acct_T0 Y)  : {result['type12_T0_year']}")
    print(f"B23 (type_12 acct_T0 Q)  : {result['type12_T0_val']}")
    print(f"C22 (type_12 acct_T-1 Y) : {result['type12_T1_year']}")
    print(f"C23 (type_12 acct_T-1 Q) : {result['type12_T1_val']}")
    print(f"E22 (type_3 acct_T0 Y)   : {result['type3_T0_year']}")
    print(f"E23 (type_3 acct_T0 Q)   : {result['type3_T0_val']}")
    print(f"F22 (type_3 acct_T-1 Y)  : {result['type3_T1_year']}")
    print(f"F23 (type_3 acct_T-1 Q)  : {result['type3_T1_val']}")