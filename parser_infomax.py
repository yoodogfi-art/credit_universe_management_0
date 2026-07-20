# -*- coding: utf-8 -*-
"""
VBA Sub parser_info() 의 파이썬(openpyxl) 구현

원본 로직 요약
--------------
1. "info_4214" 시트의 B열(종목명, raw)을 기준으로 "처음 등장하는 행"만 남겨서
   종목명 -> code_fn(AI열), sector(AL열), 원본행 매핑을 만든다.
2. 쓰레기명(IsTrashName) / 출력 제외 대상(IsExcludedFromOutput) 패턴에 해당하는
   종목명은 제외하고, 나머지를 오름차순 정렬한다.
3. 정렬된 종목 목록을 순회하며 bond(회사채) / CP(기업어음) / corp(무보증사채 등)
   등급, 일자, 등급 일치 여부(T/F), 등급전망(긍정/안정/부정) 개수를 계산한다.
4. 결과를 "parsed_infomax.json" 파일로 저장한다.

4214 -> "일반"으로 검색하기 때문에 공사 등이 섞여 들어갈 수도 있다.
컬럼 인덱스는 VBA arrInfo(row, N) 에서 쓰인 숫자를 그대로 사용했습니다.
(A=1, B=2, C=3, ... Z=26, AA=27, ... AP=42)

수정 내역
--------
- (버그) 파일 최상단에 있던 `openpyxl.read_xlsx("info_4214.xlsx")` 삭제.
  openpyxl 에는 read_xlsx 라는 메서드가 없어서 import 시점에 바로 AttributeError 가 남.
- (버그) 계산 로직(2~4단계)이 함수 밖 모듈 레벨로 빠져나가 있었고, 함수 내부에는
  `for name in only_a: ... ...` 라는 아무 일도 안 하는 껍데기 루프만 남아 있었음.
  -> parser_info() 는 아무것도 계산하지 않고 None 을 반환했고,
     모듈 레벨 코드는 함수 지역변수(code_a, arr_info, row_map_a, sector_a)를
     참조하다가 import 시점에 NameError 로 죽었음.
  전체 로직을 parser_info() 함수 내부로 다시 합쳐서 정리함.
- (변경) 결과 저장 방식을 엑셀 out_info 시트 -> "parsed_infomax.json" 로 변경.
"""

import json
import openpyxl


# ------------------------------------------------------------------
# 등급전망(outlook) 판정에 쓰는 한글 키워드
# ------------------------------------------------------------------
POS_STR = "긍정적"        # 긍정적
NEU_STR = "안정적"        # 안정적
NEG_STR = "부정적"        # 부정적
NEG_STR2 = "부정적검토"    # 부정적검토


def tf_check(base_rating: str, r1: str, r2: str, r3: str) -> str:
    """
    TFCheck 대응 함수 [신용등급 parser]

    기준등급(base_rating)과 r1/r2/r3 중 값이 있는 것들을 모아서,
    하나 이상 존재하고 전부 같으면 'T', 하나라도 다르면 'F', 전부 비어있으면 ''.
    """
    ratings = [r for r in (base_rating, r1, r2, r3) if r != ""]

    if len(ratings) == 0:
        return ""

    first = ratings[0]
    for r in ratings:
        if r != first:
            return "F"

    return "T"


def count_outlook(outlook: str, counts: dict) -> None:
    """
    CountOutlook 대응 함수.
    outlook 문자열이 긍정적/안정적/부정적(검토) 중 무엇인지에 따라
    counts = {"pos": 0, "neu": 0, "neg": 0} 를 누적 갱신.
    """
    if outlook == POS_STR:
        counts["pos"] += 1
    elif outlook == NEU_STR:
        counts["neu"] += 1
    elif outlook in (NEG_STR, NEG_STR2):
        counts["neg"] += 1


def is_excluded_case(s: str) -> bool:
    """
    코드/섹터 매핑 단계에서 제외할 예외 케이스 판정.

    아래 중 하나라도 해당하면 True:
      - '-', '(', ')' 포함
      - 'MBS' 포함
      - '제'가 '차'보다 앞에 나오는 경우
      - '유동화' 포함
      - 숫자 뒤에 '차'가 오는 경우
      - 숫자 뒤에 '호'가 오는 경우
      - '스팔', '스페' 포함
      - 마지막 글자가 '우' 또는 'C'
    """
    t = s.strip()

    if "-" in t or "(" in t or ")" in t:
        return True

    if "MBS" in t:
        return True

    pos_je = t.find("제")
    pos_cha = t.find("차")
    if pos_je != -1 and pos_cha != -1 and pos_je < pos_cha:
        return True

    if "유동화" in t:
        return True

    for i in range(len(t) - 1):
        if t[i].isdigit() and t[i + 1] == "차":
            return True
        if t[i].isdigit() and t[i + 1] == "호":
            return True

    if "스팔" in t or "스페" in t:
        return True

    if t.endswith("우") or t.endswith("C"):
        return True

    return False


def _cell(row, col_idx):
    """arr_info 한 행(row)에서 col_idx(1-based, VBA 컬럼번호) 값을 문자열로 안전하게 추출."""
    v = row[col_idx]
    return str(v).strip() if v is not None else ""


def parser_info(
    input_path: str = "info_4214.xlsx",
    sheet_name: str = "Sheet1",
    json_output_path: str = "parsed_infomax.json",
) -> str:

    # ------------------------------------------------------------------
    # 0) 엑셀 파일 열기 (함수 안에서 로드 -> import 시점 부작용 없음)
    # ------------------------------------------------------------------
    wb = openpyxl.load_workbook(input_path, data_only=True)
    ws_ref = wb[sheet_name]

    # ------------------------------------------------------------------
    # B열 기준 마지막 데이터 행 찾기
    # ------------------------------------------------------------------
    last_row = ws_ref.max_row
    while (
        last_row > 1
        and ws_ref.cell(row=last_row, column=2).value in (None, "")
    ):
        last_row -= 1

    # ------------------------------------------------------------------
    # A2:AP 범위를 메모리로 로드
    # (인덱스를 VBA 1-based 컬럼번호와 맞추기 위해 맨 앞에 None 패딩)
    # ------------------------------------------------------------------
    arr_info = []
    for row in ws_ref.iter_rows(min_row=2, max_row=last_row, min_col=1, max_col=42):
        arr_info.append([None] + [c.value for c in row])

    # ------------------------------------------------------------------
    # 1) 종목명 -> 코드 / 섹터 / 행번호 매핑 (첫 등장 행만 유지)
    # ------------------------------------------------------------------
    code_a = {}
    sector_a = {}
    row_map_a = {}

    for i, row in enumerate(arr_info):
        raw = _cell(row, 2)
        if raw != "" and raw not in code_a:
            code_a[raw] = _cell(row, 35)
            sector_a[raw] = _cell(row, 38)
            row_map_a[raw] = i

    # ------------------------------------------------------------------
    # 2) 예외 케이스 제거 후 오름차순 정렬 (VBA: SimpleSort)
    # ------------------------------------------------------------------
    only_a = [key for key in code_a.keys() if not is_excluded_case(key)]
    only_a.sort()

    # ------------------------------------------------------------------
    # 3) 종목별 계산
    # ------------------------------------------------------------------
    out_rows = []

    for name in only_a:
        info_row = arr_info[row_map_a[name]]

        # bond 기준등급 = C열(3), corp 기준등급 = V열(22)
        # "CANC" 는 취소된 등급이므로 빈 값 취급
        bond_base_rating = _cell(info_row, 3)
        corp_base_rating = _cell(info_row, 22)

        if bond_base_rating == "CANC":
            bond_base_rating = ""
        if corp_base_rating == "CANC":
            corp_base_rating = ""

        # 둘 다 없으면 해당 종목은 결과에서 제외 (VBA: GoTo NextEntity)
        if bond_base_rating == "" and corp_base_rating == "":
            continue

        row_out = {
            "code_info": code_a[name],
            "code_fn": "",             # VBA 원본에서도 항상 빈 문자열
            "sector": sector_a[name],
            "name": name,
        }

        # --- Bond (회사채) ---
        row_out["bond_rating"] = bond_base_rating
        row_out["bond_date"] = _cell(info_row, 4)   # D열

        if bond_base_rating == "":
            row_out["bond_TF"] = ""
        else:
            # E열(5), H열(8), K열(11) 각 평가사 등급이 기준등급과 일치하는지 확인
            row_out["bond_TF"] = tf_check(
                bond_base_rating,
                _cell(info_row, 5), _cell(info_row, 8), _cell(info_row, 11),
            )

        # bond 등급전망(F열=6, I열=9, L열=12)은 기준등급 유무와 상관없이 항상 집계
        bond_counts = {"pos": 0, "neu": 0, "neg": 0}
        for col_idx in (6, 9, 12):
            count_outlook(_cell(info_row, col_idx), bond_counts)
        row_out["bond_pos"] = bond_counts["pos"]
        row_out["bond_neu"] = bond_counts["neu"]
        row_out["bond_neg"] = bond_counts["neg"]

        # --- CP (기업어음) ---
        cp_base_rating = _cell(info_row, 14)  # N열
        row_out["cp_rating"] = cp_base_rating
        row_out["cp_date"] = _cell(info_row, 15)  # O열

        if cp_base_rating == "":
            row_out["cp_TF"] = ""
        else:
            # P열(16), R열(18), T열(20)
            row_out["cp_TF"] = tf_check(
                cp_base_rating,
                _cell(info_row, 16), _cell(info_row, 18), _cell(info_row, 20),
            )

        # --- Corp (무보증사채 등) ---
        row_out["corp_rating"] = corp_base_rating
        row_out["corp_date"] = _cell(info_row, 23)  # W열

        if corp_base_rating == "":
            row_out["corp_TF"] = ""
        else:
            # X열(24), AA열(27), AD열(30)
            row_out["corp_TF"] = tf_check(
                corp_base_rating,
                _cell(info_row, 24), _cell(info_row, 27), _cell(info_row, 30),
            )

        # corp 등급전망(Y열=25, AB열=28, AE열=31)도 항상 집계
        corp_counts = {"pos": 0, "neu": 0, "neg": 0}
        for col_idx in (25, 28, 31):
            count_outlook(_cell(info_row, col_idx), corp_counts)
        row_out["corp_pos"] = corp_counts["pos"]
        row_out["corp_neu"] = corp_counts["neu"]
        row_out["corp_neg"] = corp_counts["neg"]

        out_rows.append(row_out)

    # ------------------------------------------------------------------
    # 4) JSON으로 저장 (엑셀 out_info 시트 대신)
    # ------------------------------------------------------------------
    with open(json_output_path, "w", encoding="utf-8") as f:
        json.dump(out_rows, f, ensure_ascii=False, indent=2)

    print(f"Done. {len(out_rows)} companies. -> {json_output_path}")
    return json_output_path

if __name__ == "__main__":
    parser_info()