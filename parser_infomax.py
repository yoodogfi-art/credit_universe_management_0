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

컬럼 인덱스는 VBA arrInfo(row, N) 에서 쓰인 숫자를 그대로 사용했습니다.
(A=1, B=2, C=3, ... Z=26, AA=27, ... AP=42)

수정 내역
--------
- (버그) 파일 최상단에 있던 `openpyxl.read_xlsx("info_4214.xlsx")` 삭제.
  openpyxl 에는 read_xlsx 라는 메서드가 없어서 import 시점에 바로 AttributeError 가 남.
  엑셀 로딩은 parser_info() 함수 안의 `openpyxl.load_workbook(input_path, ...)` 에서
  올바르게 처리하고 있었으므로, 그 부분만 쓰면 됨.
- (변경) 결과 저장 방식을 엑셀 out_info 시트 -> "parsed_infomax.json" 로 변경.
"""

import json
import openpyxl
from openpyxl.utils import get_column_letter


# ------------------------------------------------------------------
# 등급전망(outlook) 판정에 쓰는 한글 키워드
# VBA 원본은 ChrW(&Hxxxx) 유니코드 코드포인트로 하드코딩되어 있었으나
# 실제 문자로 풀면 아래와 같습니다.
# ------------------------------------------------------------------
POS_STR = "긍정적"        # 긍정적
NEU_STR = "안정적"        # 안정적
NEG_STR = "부정적"        # 부정적
NEG_STR2 = "부정적검토"    # 부정적검토


def tf_check(base_rating: str, r1: str, r2: str, r3: str) -> str:
    """
    TFCheck 대응 함수.
    기준등급(base_rating)과 r1/r2/r3 중 하나라도 값이 있으면서 기준등급과 다르면 'F',
    전부 같거나 비어있으면 'T'. base_rating 이 비어있으면 빈 문자열.
    ^ 이부분 바꿔야할듯.
    기준등급(base_rating)과 r1/r2/r3 중 하나라도 값이 있으면서, 두개 이상일 경우 모두 동일한 등급이 아닐때 'F',
    하나 이상 존재하고 전부 같으면 T, 비어있으면 NULL
    """
    for r in (r1, r2, r3):
        if r != "" and base_rating =""
        return "F"

    if base_rating == "":
        return ""
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


def is_excluded_from_output(s: str) -> bool:
    """
    IsExcludedFromOutput 대응 함수. (최종 출력 대상에서 제외할 종목명 판정)

    아래 중 하나라도 해당하면 True(제외):
      - '-', '(', ')' 문자를 포함
      - 'MBS' 문자열을 포함
      - '제' 가 '차' 보다 앞에 나오는 경우 (예: "제1차 ...")
      - '유동화' 를 포함
      - 숫자 바로 뒤에 '차' 가 오는 경우 (예: "3차")
    """
    t = s

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

    return False


def is_trash_name(s: str) -> bool:
    """
    IsTrashName 대응 함수. (코드/섹터 매핑을 만드는 단계에서부터 아예 걸러낼 종목명 판정)

    아래 중 하나라도 해당하면 True(쓰레기명):
      - '(' 또는 ')' 를 포함
      - 'MBS' 를 포함
      - 숫자 바로 뒤에 '호' 가 오는 경우 (예: "5호")
      - '스팔' 또는 '스페' 를 포함
      - 마지막 글자가 '우' 또는 'C' 인 경우 (우선주 등)
    """
    t = s.strip()

    if "(" in t or ")" in t:
        return True
    if "MBS" in t:
        return True

    for i in range(len(t) - 1):
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


def parser_info(input_path: str,
                 json_output_path: str = "parsed_infomax.json",
                 source_sheet: str = "info_4214") -> str:
    """
    VBA Sub parser_info() 이식. (결과는 엑셀이 아닌 JSON 파일로 저장)

    Parameters
    ----------
    input_path       : 원본 엑셀 파일 경로 (info_4214 시트 포함)
    json_output_path : 결과를 저장할 JSON 파일 경로 (기본 "parsed_infomax.json")
    source_sheet      : 원본 시트명 (기본 "info_4214")

    Returns
    -------
    저장된 JSON 파일 경로
    """
    # ------------------------------------------------------------------
    # 엑셀 로딩
    # data_only=True : 수식이 아닌, 마지막으로 계산된 값(캐시된 값)을 읽음
    # ------------------------------------------------------------------
    wb = openpyxl.load_workbook(input_path, data_only=True)
    ws_ref = wb[source_sheet]

    # ------------------------------------------------------------------
    # B열(=2) 기준 마지막 데이터 행 찾기 (VBA: Cells(Rows.Count,"B").End(xlUp))
    # ------------------------------------------------------------------
    last_row = ws_ref.max_row
    while last_row > 1 and ws_ref.cell(row=last_row, column=2).value in (None, ""):
        last_row -= 1

    # ------------------------------------------------------------------
    # A2:AP(last_row) 범위를 통째로 로드 (VBA: arrInfo)
    # 인덱스 0에 더미(None)를 넣어 VBA와 동일하게 "컬럼 번호 = 리스트 인덱스"가
    # 되도록 맞춤 (arr_info[i][1] = A열, arr_info[i][2] = B열, ...)
    # ------------------------------------------------------------------
    arr_info = []
    for row in ws_ref.iter_rows(min_row=2, max_row=last_row, min_col=1, max_col=42):
        arr_info.append([None] + [c.value for c in row])

    # ------------------------------------------------------------------
    # 1) 종목명(B열, raw) 별로 "첫 등장 행"만 남겨 매핑 생성
    #    code_a   : 종목명 -> code_fn (AI열=35)
    #    sector_a : 종목명 -> sector  (AL열=38)
    #    row_map_a: 종목명 -> arr_info 내 인덱스 (첫 등장 행)
    # ------------------------------------------------------------------
    code_a, sector_a, row_map_a = {}, {}, {}
    for i, row in enumerate(arr_info):
        raw = _cell(row, 2)  # B열
        if raw != "" and raw not in code_a:
            code_a[raw] = _cell(row, 35)     # AI열
            sector_a[raw] = _cell(row, 38)   # AL열
            row_map_a[raw] = i

    # ------------------------------------------------------------------
    # 2) 쓰레기명 / 출력 제외 대상 필터링 후 오름차순 정렬 (VBA: SimpleSort)
    # ------------------------------------------------------------------
    only_a = [
        key for key in code_a.keys()
        if not is_trash_name(key) and not is_excluded_from_output(key)
    ]
    only_a.sort()  # StrComp(vbBinaryCompare) 오름차순과 동일한 결과

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
    # 사용 예시:
    # python parser_info.py info_4214.xlsx
    # python parser_info.py info_4214.xlsx custom_output.json
    import sys

    if len(sys.argv) < 2:
        print("사용법: python parser_info.py <입력.xlsx> [출력.json]")
    else:
        in_path = sys.argv[1]
        out_path = sys.argv[2] if len(sys.argv) > 2 else "parsed_infomax.json"
        parser_info(in_path, out_path)