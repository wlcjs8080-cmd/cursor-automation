# -*- coding: utf-8 -*-
"""
Step 2: Part별 이전 교체일 (K57~K63) 자동 입력 (xlwings)
- 레포트 파일을 엑셀에서 열어둔 상태에서 터미널에서 python excel_step2_parts.py 실행
- xw.books.active로 열려 있는 레포트 접근 → N11(S/N), I57~I63(품번) 읽기
- 마스터 시트 'New Alarm 및 사용Part 이력'에서 S/N(L)+품번(AC) 매칭 → 오늘 이전 최근 작업일자(P) → K57~K63 입력
- .cursor/rules/ RULE.md 및 .cursorrules 준수 (openpyxl 금지, try-finally로 마스터 Excel 종료)
"""

from pathlib import Path
from datetime import datetime, date, timedelta

import xlwings as xw

# ---------------------------------------------------------------------------
# 설정 (.cursorrules / project-context RULE.md 기준)
# ---------------------------------------------------------------------------
BASE_PATH = Path(r"C:\정동교\문서 자동화 TEST\커서 바이브코딩 자동화 관련")
MASTER_FOLDER = BASE_PATH / "마스터 시트"
REPORT_SHEET_NAME = "PM REPORT (새양식)"
MASTER_SHEET_ALARM = "New Alarm 및 사용Part 이력"

# 레포트 시트 셀
CELL_SN = "N11"
RANGE_PART_NO = "I57:I63"   # 품번 (사람이 입력)
RANGE_PREV_DATE = "K57:K63" # Part별 이전 교체일 (여기에 입력)

# New Alarm 및 사용Part 이력: 머릿말 2행, 데이터 3행부터
ALARM_HEADER_ROW = 2
ALARM_DATA_START_ROW = 3
COL_ALARM_SN = 12    # L열 = S/N
COL_ALARM_PART = 29  # AC열 = 품번
COL_ALARM_WORK_DATE = 16  # P열 = 작업일자
COL_ALARM_PARTNAME = 28   # AB열 = 파트명
COL_ALARM_SPEC = 37       # AK열 = 규격
NO_HISTORY_TEXT = "교체이력 없음"


def _get_sheet_by_name(book, name):
    """북에서 시트명으로 시트 반환, 없으면 None."""
    for s in book.sheets:
        if s.name == name:
            return s
    return None


def _to_date(val):
    """셀 값을 오늘과 비교 가능한 date로 변환. 불가능하면 None."""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    if isinstance(val, (int, float)):
        try:
            return date(1899, 12, 30) + timedelta(days=int(val))
        except (ValueError, OverflowError):
            return None
    s = str(val).strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%d/%m/%Y", "%Y.%m.%d"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    return None


def _read_all_data(sheet):
    """시트의 used_range 전체를 2D 리스트로 한 번에 읽어온다. 빈 시트면 빈 리스트 반환."""
    try:
        used = sheet.used_range
        if used is None:
            return []
        data = used.value
        if data is None:
            return []
        # 단일 셀이면 [[값]] 형태로 통일
        if not isinstance(data, list):
            return [[data]]
        # 단일 행이면 [값, 값, ...] → [[값, 값, ...]]
        if len(data) > 0 and not isinstance(data[0], list):
            return [data]
        return data
    except Exception:
        return []


def _build_sn_part_to_latest_date(all_data, today):
    """
    전체 데이터(2D 리스트)에서
    (S/N, 품번) -> 오늘 이전 가장 최근 작업일자(엑셀에 넣을 원본 값) 딕셔너리 생성.
    all_data[0] = 엑셀 1행, all_data[2] = 엑셀 3행(데이터 시작)
    """
    result = {}  # (sn_norm, part_norm) -> (best_date, best_raw_value)
    # 데이터는 3행부터 → 인덱스 2부터
    for row_idx in range(ALARM_DATA_START_ROW - 1, len(all_data)):
        row = all_data[row_idx]
        # 열 번호는 1부터 시작 → 리스트 인덱스는 열번호-1
        if len(row) < max(COL_ALARM_SN, COL_ALARM_PART, COL_ALARM_WORK_DATE):
            continue
        sn_val = row[COL_ALARM_SN - 1]       # L열
        part_val = row[COL_ALARM_PART - 1]    # AC열
        work_val = row[COL_ALARM_WORK_DATE - 1]  # P열
        d = _to_date(work_val)
        if d is None or d >= today:
            continue
        sn_norm = str(sn_val).strip() if sn_val is not None else ""
        part_norm = str(part_val).strip() if part_val is not None else ""
        if not sn_norm and not part_norm:
            continue
        key = (sn_norm, part_norm)
        if key not in result or d > result[key][0]:
            result[key] = (d, work_val)
    return {k: v[1] for k, v in result.items()}  # (sn, part) -> raw value for Excel


def _build_part_to_name_spec(all_data):
    """
    전체 데이터(2D 리스트)에서 품번(AC)만 매칭하여
    품번 -> (파트명 AB, 규격 AK) 딕셔너리 생성. 같은 품번 여러 행이면 첫 행 기준.
    """
    result = {}
    for row_idx in range(ALARM_DATA_START_ROW - 1, len(all_data)):
        row = all_data[row_idx]
        if len(row) < max(COL_ALARM_PART, COL_ALARM_PARTNAME, COL_ALARM_SPEC):
            continue
        part_val = row[COL_ALARM_PART - 1]    # AC열
        part_norm = str(part_val).strip() if part_val is not None else ""
        if not part_norm or part_norm in result:
            continue
        part_name = row[COL_ALARM_PARTNAME - 1]  # AB열
        spec_val = row[COL_ALARM_SPEC - 1]        # AK열
        result[part_norm] = (part_name, spec_val)
    return result


def run_step2():
    """
    열려 있는 레포트에서 S/N·품번 읽고, 마스터에서 이전 교체일 조회 후 K57~K63에 입력.
    레포트는 저장하지 않음.
    """
    # 1) 열려 있는 레포트 접근
    try:
        report_book = xw.books.active
    except Exception:
        print("오류: 활성 워크북이 없습니다. 레포트 파일을 엑셀에서 먼저 열어두세요.")
        return
    if report_book is None:
        print("오류: 활성 워크북이 없습니다. 레포트 파일을 엑셀에서 먼저 열어두세요.")
        return

    if REPORT_SHEET_NAME not in [s.name for s in report_book.sheets]:
        print(f"오류: 이 워크북에 '{REPORT_SHEET_NAME}' 시트가 없습니다.")
        return
    report_sheet = report_book.sheets[REPORT_SHEET_NAME]

    # 2) N11에서 S/N 읽기
    sn_val = report_sheet.range(CELL_SN).value
    sn_str = str(sn_val).strip() if sn_val is not None else ""
    if not sn_str:
        print("오류: N11에 S/N이 없습니다.")
        return

    # 3) I57~I63에서 품번 읽기 (비어 있는 셀은 건너뜀)
    part_numbers = []
    for row_no in range(57, 64):
        cell = report_sheet.range((row_no, 9))  # I열 = 9
        val = cell.value
        part_numbers.append((row_no, val))

    # 3-2) 우측 V57~V61에서 품번 읽기
    part_numbers_right = []
    for row_no in range(57, 62):  # 57~61
        cell = report_sheet.range((row_no, 22))  # V열 = 22
        val = cell.value
        part_numbers_right.append((row_no, val))

    # 3-1) 점검일(V8) 읽기 — 사용일 계산용
    v8_val = report_sheet.range("V8").value
    inspection_date = _to_date(v8_val)

    # 4) 마스터 시트 폴더에서 파일 목록 (백업·임시 제외)
    if not MASTER_FOLDER.is_dir():
        print(f"오류: 마스터 폴더를 찾을 수 없습니다. {MASTER_FOLDER}")
        return
    master_files = list(MASTER_FOLDER.glob("*.xlsx")) + list(MASTER_FOLDER.glob("*.xlsm"))
    master_files = [p for p in master_files if "_backup_" not in p.name and not p.name.startswith("~$")]
    if not master_files:
        print(f"마스터 파일이 없습니다. ({MASTER_FOLDER})")
        return
    master_path = master_files[0]

    # 5) 마스터를 백그라운드(visible=False)로 열기 → try-finally로 반드시 종료
    app_master = None
    master_book = None
    try:
        app_master = xw.App(visible=False)
        app_master.display_alerts = False
        master_book = app_master.books.open(str(master_path))

        alarm_sheet = _get_sheet_by_name(master_book, MASTER_SHEET_ALARM)

        # ★ 전체 데이터를 한 번에 읽기 (속도 개선 핵심)
        if alarm_sheet is not None:
            all_data = _read_all_data(alarm_sheet)
        else:
            all_data = []

        today = date.today()
        sn_part_to_date = _build_sn_part_to_latest_date(all_data, today)
        part_to_name_spec = _build_part_to_name_spec(all_data)

        # 6-1) 점검일(V8)이 없으면 오류 메시지 후 L57~L63, Y57~Y61 전부 0 입력
        if inspection_date is None:
            print("오류: V8(점검일)이 비어있거나 날짜로 변환할 수 없습니다. L57~L63, Y57~Y61에 0을 입력합니다.")
            for r in range(57, 64):
                report_sheet.range((r, 12)).value = 0  # L열 = 12
            for r in range(57, 62):
                report_sheet.range((r, 25)).value = 0  # Y열 = 25

        # 7) 각 품번마다 매칭 → B(파트명), F(규격), K(이전 교체일), L(사용일) 입력, 같은 행 I열 글자 크기 적용
        for row_no, part_val in part_numbers:
            if part_val is None or str(part_val).strip() == "":
                continue
            part_norm = str(part_val).strip()
            key = (sn_str, part_norm)
            found_date = sn_part_to_date.get(key)
            name_spec = part_to_name_spec.get(part_norm)

            i_cell = report_sheet.range((row_no, 9))   # I열 = 품번
            b_cell = report_sheet.range((row_no, 2))   # B열 = 파트명
            f_cell = report_sheet.range((row_no, 6))   # F열 = 규격
            k_cell = report_sheet.range((row_no, 11))  # K열 = 이전 교체일
            l_cell = report_sheet.range((row_no, 12))  # L열 = 사용일

            if name_spec is not None:
                part_name_val, spec_val = name_spec
                b_cell.value = part_name_val
                f_cell.value = spec_val
            else:
                b_cell.value = NO_HISTORY_TEXT
                f_cell.value = NO_HISTORY_TEXT

            if found_date is not None:
                k_cell.value = found_date
            else:
                k_cell.value = NO_HISTORY_TEXT
                print(f"매칭 없음: S/N={sn_str}, 품번={part_norm} (행 {row_no})")

            # 사용일(L열): 점검일(V8) - 이전 교체일(K). 교체이력 없음이면 0
            if inspection_date is not None:
                prev_d = _to_date(found_date) if found_date is not None else None
                days_used = (inspection_date - prev_d).days if prev_d is not None else 0
                l_cell.value = days_used

            # 셀에 맞춤 축소(ShrinkToFit) 적용 (좌측)
            try:
                i_cell.api.ShrinkToFit = True
            except Exception:
                pass
            try:
                b_cell.api.ShrinkToFit = True
            except Exception:
                pass
            try:
                f_cell.api.ShrinkToFit = True
            except Exception:
                pass
            try:
                k_cell.api.ShrinkToFit = True
            except Exception:
                pass
            try:
                l_cell.api.ShrinkToFit = True
            except Exception:
                pass

            try:
                fs = i_cell.font.size
                b_cell.font.size = fs
                f_cell.font.size = fs
                k_cell.font.size = fs
                l_cell.font.size = fs
            except Exception:
                pass

        # 8) 우측 Part 사용 이력: O(파트명), S(규격), X(이전교체일), Y(사용일) — V57~V61 품번 기준
        for row_no, part_val in part_numbers_right:
            if part_val is None or str(part_val).strip() == "":
                continue
            part_norm = str(part_val).strip()
            key = (sn_str, part_norm)
            found_date = sn_part_to_date.get(key)
            name_spec = part_to_name_spec.get(part_norm)

            v_cell = report_sheet.range((row_no, 22))  # V열 = 품번
            o_cell = report_sheet.range((row_no, 15))  # O열 = 파트명
            s_cell = report_sheet.range((row_no, 19))  # S열 = 규격
            x_cell = report_sheet.range((row_no, 24))  # X열 = 이전 교체일
            y_cell = report_sheet.range((row_no, 25))  # Y열 = 사용일

            if name_spec is not None:
                part_name_val, spec_val = name_spec
                o_cell.value = part_name_val
                s_cell.value = spec_val
            else:
                o_cell.value = NO_HISTORY_TEXT
                s_cell.value = NO_HISTORY_TEXT

            if found_date is not None:
                x_cell.value = found_date
            else:
                x_cell.value = NO_HISTORY_TEXT
                print(f"매칭 없음(우측): S/N={sn_str}, 품번={part_norm} (행 {row_no})")

            if inspection_date is not None:
                prev_d = _to_date(found_date) if found_date is not None else None
                days_used = (inspection_date - prev_d).days if prev_d is not None else 0
                y_cell.value = days_used

            # 셀에 맞춤 축소(ShrinkToFit) 적용 (우측)
            try:
                v_cell.api.ShrinkToFit = True
            except Exception:
                pass
            try:
                o_cell.api.ShrinkToFit = True
            except Exception:
                pass
            try:
                s_cell.api.ShrinkToFit = True
            except Exception:
                pass
            try:
                x_cell.api.ShrinkToFit = True
            except Exception:
                pass
            try:
                y_cell.api.ShrinkToFit = True
            except Exception:
                pass

            try:
                fs = v_cell.font.size
                o_cell.font.size = fs
                s_cell.font.size = fs
                x_cell.font.size = fs
                y_cell.font.size = fs
            except Exception:
                pass

    finally:
        if master_book is not None:
            try:
                master_book.close(save=False)
            except Exception:
                pass
        if app_master is not None:
            try:
                app_master.quit()
            except Exception:
                pass

    # 레포트는 저장하지 않음 (사람이 확인 후 직접 저장)
    print("Step 2 완료. K57~K63을 확인한 뒤 필요하면 저장하세요.")


if __name__ == "__main__":
    run_step2()
