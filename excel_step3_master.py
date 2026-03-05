# -*- coding: utf-8 -*-
"""
Step 3: 레포트 + 스케줄 → 마스터(New Alarm 및 사용Part 이력) 기입
- 스케줄 시트에서 AC열="완료"인 행을 찾아 해당 레포트 파일을 OPEN
- 레포트/스케줄 데이터를 모아서 마스터 시트에 새 행으로 추가한다.
- 유상/무상(Z63)에 따라 "작업" 행 1개 + "파트" 행 여러 개를 생성한다.
- .cursor/rules/ RULE.md 및 .cursorrules 규칙 준수:
  - xlwings만 사용 (openpyxl 절대 금지)
  - 실행 전 자동 백업 생성 (마스터 파일 대상)
  - backup / ~$ 임시 파일 제외
  - 마스터 시트는 visible=False + try-finally로 Excel 프로세스 종료
  - 마스터 시트 필터 자동 해제 후 작업
"""

import sys
from pathlib import Path
from datetime import datetime, date, timedelta

import xlwings as xw

# ---------------------------------------------------------------------------
# 경로 및 시트 설정 (project-context / .cursorrules 기준)
# ---------------------------------------------------------------------------
BASE_PATH = Path(r"C:\정동교\문서 자동화 TEST\커서 바이브코딩 자동화 관련")
SCHEDULE_FOLDER = BASE_PATH / "스케쥴 시트"
MASTER_FOLDER = BASE_PATH / "마스터 시트"
CUSTOMER_FOLDER = BASE_PATH / "고객사 폴더"

SCHEDULE_SHEET_NAME = "Sheet1"
REPORT_SHEET_NAME = "PM REPORT (새양식)"
MASTER_SHEET_ALARM = "New Alarm 및 사용Part 이력"

HEADER_ROWS = [93, 138, 183, 229, 274, 319, 364]
DATA_ROWS_PER_BLOCK = 41

# 스케줄 열 (1-based)
COL_VISIT_1 = 12   # L열 = 1차 방문
COL_VISIT_2 = 13   # M열 = 2차 방문
COL_VISIT_3 = 14   # N열 = 3차 방문
COL_CUSTOMER = 16  # P열 = 고객사
COL_MODEL = 25     # Y열 = MODEL
COL_SERIAL = 27    # AA열 = S/N
COL_WORK = 20      # T열 = 업무내용
COL_STAFF = 22     # V열 = 작업인원
COL_LINE = 24      # X열 = 라인
COL_UNIT = 26      # Z열 = 설비호기
COL_PROCESS_DONE = 29  # AC열 = 처리완료 ("완료")

# 마스터: 머릿말 2행, 데이터 3행부터
MASTER_DATA_START_ROW = 3


def backup_file(file_path: Path) -> Path:
    """원본 파일을 백업 (파일명_backup_YYYYMMDD_HHMMSS.xlsx)."""
    stem = file_path.stem
    suffix = file_path.suffix
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = file_path.parent / f"{stem}_backup_{stamp}{suffix}"
    import shutil

    shutil.copy2(file_path, backup_path)
    return backup_path


def sanitize_filename(s):
    """파일명에 사용할 수 없는 문자 제거."""
    if s is None:
        return ""
    s = str(s).strip()
    for c in r'\/:*?"<>|':
        s = s.replace(c, "")
    return s[:200]


def make_report_filename(customer, yyyymmdd, model, unit, serial, work):
    """
    Step 1과 동일한 파일명 규칙:
    고객사_날짜_MODEL_호기_(S/N)_업무내용 件.xlsx
    """
    c = sanitize_filename(customer)
    e = sanitize_filename(model)
    u = sanitize_filename(unit)
    s = sanitize_filename(serial)
    w = sanitize_filename(work)
    return f"{c}_{yyyymmdd}_{e}_{u}_({s})_{w} 件.xlsx"


def find_customer_report_folder(customer_name):
    """고객사 폴더 안에서 고객사명이 포함된 폴더/REPORT 경로 찾기."""
    if not customer_name or not str(customer_name).strip():
        return None
    name = str(customer_name).strip()
    if not CUSTOMER_FOLDER.is_dir():
        return None
    for path in CUSTOMER_FOLDER.iterdir():
        if not path.is_dir():
            continue
        if name in path.name:
            report_dir = path / "REPORT"
            if report_dir.is_dir():
                return report_dir
            return None
    return None


def _get_sheet_by_name(book, target_name):
    """시트명이 target_name과 일치하는 시트 반환 (앞뒤 공백 무시). 없으면 None."""
    name_stripped = target_name.strip()
    for s in book.sheets:
        if s.name.strip() == name_stripped:
            return s
    return None


def _read_all_data(sheet):
    """시트의 used_range 전체를 2D 리스트로 한 번에 읽어온다."""
    try:
        used = sheet.used_range
        if used is None:
            return []
        data = used.value
        if data is None:
            return []
        if not isinstance(data, list):
            return [[data]]
        if len(data) > 0 and not isinstance(data[0], list):
            return [data]
        return data
    except Exception:
        return []


def get_visit_date_and_yyyymmdd(val_l, val_m, val_n):
    """
    방문일정: N(3차) > M(2차) > L(1차) 우선순위로 선택.
    파일명용 YYYYMMDD 문자열도 함께 반환.
    """
    for val in (val_n, val_m, val_l):
        if val is None:
            continue
        if isinstance(val, (datetime, date)):
            d = val.date() if isinstance(val, datetime) else val
        elif isinstance(val, (int, float)):
            # Excel 날짜 직렬값 가정
            try:
                d = date(1899, 12, 30) + timedelta(days=int(val))
            except Exception:
                continue
        else:
            try:
                d = datetime.strptime(str(val).strip()[:10], "%Y-%m-%d").date()
            except Exception:
                continue
        return d, d.strftime("%Y%m%d")
    return None, None


def find_master_next_row(master_sheet):
    """
    마스터 시트에서 D(4), L(12), P(16) 세 개가 모두 채워진 마지막 행을 찾고,
    그 아래(다음) 행 번호를 반환한다.
    """
    data = _read_all_data(master_sheet)
    if not data:
        return MASTER_DATA_START_ROW

    last_row = MASTER_DATA_START_ROW - 1
    for idx in range(len(data) - 1, MASTER_DATA_START_ROW - 2, -1):
        row = data[idx]
        # 최소 16열까지 있는지 확인
        if len(row) < 16:
            continue
        val_d = row[3]
        val_l = row[11]
        val_p = row[15]
        if (val_d is not None and str(val_d).strip() != "") and \
           (val_l is not None and str(val_l).strip() != "") and \
           (val_p is not None and str(val_p).strip() != ""):
            last_row = idx + 1  # 1-based
            break
    if last_row < MASTER_DATA_START_ROW:
        last_row = MASTER_DATA_START_ROW - 1
    return last_row + 1


def is_master_in_use_by_temp_file():
    """마스터 시트 폴더 내 ~$ 임시 파일 존재 여부로 사용 중인지 확인."""
    if not MASTER_FOLDER.is_dir():
        return False
    for p in MASTER_FOLDER.iterdir():
        if p.is_file() and p.name.startswith("~$") and p.suffix.lower() in (".xlsx", ".xlsm"):
            return True
    return False


def main():
    # 폴더 존재 확인
    if not SCHEDULE_FOLDER.is_dir():
        print(f"오류: 스케줄 폴더를 찾을 수 없습니다. {SCHEDULE_FOLDER}")
        return
    if not MASTER_FOLDER.is_dir():
        print(f"오류: 마스터 폴더를 찾을 수 없습니다. {MASTER_FOLDER}")
        return

    schedule_files = [p for p in list(SCHEDULE_FOLDER.glob("*.xlsx")) + list(SCHEDULE_FOLDER.glob("*.xlsm"))
                      if "_backup_" not in p.name and not p.name.startswith("~$")]
    if not schedule_files:
        print(f"처리할 스케줄 파일이 없습니다. ({SCHEDULE_FOLDER})")
        return

    master_files = [p for p in list(MASTER_FOLDER.glob("*.xlsx")) + list(MASTER_FOLDER.glob("*.xlsm"))
                    if "_backup_" not in p.name and not p.name.startswith("~$")]
    if not master_files:
        print(f"마스터 파일이 없습니다. ({MASTER_FOLDER})")
        return

    schedule_path = schedule_files[0]
    master_path = master_files[0]

    # 마스터 시트 사용 중(~$ 임시 파일) 체크
    if is_master_in_use_by_temp_file():
        print("마스터 시트가 사용 중입니다. 닫히면 다시 실행하세요.")
        return

    # 마스터 파일 백업
    try:
        backup = backup_file(master_path)
        print(f"마스터 백업 완료: {backup.name}")
    except Exception as e:
        print(f"마스터 백업 실패: {e}")
        sys.exit(1)

    app = None
    try:
        app = xw.App(visible=False)
        app.display_alerts = False

        # 스케줄 열기 (읽기 전용 용도)
        schedule_wb = app.books.open(str(schedule_path))
        if SCHEDULE_SHEET_NAME not in [s.name for s in schedule_wb.sheets]:
            print(f"오류: 스케줄 시트에 '{SCHEDULE_SHEET_NAME}'이(가) 없습니다.")
            schedule_wb.close()
            return
        schedule_sheet = schedule_wb.sheets[SCHEDULE_SHEET_NAME]

        # 마스터 열기 (쓰기)
        master_wb = app.books.open(str(master_path))

        # 읽기 전용 여부 체크
        try:
            if bool(getattr(master_wb.api, "ReadOnly", False)):
                print("마스터 시트가 읽기 전용으로 열려 있습니다. 닫히면 다시 실행하세요.")
                master_wb.close()
                schedule_wb.close()
                return
        except Exception:
            pass

        master_sheet = _get_sheet_by_name(master_wb, MASTER_SHEET_ALARM)
        if master_sheet is None:
            print(f"오류: 마스터 파일에 시트 '{MASTER_SHEET_ALARM}'이(가) 없습니다.")
            master_wb.close()
            schedule_wb.close()
            return

        # 새 행 시작 위치 계산
        next_row = find_master_next_row(master_sheet)

        # 기존 마지막 데이터 행의 일부 열 값 복사용 (C, E, F, K)
        base_row = next_row - 1
        if base_row >= MASTER_DATA_START_ROW:
            base_c = master_sheet.range((base_row, 3)).value  # C열
            base_e = master_sheet.range((base_row, 5)).value  # E열
            base_f = master_sheet.range((base_row, 6)).value  # F열
            base_k = master_sheet.range((base_row, 11)).value  # K열
        else:
            base_c = base_e = base_f = base_k = None

        success_count = 0
        skip_count = 0

        # 3월(HEADER_ROWS >= 183) 블록부터, AC="완료" 행만 처리
        for header_row in [r for r in HEADER_ROWS if r >= 183]:
            start_row = header_row + 1
            end_row = header_row + DATA_ROWS_PER_BLOCK
            for row_num in range(start_row, end_row + 1):
                try:
                    val_ac = schedule_sheet.range((row_num, COL_PROCESS_DONE)).value
                    if val_ac is None or str(val_ac).strip() != "완료":
                        continue

                    customer = schedule_sheet.range((row_num, COL_CUSTOMER)).value
                    model = schedule_sheet.range((row_num, COL_MODEL)).value
                    serial = schedule_sheet.range((row_num, COL_SERIAL)).value
                    work = schedule_sheet.range((row_num, COL_WORK)).value
                    staff = schedule_sheet.range((row_num, COL_STAFF)).value
                    line = schedule_sheet.range((row_num, COL_LINE)).value
                    unit = schedule_sheet.range((row_num, COL_UNIT)).value
                    val_l = schedule_sheet.range((row_num, COL_VISIT_1)).value
                    val_m = schedule_sheet.range((row_num, COL_VISIT_2)).value
                    val_n = schedule_sheet.range((row_num, COL_VISIT_3)).value

                    visit_val, yyyymmdd = get_visit_date_and_yyyymmdd(val_l, val_m, val_n)
                    if visit_val is None or yyyymmdd is None:
                        print(f"행 {row_num}: 방문일정 없음 → 건너뜀")
                        skip_count += 1
                        continue

                    if not customer or not str(customer).strip():
                        print(f"행 {row_num}: 고객사명 누락 → 건너뜀")
                        skip_count += 1
                        continue
                    if not model or not str(model).strip():
                        print(f"행 {row_num}: MODEL(Y열) 누락 → 건너뜀")
                        skip_count += 1
                        continue
                    if not serial or not str(serial).strip():
                        print(f"행 {row_num}: S/N(AA열) 누락 → 건너뜀")
                        skip_count += 1
                        continue
                    if not work or not str(work).strip():
                        print(f"행 {row_num}: 업무내용(T열) 누락 → 건너뜀")
                        skip_count += 1
                        continue

                    # 고객사 REPORT 폴더
                    report_dir = find_customer_report_folder(customer)
                    if report_dir is None:
                        print(f"행 {row_num}: 고객사 REPORT 폴더를 찾을 수 없음 (고객사: {customer})")
                        skip_count += 1
                        continue

                    filename = make_report_filename(customer, yyyymmdd, model, unit, serial, work)
                    report_path = report_dir / filename
                    if not report_path.is_file():
                        print(f"행 {row_num}: 레포트 파일을 찾을 수 없음 ({report_path})")
                        skip_count += 1
                        continue

                    # 레포트 열기 (읽기 전용으로만 사용)
                    report_wb = app.books.open(str(report_path))
                    try:
                        if REPORT_SHEET_NAME not in [s.name for s in report_wb.sheets]:
                            print(f"행 {row_num}: 레포트에 '{REPORT_SHEET_NAME}' 시트 없음 → 건너뜀")
                            skip_count += 1
                            continue
                        rs = report_wb.sheets[REPORT_SHEET_NAME]

                        # 유상/무상 판단 (Z63)
                        charge = rs.range("Z63").value
                        charge_str = str(charge).strip() if charge is not None else ""
                        if charge_str not in ("유상", "무상"):
                            # 업체/빈칸 등은 건너뜀
                            print(f"행 {row_num}: Z63='{charge_str}' (유상/무상 아님) → 건너뜀")
                            skip_count += 1
                            continue

                        # KK 데이터 (작업/파트 공통)
                        kk_customer = rs.range("E10").value   # 고객사
                        kk_sn = rs.range("N11").value         # S/N
                        kk_model = rs.range("N10").value      # MODEL
                        kk_turn_on = rs.range("V10").value    # TURN ON
                        kk_date = rs.range("V8").value        # 작업일자
                        kk_prev = rs.range("V11").value       # 이전 방문일
                        kk_start_time = rs.range("V9").value  # 시작시간
                        kk_end_time = rs.range("Y9").value    # 종료시간
                        kk_time = rs.range("AA9").value       # 작업시간
                        kk_problem = rs.range("B17").value    # 문제/현상
                        kk_cause = rs.range("B19").value      # 원인

                        # 좌측/우측 Part 정보 수집
                        left_parts = []
                        for r in range(57, 64):
                            part_no = rs.range((r, 9)).value  # I열
                            if part_no is None or str(part_no).strip() == "":
                                continue
                            part_name = rs.range((r, 2)).value   # B열
                            spec = rs.range((r, 6)).value        # F열
                            used_days = rs.range((r, 12)).value  # L열
                            left_parts.append((part_name, part_no, used_days, spec))

                        right_parts = []
                        for r in range(57, 62):
                            part_no = rs.range((r, 22)).value  # V열
                            if part_no is None or str(part_no).strip() == "":
                                continue
                            part_name = rs.range((r, 15)).value  # O열
                            spec = rs.range((r, 19)).value       # S열
                            used_days = rs.range((r, 25)).value  # Y열
                            right_parts.append((part_name, part_no, used_days, spec))

                        # "작업" 행 (유상/무상 공통)
                        if charge_str in ("유상", "무상"):
                            row = next_row
                            master_sheet.range((row, 2)).value = "작업"  # B
                            master_sheet.range((row, 3)).value = base_c  # C
                            master_sheet.range((row, 4)).value = kk_customer  # D 고객사
                            master_sheet.range((row, 5)).value = base_e  # E
                            master_sheet.range((row, 6)).value = base_f  # F
                            master_sheet.range((row, 7)).value = line         # G 라인
                            master_sheet.range((row, 10)).value = unit        # J 설비호기
                            master_sheet.range((row, 11)).value = base_k      # K
                            master_sheet.range((row, 12)).value = kk_sn       # L S/N
                            master_sheet.range((row, 13)).value = kk_model    # M MODEL
                            master_sheet.range((row, 15)).value = kk_turn_on  # O TURN ON
                            master_sheet.range((row, 16)).value = kk_date     # P 작업일자
                            master_sheet.range((row, 17)).value = kk_start_time  # Q 시작시간
                            master_sheet.range((row, 18)).value = kk_end_time    # R 종료시간
                            master_sheet.range((row, 19)).value = kk_time     # S 작업시간 (기존값 무조건 덮어쓰기)
                            master_sheet.range((row, 20)).value = staff       # T 작업인원
                            master_sheet.range((row, 24)).value = kk_problem  # X 문제(현상)
                            master_sheet.range((row, 25)).value = kk_cause    # Y 원인
                            master_sheet.range((row, 39)).value = kk_prev     # AM 이전방문일
                            master_sheet.range((row, 28)).value = "인건비"   # AB 인건비
                            next_row += 1

                        # "파트" 행들 (유상/무상 공통, 품번 수만큼)
                        def write_part_row(part_name, part_no, used_days, spec_val):
                            nonlocal next_row
                            # 레포트에서 "교체이력 없음"으로 표시된 품번은 마스터에 기록하지 않음
                            if isinstance(part_name, str) and part_name.strip() == "교체이력 없음":
                                return
                            if isinstance(spec_val, str) and spec_val.strip() == "교체이력 없음":
                                return
                            row = next_row
                            # B~AO까지는 "작업"과 동일한 KK 데이터 기반
                            master_sheet.range((row, 2)).value = "파트"       # B 구분
                            master_sheet.range((row, 3)).value = base_c       # C
                            master_sheet.range((row, 4)).value = kk_customer  # D 고객사
                            master_sheet.range((row, 5)).value = base_e       # E
                            master_sheet.range((row, 6)).value = base_f       # F
                            master_sheet.range((row, 7)).value = line         # G 라인
                            master_sheet.range((row, 10)).value = unit        # J 설비호기
                            master_sheet.range((row, 11)).value = base_k      # K
                            master_sheet.range((row, 12)).value = kk_sn       # L S/N
                            master_sheet.range((row, 13)).value = kk_model    # M MODEL
                            master_sheet.range((row, 15)).value = kk_turn_on  # O TURN ON
                            master_sheet.range((row, 16)).value = kk_date     # P 작업일자
                            master_sheet.range((row, 17)).value = kk_start_time  # Q 시작시간
                            master_sheet.range((row, 18)).value = kk_end_time    # R 종료시간
                            master_sheet.range((row, 19)).value = kk_time     # S 작업시간
                            master_sheet.range((row, 20)).value = staff       # T 작업인원
                            master_sheet.range((row, 24)).value = kk_problem  # X 문제(현상)
                            master_sheet.range((row, 25)).value = kk_cause    # Y 원인
                            master_sheet.range((row, 39)).value = kk_prev     # AM 이전방문일
                            # 파트 정보 (AB, AC, AG, AK)
                            master_sheet.range((row, 28)).value = part_name   # AB 파트명
                            master_sheet.range((row, 29)).value = part_no     # AC 품번
                            master_sheet.range((row, 33)).value = used_days   # AG 사용일
                            master_sheet.range((row, 37)).value = spec_val    # AK 규격
                            next_row += 1

                        for (pn, pno, ud, sp) in left_parts:
                            write_part_row(pn, pno, ud, sp)
                        for (pn, pno, ud, sp) in right_parts:
                            write_part_row(pn, pno, ud, sp)

                        success_count += 1
                        print(f"행 {row_num}: 마스터 기입 완료 (작업/파트)")

                    finally:
                        report_wb.close()

                except Exception as e:
                    skip_count += 1
                    print(f"행 {row_num}: 오류 - {e}")

        master_wb.save()
        master_wb.close()
        schedule_wb.close()

        print("-" * 50)
        print(f"총 {success_count}건 마스터 기입, {skip_count}건 건너뜀")

    finally:
        if app is not None:
            app.quit()


if __name__ == "__main__":
    main()

