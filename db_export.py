# -*- coding: utf-8 -*-
"""
DB → 엑셀 내보내기 스크립트.

- master.db  : master_alarm 전체를 "New Alarm 및 사용Part 이력" 시트(3행부터)에 덮어쓰기
- schedule.db: status="완료"인 행을 스케줄 엑셀 AC열(29열)에 "완료"로 반영
"""

import shutil
import sqlite3
from datetime import datetime
from pathlib import Path

import xlwings as xw

# ---------------------------------------------------------------------------
# 경로 및 기본 설정 (project-context / .cursorrules 기준)
# ---------------------------------------------------------------------------
BASE_PATH = Path(r"C:\정동교\문서 자동화 TEST\커서 바이브코딩 자동화 관련")
MASTER_FOLDER = BASE_PATH / "마스터 시트"
SCHEDULE_FOLDER = BASE_PATH / "스케쥴 시트"
DB_FOLDER = BASE_PATH / "db"

MASTER_DB_PATH = DB_FOLDER / "master.db"
SCHEDULE_DB_PATH = DB_FOLDER / "schedule.db"

MASTER_SHEET_ALARM = "New Alarm 및 사용Part 이력"
SCHEDULE_SHEET_NAME = "Sheet1"

# 스케줄 시트 열 (1-based, project-context RULE.md 기준)
COL_CUSTOMER = 16      # P열 = 고객사
COL_MODEL = 25         # Y열 = MODEL
COL_SN = 27            # AA열 = S/N
COL_WORK = 20          # T열 = 업무내용
COL_VISIT_DONE = 28    # AB열 = 방문완료
COL_PROCESS_DONE = 29  # AC열 = 처리완료

# HEADER_ROWS / DATA_ROWS_PER_BLOCK (스케줄 시트 구조와 동일)
HEADER_ROWS = [93, 138, 183, 229, 274, 319, 364]
DATA_ROWS_PER_BLOCK = 41


def backup_file(file_path: Path) -> Path:
    """원본 파일을 백업 (파일명_backup_YYYYMMDD_HHMMSS.xlsx)."""
    stem = file_path.stem
    suffix = file_path.suffix
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = file_path.parent / f"{stem}_backup_{stamp}{suffix}"
    shutil.copy2(file_path, backup_path)
    return backup_path


def _get_first_excel_file(folder: Path) -> Path | None:
    """폴더에서 첫 번째 유효한 엑셀 파일을 찾는다 (backup/~$ 제외)."""
    if not folder.is_dir():
        return None
    files = list(folder.glob("*.xlsx")) + list(folder.glob("*.xlsm"))
    files = [p for p in files if "_backup_" not in p.name and not p.name.startswith("~$")]
    return files[0] if files else None


def is_in_use_by_temp_file(folder: Path) -> bool:
    """폴더 내 ~$ 임시 파일 존재 여부로 사용 중인지 확인."""
    if not folder.is_dir():
        return False
    for p in folder.iterdir():
        if p.is_file() and p.name.startswith("~$") and p.suffix.lower() in (".xlsx", ".xlsm"):
            return True
    return False


def export_master_db_to_excel() -> int:
    """
    master.db → 마스터 엑셀 "New Alarm 및 사용Part 이력" 시트로 내보내기.
    반환: 내보낸 행 수.
    """
    if not MASTER_DB_PATH.is_file():
        print(f"오류: master.db를 찾을 수 없습니다. ({MASTER_DB_PATH})")
        return 0

    master_path = _get_first_excel_file(MASTER_FOLDER)
    if master_path is None:
        print(f"오류: 마스터 파일을 찾을 수 없습니다. ({MASTER_FOLDER})")
        return 0

    if is_in_use_by_temp_file(MASTER_FOLDER):
        print("마스터 시트가 열려있습니다. 닫고 다시 실행하세요.")
        return 0

    # 내보내기 전 자동 백업
    try:
        backup = backup_file(master_path)
        print(f"마스터 백업 완료: {backup.name}")
    except Exception as e:
        print(f"마스터 백업 실패: {e}")
        return 0

    app = None
    try:
        app = xw.App(visible=False)
        app.display_alerts = False

        wb = app.books.open(str(master_path))
        if MASTER_SHEET_ALARM not in [s.name for s in wb.sheets]:
            print(f"오류: 마스터 파일에 시트 '{MASTER_SHEET_ALARM}'이(가) 없습니다.")
            wb.close()
            return 0
        sheet = wb.sheets[MASTER_SHEET_ALARM]

        # 1) 엑셀 시트에서 마지막 데이터 행 찾기 (3행부터)
        used = sheet.used_range
        last_row = used.last_cell.row if used is not None else 2
        if last_row < 3:
            last_row = 2
        # D열(4) 기준으로 아래에서 위로 스캔해 실제 데이터가 있는 마지막 행을 찾음
        excel_data_last = 2
        for r in range(last_row, 2, -1):
            val = sheet.range((r, 4)).value  # D열=고객사
            if val is not None and str(val).strip() != "":
                excel_data_last = r
                break
        if excel_data_last < 3:
            excel_data_last = 2

        # 엑셀에 이미 있는 데이터 행 수 (3행부터 연속 구간이라고 가정)
        existing_count = excel_data_last - 2

        # 2) master.db에서 총 행 수와 새 행들 조회
        conn = sqlite3.connect(str(MASTER_DB_PATH))
        try:
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*) FROM master_alarm")
            total_count = cur.fetchone()[0]

            if total_count <= existing_count:
                print(f"마스터 DB 총 {total_count}건, 엑셀에 이미 {existing_count}건 존재 → 추가 없음")
                wb.save()
                wb.close()
                return 0

            new_count = total_count - existing_count
            print(f"마스터 DB 총 {total_count}건, 엑셀 {existing_count}건 → 새 행 {new_count}건 추가")

            cur.execute(
                """
                SELECT
                    no, category, department, customer, division, site,
                    line, main_process, sub_process, unit, chamber,
                    sn, model, type, turn_on, work_date,
                    work_time_q, end_time_r, work_time_s, staff_count, man_hour,
                    major_class, minor_class, problem, cause, action,
                    part_type, part_name, part_no, customer_code, qty,
                    warranty, used_days, charge_type, cost, price,
                    spec, warranty_out, prev_visit, month, elapsed_days,
                    initial_defect, charge_flag
                FROM master_alarm
                ORDER BY rowid
                LIMIT ? OFFSET ?
                """,
                (new_count, existing_count),
            )
            new_rows = cur.fetchall()
        finally:
            conn.close()

        # 3) 엑셀 마지막 데이터 행 다음 행부터 새 행만 추가
        if new_rows:
            start_row = excel_data_last + 1
            start_cell = sheet.range((start_row, 1))
            data_2d = [list(r) for r in new_rows]
            start_cell.value = data_2d

        wb.save()
        wb.close()
        return len(new_rows)
    finally:
        if app is not None:
            app.quit()


def export_schedule_db_to_excel() -> int:
    """
    schedule.db에서 status="완료"인 행을 스케줄 엑셀 AC열(29열)에 "완료"로 반영.
    반환: 완료 반영 건수.
    """
    if not SCHEDULE_DB_PATH.is_file():
        print(f"오류: schedule.db를 찾을 수 없습니다. ({SCHEDULE_DB_PATH})")
        return 0

    schedule_path = _get_first_excel_file(SCHEDULE_FOLDER)
    if schedule_path is None:
        print(f"오류: 스케줄 파일을 찾을 수 없습니다. ({SCHEDULE_FOLDER})")
        return 0

    if is_in_use_by_temp_file(SCHEDULE_FOLDER):
        print("스케줄 시트가 열려있습니다. 닫고 다시 실행하세요.")
        return 0

    # 내보내기 전 자동 백업
    try:
        backup = backup_file(schedule_path)
        print(f"스케줄 백업 완료: {backup.name}")
    except Exception as e:
        print(f"스케줄 백업 실패: {e}")
        return 0

    # schedule.db에서 완료 행 조회
    conn = sqlite3.connect(str(SCHEDULE_DB_PATH))
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    try:
        cur.execute(
            """
            SELECT
                customer,
                model,
                sn,
                work
            FROM schedule
            WHERE status = '완료'
            """
        )
        rows = cur.fetchall()
    finally:
        conn.close()

    if not rows:
        print("status='완료' 인 스케줄 행이 없습니다.")
        return 0

    app = None
    updated = 0
    try:
        app = xw.App(visible=False)
        app.display_alerts = False

        wb = app.books.open(str(schedule_path))
        if SCHEDULE_SHEET_NAME not in [s.name for s in wb.sheets]:
            print(f"오류: 스케줄 시트에 '{SCHEDULE_SHEET_NAME}'이(가) 없습니다.")
            wb.close()
            return 0
        sheet = wb.sheets[SCHEDULE_SHEET_NAME]

        # 각 완료 행을 스케줄 시트에서 찾아 AC열에 "완료" 기록
        for row in rows:
            customer = row["customer"]
            model = row["model"]
            sn = row["sn"]
            work = row["work"]

            found = False
            for header_row in HEADER_ROWS:
                start_row = header_row + 1
                end_row = header_row + DATA_ROWS_PER_BLOCK
                for r in range(start_row, end_row + 1):
                    val_ab = sheet.range((r, COL_VISIT_DONE)).value
                    if str(val_ab).strip().upper() != "O":
                        continue

                    c = sheet.range((r, COL_CUSTOMER)).value
                    m = sheet.range((r, COL_MODEL)).value
                    s = sheet.range((r, COL_SN)).value
                    w = sheet.range((r, COL_WORK)).value

                    if (
                        str(c).strip() == str(customer).strip()
                        and str(m).strip() == str(model).strip()
                        and str(s).strip() == str(sn).strip()
                        and str(w).strip() == str(work).strip()
                    ):
                        sheet.range((r, COL_PROCESS_DONE)).value = "완료"
                        updated += 1
                        found = True
                        break
                if found:
                    break

        wb.save()
        wb.close()
        return updated
    finally:
        if app is not None:
            app.quit()


def main() -> None:
    master_count = export_master_db_to_excel()
    schedule_count = export_schedule_db_to_excel()
    print(f"마스터 {master_count}건 내보내기 완료, 스케줄 {schedule_count}건 완료 반영")


if __name__ == "__main__":
    main()

