# -*- coding: utf-8 -*-
"""
마스터/스케줄 엑셀 → SQLite DB 초기 변환 스크립트.

- master.db  : "New Alarm 및 사용Part 이력" 전체(A~AQ, 43열)
- schedule.db: 스케줄 시트(Sheet1) 중 AB열=O인 행 + status 컬럼
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

SCHEDULE_SHEET_NAME = "Sheet1"
MASTER_SHEET_ALARM = "New Alarm 및 사용Part 이력"

# 스케줄 시트 열 (1-based, project-context RULE.md 기준)
COL_RECEIPT_DATE = 10  # J열 = 접수일자
COL_VISIT_PLAN = 11    # K열 = 방문계획
COL_VISIT_1 = 12       # L열 = 1차방문
COL_VISIT_2 = 13       # M열 = 2차방문
COL_VISIT_3 = 14       # N열 = 3차방문
COL_REASON = 15        # O열 = 미준수사유
COL_CUSTOMER = 16      # P열 = 고객사
COL_MANAGER = 17       # Q열 = 담당자
COL_RECEIVER = 18      # R열 = 접수자
COL_MAT_WORKER = 19    # S열 = MAT작업자
COL_WORK = 20          # T열 = 업무내용
COL_CHARGE_TYPE = 21   # U열 = 유무상
COL_PEOPLE = 22        # V열 = 인원
COL_PROCESS = 23       # W열 = 공정
COL_LINE = 24          # X열 = 라인
COL_MODEL = 25         # Y열 = MODEL
COL_UNIT = 26          # Z열 = 설비호기
COL_SN = 27            # AA열 = S/N
COL_VISIT_DONE = 28    # AB열 = 방문완료
COL_PROCESS_DONE = 29  # AC열 = 처리완료

# 스케줄 시트 HEADER_ROWS (project-context RULE.md 기준)
HEADER_ROWS = [93, 138, 183, 229, 274, 319, 364]
DATA_ROWS_PER_BLOCK = 41


def backup_db_file(db_path: Path) -> None:
    """기존 DB 파일이 있으면 master_backup_YYYYMMDD_HHMMSS.db 형태로 백업."""
    if not db_path.exists():
        return
    DB_FOLDER.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"{db_path.stem}_backup_{stamp}{db_path.suffix}"
    backup_path = db_path.with_name(backup_name)
    shutil.copy2(db_path, backup_path)


def _get_first_excel_file(folder: Path) -> Path | None:
    """폴더에서 첫 번째 유효한 엑셀 파일을 찾는다 (backup/~$ 제외)."""
    if not folder.is_dir():
        return None
    files = list(folder.glob("*.xlsx")) + list(folder.glob("*.xlsm"))
    files = [p for p in files if "_backup_" not in p.name and not p.name.startswith("~$")]
    return files[0] if files else None


def _get_sheet_by_name(book, target_name: str):
    """시트명이 target_name과 일치하는 시트 반환 (앞뒤 공백 무시). 없으면 None."""
    name_stripped = target_name.strip()
    for s in book.sheets:
        if s.name.strip() == name_stripped:
            return s
    return None


def is_visit_done_o(val) -> bool:
    """AB열 값이 대문자 'O'인지 확인."""
    if val is None:
        return False
    return str(val).strip().upper() == "O"


def init_master_db(master_path: Path) -> int:
    """마스터 엑셀 → master.db 변환."""
    DB_FOLDER.mkdir(parents=True, exist_ok=True)
    db_path = DB_FOLDER / "master.db"
    backup_db_file(db_path)

    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute("DROP TABLE IF EXISTS master_alarm")
        # project-context/RULE.md의 전체 열 구조(A~AQ, 43열)를 그대로 매핑
        cur.execute(
            """
            CREATE TABLE master_alarm (
                no              TEXT,
                category        TEXT,
                department      TEXT,
                customer        TEXT,
                division        TEXT,
                site            TEXT,
                line            TEXT,
                main_process    TEXT,
                sub_process     TEXT,
                unit            TEXT,
                chamber         TEXT,
                sn              TEXT,
                model           TEXT,
                type            TEXT,
                turn_on         TEXT,
                work_date       TEXT,
                work_time_q     TEXT,
                end_time_r      TEXT,
                work_time_s     TEXT,
                staff_count     TEXT,
                man_hour        TEXT,
                major_class     TEXT,
                minor_class     TEXT,
                problem         TEXT,
                cause           TEXT,
                action          TEXT,
                part_type       TEXT,
                part_name       TEXT,
                part_no         TEXT,
                customer_code   TEXT,
                qty             TEXT,
                warranty        TEXT,
                used_days       TEXT,
                charge_type     TEXT,
                cost            TEXT,
                price           TEXT,
                spec            TEXT,
                warranty_out    TEXT,
                prev_visit      TEXT,
                month           TEXT,
                elapsed_days    TEXT,
                initial_defect  TEXT,
                charge_flag     TEXT
            )
            """
        )

        app = None
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            wb = app.books.open(str(master_path), read_only=True)
            sheet = _get_sheet_by_name(wb, MASTER_SHEET_ALARM)
            if sheet is None:
                print(f"오류: 마스터 파일에 시트 '{MASTER_SHEET_ALARM}'이(가) 없습니다.")
                return 0

            used = sheet.used_range
            data = used.value
            if data is None:
                return 0

            # data를 2차원 리스트로 정규화
            if not isinstance(data, list):
                data = [[data]]
            if len(data) > 0 and not isinstance(data[0], list):
                data = [data]

            # 머릿말 2행, 데이터 3행부터 → 인덱스 2부터
            row_count = 0
            for idx in range(2, len(data)):
                row = data[idx]
                if row is None:
                    continue
                # 최소 43열까지 패딩
                if len(row) < 43:
                    row = list(row) + [None] * (43 - len(row))
                values = row[:43]
                if all(v is None or str(v).strip() == "" for v in values):
                    continue
                cur.execute(
                    """
                    INSERT INTO master_alarm VALUES (
                        ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                        ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                        ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                        ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                        ?, ?
                    )
                    """,
                    [str(v) if v is not None else None for v in values],
                )
                row_count += 1

            conn.commit()
            return row_count
        finally:
            if app is not None:
                app.quit()
    finally:
        conn.close()


def init_schedule_db(schedule_path: Path) -> int:
    """스케줄 엑셀 → schedule.db 변환."""
    DB_FOLDER.mkdir(parents=True, exist_ok=True)
    db_path = DB_FOLDER / "schedule.db"
    backup_db_file(db_path)

    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute("DROP TABLE IF EXISTS schedule")
        # project-context/RULE.md의 스케줄 열 매핑 + status 컬럼
        cur.execute(
            """
            CREATE TABLE schedule (
                receipt_date   TEXT,  -- J
                visit_plan     TEXT,  -- K
                visit_1        TEXT,  -- L
                visit_2        TEXT,  -- M
                visit_3        TEXT,  -- N
                reason         TEXT,  -- O
                customer       TEXT,  -- P
                manager        TEXT,  -- Q
                receiver       TEXT,  -- R
                mat_worker     TEXT,  -- S
                work           TEXT,  -- T
                charge_type    TEXT,  -- U
                people         TEXT,  -- V
                process        TEXT,  -- W
                line           TEXT,  -- X
                model          TEXT,  -- Y
                unit           TEXT,  -- Z
                sn             TEXT,  -- AA
                visit_done     TEXT,  -- AB
                process_done   TEXT,  -- AC
                status         TEXT   -- 미완료/처리중/완료/기존완료 등
            )
            """
        )

        app = None
        inserted = 0
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            wb = app.books.open(str(schedule_path), read_only=True)
            if SCHEDULE_SHEET_NAME not in [s.name for s in wb.sheets]:
                print(f"오류: 스케줄 시트에 '{SCHEDULE_SHEET_NAME}'이(가) 없습니다.")
                return 0
            sheet = wb.sheets[SCHEDULE_SHEET_NAME]

            for header_row in HEADER_ROWS:
                start_row = header_row + 1
                end_row = header_row + DATA_ROWS_PER_BLOCK
                for row_num in range(start_row, end_row + 1):
                    val_ab = sheet.range((row_num, COL_VISIT_DONE)).value
                    if not is_visit_done_o(val_ab):
                        continue

                    val_ac = sheet.range((row_num, COL_PROCESS_DONE)).value
                    if val_ac is not None and str(val_ac).strip():
                        status = "기존완료"
                    else:
                        status = "미완료"

                    row_values = [
                        sheet.range((row_num, COL_RECEIPT_DATE)).value,   # J
                        sheet.range((row_num, COL_VISIT_PLAN)).value,     # K
                        sheet.range((row_num, COL_VISIT_1)).value,        # L
                        sheet.range((row_num, COL_VISIT_2)).value,        # M
                        sheet.range((row_num, COL_VISIT_3)).value,        # N
                        sheet.range((row_num, COL_REASON)).value,         # O
                        sheet.range((row_num, COL_CUSTOMER)).value,       # P
                        sheet.range((row_num, COL_MANAGER)).value,        # Q
                        sheet.range((row_num, COL_RECEIVER)).value,       # R
                        sheet.range((row_num, COL_MAT_WORKER)).value,     # S
                        sheet.range((row_num, COL_WORK)).value,           # T
                        sheet.range((row_num, COL_CHARGE_TYPE)).value,    # U
                        sheet.range((row_num, COL_PEOPLE)).value,         # V
                        sheet.range((row_num, COL_PROCESS)).value,        # W
                        sheet.range((row_num, COL_LINE)).value,           # X
                        sheet.range((row_num, COL_MODEL)).value,          # Y
                        sheet.range((row_num, COL_UNIT)).value,           # Z
                        sheet.range((row_num, COL_SN)).value,             # AA
                        val_ab,                                           # AB
                        val_ac,                                           # AC
                        status,
                    ]

                    # 최소한 고객사/업무내용이 없는 완전 빈 행은 건너뜀
                    if all(v is None or str(v).strip() == "" for v in row_values[:-1]):
                        continue

                    cur.execute(
                        """
                        INSERT INTO schedule (
                            receipt_date, visit_plan, visit_1, visit_2, visit_3,
                            reason, customer, manager, receiver, mat_worker,
                            work, charge_type, people, process, line,
                            model, unit, sn, visit_done, process_done, status
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        [str(v) if v is not None else None for v in row_values],
                    )
                    inserted += 1

            conn.commit()
            return inserted
        finally:
            if app is not None:
                app.quit()
    finally:
        conn.close()


def main() -> None:
    if not MASTER_FOLDER.is_dir():
        print(f"오류: 마스터 폴더를 찾을 수 없습니다. {MASTER_FOLDER}")
        return
    if not SCHEDULE_FOLDER.is_dir():
        print(f"오류: 스케줄 폴더를 찾을 수 없습니다. {SCHEDULE_FOLDER}")
        return

    master_path = _get_first_excel_file(MASTER_FOLDER)
    if master_path is None:
        print(f"오류: 마스터 파일을 찾을 수 없습니다. ({MASTER_FOLDER})")
        return

    schedule_path = _get_first_excel_file(SCHEDULE_FOLDER)
    if schedule_path is None:
        print(f"오류: 스케줄 파일을 찾을 수 없습니다. ({SCHEDULE_FOLDER})")
        return

    master_count = init_master_db(master_path)
    schedule_count = init_schedule_db(schedule_path)

    print(f"마스터 {master_count}건, 스케줄 {schedule_count}건 변환 완료")


if __name__ == "__main__":
    main()

