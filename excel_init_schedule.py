# -*- coding: utf-8 -*-
"""
엑셀 스케줄 시트 초기화 프로그램 (xlwings 버전)
- 스케쥴 시트 폴더 내 엑셀 파일의 Sheet1에서 기존완료 처리 및 AC열 시트 보호
- .cursorrules 프로젝트 규칙 준수: xlwings만 사용, 백업 자동 생성, 서식 보존
"""

import sys
import shutil
from pathlib import Path
from datetime import datetime, date

import xlwings as xw

# ---------------------------------------------------------------------------
# 설정 (.cursorrules 기준)
# ---------------------------------------------------------------------------
BASE_PATH = Path(r"C:\정동교\문서 자동화 TEST\커서 바이브코딩 자동화 관련")
SCHEDULE_FOLDER = BASE_PATH / "스케쥴 시트"
SHEET_NAME = "Sheet1"
HEADER_ROWS = [93, 138, 183, 229, 274, 319, 364]
DATA_ROWS_PER_BLOCK = 41
PROTECT_PASSWORD = "mat2026"

# 열 인덱스 (1-based, .cursorrules 기준)
COL_VISIT_DONE = 28   # AB열 = 방문완료
COL_1ST_VISIT = 12   # L열 = 1차 방문 일정
COL_2ND_VISIT = 13   # M열 = 2차 방문 일정
COL_3RD_VISIT = 14   # N열 = 3차 이상 마지막 방문일정
COL_PROCESS_DONE = 29  # AC열 = 처리완료(작성금지)


def is_excel_file_open(file_path):
    """
    해당 엑셀 파일이 이미 Excel에서 열려 있는지 확인.
    열려 있으면 True, 아니면 False.
    """
    try:
        import win32com.client
        app = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        return False
    try:
        target_abs = Path(file_path).resolve()
        for wb in app.Workbooks:
            try:
                wb_path = Path(wb.FullName).resolve()
                if wb_path == target_abs:
                    return True
            except Exception:
                continue
    except Exception:
        pass
    return False


def backup_file(file_path):
    """원본 파일을 백업 (파일명_backup_YYYYMMDD_HHMMSS.xlsx)"""
    stem = file_path.stem
    suffix = file_path.suffix
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = file_path.parent / f"{stem}_backup_{stamp}{suffix}"
    shutil.copy2(file_path, backup_path)
    return backup_path


def has_date_value(val):
    """셀에 날짜(또는 날짜로 해석 가능한 값)가 있는지 확인"""
    if val is None:
        return False
    if isinstance(val, (datetime, date)):
        return True
    if isinstance(val, (int, float)):
        # 엑셀 날짜는 숫자(serial)로 저장될 수 있음
        if 1 < val < 2958466:
            return True
        return False
    s = str(val).strip()
    if not s:
        return False
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%d/%m/%Y", "%Y.%m.%d"):
        try:
            datetime.strptime(s[:10], fmt)
            return True
        except ValueError:
            continue
    return False


def is_visit_done_o(val):
    """AB열(방문완료) 값이 영문 대문자 'O'인지 확인"""
    if val is None:
        return False
    return str(val).strip().upper() == "O"


def process_file(app, file_path):
    """
    엑셀 파일 하나 처리.
    반환: (기존완료 처리 수, 방문일정 없어서 건너뜀 수)
    """
    wb = app.books.open(str(file_path))
    try:
        if SHEET_NAME not in [s.name for s in wb.sheets]:
            raise ValueError(f"시트 '{SHEET_NAME}'이(가) 없습니다.")

        sheet = wb.sheets[SHEET_NAME]
        total_marked = 0
        total_skipped = 0

        # 시트가 이미 보호되어 있으면 해제 (비밀번호 있을 수 있음)
        try:
            sheet.api.Unprotect(PROTECT_PASSWORD)
        except Exception:
            try:
                sheet.api.Unprotect()
            except Exception:
                pass

        # 전체 셀 잠금 해제 후 AC열만 잠금 (데이터 범위 고려해 충분한 행까지)
        max_row = HEADER_ROWS[-1] + DATA_ROWS_PER_BLOCK + 50
        try:
            sheet.range((1, 1), (max_row, 29)).api.Locked = False
            sheet.range((1, COL_PROCESS_DONE), (max_row, COL_PROCESS_DONE)).api.Locked = True
        except Exception:
            pass

        for header_row in [r for r in HEADER_ROWS if r >= 183]:
            start_row = header_row + 1
            end_row = header_row + DATA_ROWS_PER_BLOCK
            for row_num in range(start_row, end_row + 1):
                try:
                    # xlwings는 1-based 인덱스
                    val_ab = sheet.range((row_num, COL_VISIT_DONE)).value
                    val_ac = sheet.range((row_num, COL_PROCESS_DONE)).value
                    val_l = sheet.range((row_num, COL_1ST_VISIT)).value
                    val_m = sheet.range((row_num, COL_2ND_VISIT)).value
                    val_n = sheet.range((row_num, COL_3RD_VISIT)).value

                    if not is_visit_done_o(val_ab):
                        continue
                    if val_ac is not None and str(val_ac).strip():
                        continue
                    has_any_date = (
                        has_date_value(val_l)
                        or has_date_value(val_m)
                        or has_date_value(val_n)
                    )
                    if not has_any_date:
                        total_skipped += 1
                        continue
                    sheet.range((row_num, COL_PROCESS_DONE)).value = "기존완료"  # AC열
                    total_marked += 1
                except Exception as e:
                    raise RuntimeError(f"행 {row_num} 처리 중 오류: {e}") from e

        # AC열 잠금 재적용 후 시트 보호
        try:
            sheet.range((1, COL_PROCESS_DONE), (max_row, COL_PROCESS_DONE)).api.Locked = True
            sheet.api.Protect(Password=PROTECT_PASSWORD)
        except Exception as e:
            raise RuntimeError(f"시트 보호 설정 중 오류: {e}") from e

        wb.save()
        return total_marked, total_skipped
    finally:
        wb.close()


def main():
    if not SCHEDULE_FOLDER.is_dir():
        print(f"오류: 폴더를 찾을 수 없습니다. {SCHEDULE_FOLDER}")
        return

    excel_files = list(SCHEDULE_FOLDER.glob("*.xlsx")) + list(SCHEDULE_FOLDER.glob("*.xlsm"))
    excel_files = [p for p in excel_files if "_backup_" not in p.name and not p.name.startswith("~$")]
    if not excel_files:
        print(f"처리할 엑셀 파일이 없습니다. ({SCHEDULE_FOLDER})")
        return

    # 실행 전: 처리 대상 파일 중 열려 있는 파일이 있으면 종료
    for path in excel_files:
        if is_excel_file_open(path):
            print("엑셀 파일을 닫고 다시 실행하세요.")
            sys.exit(1)

    # 실행 전: 원본 파일 자동 백업
    backups = []
    for path in excel_files:
        try:
            backup_path = backup_file(path)
            backups.append(backup_path)
            print(f"백업 완료: {backup_path.name}")
        except Exception as e:
            print(f"백업 실패 ({path.name}): {e}")
            sys.exit(1)

    app = None
    try:
        app = xw.App(visible=False)
        app.display_alerts = False
        total_marked_all = 0
        total_skipped_all = 0
        for path in excel_files:
            try:
                marked, skipped = process_file(app, path)
                total_marked_all += marked
                total_skipped_all += skipped
                print(f"[{path.name}] 기존완료 {marked}행, 건너뜀 {skipped}행")
            except Exception as e:
                print(f"[{path.name}] 오류: {e}")
        print("-" * 50)
        print(f"총 {total_marked_all}행 기존완료 처리, {total_skipped_all}행 방문일정 없어서 건너뜀")
    finally:
        if app is not None:
            app.quit()


if __name__ == "__main__":
    main()
