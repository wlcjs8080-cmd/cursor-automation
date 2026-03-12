# -*- coding: utf-8 -*-
"""
Step 1: 스케줄 → 레포트 자동 생성 (xlwings)
- AB열(방문완료)=O, AC열(처리완료) 비어 있는 행만 처리 (3월 블록부터)
- 템플릿 복사 후 셀 매핑, 마스터에서 SET UP·이전 점검일 조회, 고객사\REPORT에 저장
- .cursorrules 프로젝트 규칙 준수
"""

import sys
import shutil
import sqlite3
from pathlib import Path
from datetime import datetime, date, timedelta

import xlwings as xw

# ---------------------------------------------------------------------------
# 설정 (.cursorrules 기준)
# ---------------------------------------------------------------------------
BASE_PATH = Path(r"C:\정동교\문서 자동화 TEST\커서 바이브코딩 자동화 관련")
SCHEDULE_FOLDER = BASE_PATH / "스케쥴 시트"
MASTER_FOLDER = BASE_PATH / "마스터 시트"
TEMPLATE_FOLDER = BASE_PATH / "레포트 양식 기준"
CUSTOMER_FOLDER = BASE_PATH / "고객사 폴더"
DB_FOLDER = BASE_PATH / "db"
MASTER_DB_PATH = DB_FOLDER / "master.db"
SCHEDULE_DB_PATH = DB_FOLDER / "schedule.db"
SHEET_NAME = "Sheet1"
TEMPLATE_SHEET_NAME = "PM REPORT (새양식)"
MASTER_SHEET_NAME = "설비 Master Sheet"
MASTER_SHEET_ALARM = "New Alarm 및 사용Part 이력"
HEADER_ROWS = [93, 138, 183, 229, 274, 319, 364]
DATA_ROWS_PER_BLOCK = 41
PROTECT_PASSWORD = "mat2026"

# 스케줄 시트 열 (1-based, .cursorrules 기준)
COL_VISIT_DONE = 28   # AB열 = 방문완료
COL_PROCESS_DONE = 29  # AC열 = 처리완료(작성금지)
COL_1ST_VISIT = 12    # L열 = 1차 방문 일정
COL_2ND_VISIT = 13    # M열 = 2차 방문 일정
COL_3RD_VISIT = 14    # N열 = 3차 이상 마지막 방문일정
COL_CUSTOMER = 16     # P열 = 고객사
COL_RESPONSIBLE = 17  # Q열 = 담당자
COL_MAT = 19          # S열 = MAT 작업자
COL_WORK = 20         # T열 = 업무내용
COL_PROCESS = 23      # W열 = 공정
COL_MODEL = 25        # Y열 = MODEL
COL_UNIT = 26         # Z열 = 설비호기
COL_SERIAL = 27       # AA열 = S/N
# 마스터 설비 Master Sheet
COL_MASTER_SN = 10    # J열 = S/N (매칭 키)
COL_MASTER_SETUP = 13  # M열 = SET UP (레포트 V10)
# 마스터 New Alarm 및 사용Part 이력
COL_ALARM_SN = 12          # L열 = S/N
COL_ALARM_WORK_DATE = 16   # P열 = 작업일자 (이전 점검일 V11용)
COL_ALARM_SETUP = 15       # O열 = TURN ON (SET UP 2순위)

# 설비 Master Sheet: 머릿말 6~7행(병합), 데이터 8행부터
MASTER_DATA_START_ROW = 8
# New Alarm 및 사용Part 이력: 머릿말 2행, 데이터 3행부터
ALARM_DATA_START_ROW = 3


def is_excel_file_open(file_path):
    """해당 엑셀 파일이 이미 Excel에서 열려 있는지 확인."""
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
    """원본 파일을 백업 (파일명_backup_YYYYMMDD_HHMMSS.xlsx)."""
    stem = file_path.stem
    suffix = file_path.suffix
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = file_path.parent / f"{stem}_backup_{stamp}{suffix}"
    shutil.copy2(file_path, backup_path)
    return backup_path


def has_date_value(val):
    """셀에 날짜(또는 날짜로 해석 가능한 값)가 있는지 확인."""
    if val is None:
        return False
    if isinstance(val, (datetime, date)):
        return True
    if isinstance(val, (int, float)):
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
    """AB열(방문완료) 값이 영문 대문자 'O'인지 확인."""
    if val is None:
        return False
    return str(val).strip().upper() == "O"


def get_visit_date_and_yyyymmdd(val_l, val_m, val_n):
    """
    방문일정 N > M > L 우선순위로 날짜 반환.
    반환: (엑셀에 넣을 값, 파일명용 YYYYMMDD 문자열). 없으면 (None, None).
    """
    for val in (val_n, val_m, val_l):
        if not has_date_value(val):
            continue
        if isinstance(val, (datetime, date)):
            d = val if isinstance(val, date) else val.date()
            return val, d.strftime("%Y%m%d")
        if isinstance(val, (int, float)):
            try:
                base = date(1899, 12, 30)
                d = base + timedelta(days=int(val))
                return val, d.strftime("%Y%m%d")
            except Exception:
                continue
        s = str(val).strip()[:10]
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%d/%m/%Y", "%Y.%m.%d"):
            try:
                dt = datetime.strptime(s, fmt)
                return val, dt.strftime("%Y%m%d")
            except ValueError:
                continue
    return None, None


def sanitize_filename(s):
    """파일명에 사용할 수 없은 문자 제거."""
    if s is None:
        return ""
    s = str(s).strip()
    for c in r'\/:*?"<>|':
        s = s.replace(c, "")
    return s[:200]


def find_customer_report_folder(customer_name):
    """
    고객사 폴더 하위에서 고객사명이 포함된 폴더를 찾고, 그 안의 REPORT 폴더 경로 반환.
    없으면 None.
    """
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
            return None  # 고객사 폴더는 찾았지만 REPORT 없음
    return None


def _get_sheet_by_name(book, target_name):
    """시트명이 target_name과 일치하는 시트 반환 (앞뒤 공백 무시). 없으면 None."""
    name_stripped = target_name.strip()
    for s in book.sheets:
        if s.name.strip() == name_stripped:
            return s
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


def _build_setup_dict(master_data, alarm_data):
    """
    S/N → SET UP 딕셔너리 생성.
    1순위: 설비 Master Sheet (S/N=J열 col10, SET UP=M열 col13, 데이터 8행부터)
    2순위: New Alarm 및 사용Part 이력 (S/N=L열 col12, SET UP=AL열 col38, 데이터 3행부터)
    1순위에서 찾으면 2순위는 무시.
    """
    result = {}

    # 2순위 먼저 넣고, 1순위로 덮어쓰기 (1순위 우선)
    # New Alarm 및 사용Part 이력: 데이터 3행 → 인덱스 2
    for row_idx in range(ALARM_DATA_START_ROW - 1, len(alarm_data)):
        row = alarm_data[row_idx]
        if len(row) < max(COL_ALARM_SN, COL_ALARM_SETUP):
            continue
        sn_val = row[COL_ALARM_SN - 1]  # L열
        setup_val = row[COL_ALARM_SETUP - 1]  # AL열
        if sn_val is None:
            continue
        sn_str = str(sn_val).strip()
        if not sn_str:
            continue
        if sn_str not in result and setup_val is not None:
            result[sn_str] = setup_val

    # 1순위: 설비 Master Sheet: 데이터 8행 → 인덱스 7
    for row_idx in range(MASTER_DATA_START_ROW - 1, len(master_data)):
        row = master_data[row_idx]
        if len(row) < max(COL_MASTER_SN, COL_MASTER_SETUP):
            continue
        sn_val = row[COL_MASTER_SN - 1]  # J열
        setup_val = row[COL_MASTER_SETUP - 1]  # M열
        if sn_val is None:
            continue
        sn_str = str(sn_val).strip()
        if not sn_str:
            continue
        if setup_val is not None:
            result[sn_str] = setup_val  # 1순위이므로 덮어쓰기

    return result


def _build_prev_inspection_dict_from_rows(rows):
    """
    S/N → 오늘 이전 가장 최근 작업일자 딕셔너리 생성.
    rows: (sn, work_date) 리스트. work_date는 문자열/숫자/날짜 혼합 가능.
    """
    today = date.today()
    result = {}  # sn_str -> (best_date, best_raw_value)

    for sn_val, work_val in rows:
        if sn_val is None:
            continue
        sn_str = str(sn_val).strip()
        if not sn_str:
            continue
        if not has_date_value(work_val):
            continue

        # 날짜 변환 (이전 구현과 동일 로직)
        if isinstance(work_val, (datetime, date)):
            d = work_val.date() if isinstance(work_val, datetime) else work_val
        elif isinstance(work_val, (int, float)):
            try:
                d = date(1899, 12, 30) + timedelta(days=int(work_val))
            except (ValueError, OverflowError):
                continue
        else:
            s = str(work_val).strip()[:10]
            parsed = None
            for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%d/%m/%Y", "%Y.%m.%d"):
                try:
                    parsed = datetime.strptime(s, fmt).date()
                    break
                except ValueError:
                    continue
            if parsed is None:
                continue
            d = parsed

        if d >= today:
            continue
        if sn_str not in result or d > result[sn_str][0]:
            result[sn_str] = (d, work_val)

    return {k: v[1] for k, v in result.items()}  # sn -> raw value for Excel


def load_master_dicts_from_db():
    """master.db에서 setup_dict, prev_inspection_dict 생성."""
    if not MASTER_DB_PATH.is_file():
        print(f"오류: master.db를 찾을 수 없습니다. ({MASTER_DB_PATH})")
        return {}, {}

    conn = sqlite3.connect(str(MASTER_DB_PATH))
    try:
        cur = conn.cursor()

        # SET UP: S/N → TURN ON(O열)
        setup_dict = {}
        cur.execute("SELECT sn, turn_on FROM master_alarm WHERE sn IS NOT NULL AND turn_on IS NOT NULL")
        for sn_val, setup_val in cur.fetchall():
            sn_str = str(sn_val).strip()
            if not sn_str:
                continue
            setup_dict[sn_str] = setup_val

        # 이전 점검일: S/N → 오늘 이전 가장 최근 작업일자(work_date)
        cur.execute("SELECT sn, work_date FROM master_alarm WHERE sn IS NOT NULL AND work_date IS NOT NULL")
        rows = cur.fetchall()
        prev_inspection_dict = _build_prev_inspection_dict_from_rows(rows)

        return setup_dict, prev_inspection_dict
    finally:
        conn.close()


def make_report_filename(customer, yyyymmdd, model, unit, serial, work):
    """파일명: 고객사_날짜_설비타입_호기_(시리얼번호)_작업내용 件.xlsx (설비타입=MODEL, 호기=설비호기)"""
    c = sanitize_filename(customer)
    e = sanitize_filename(model)
    u = sanitize_filename(unit)
    s = sanitize_filename(serial)
    w = sanitize_filename(work)
    return f"{c}_{yyyymmdd}_{e}_{u}_({s})_{w} 件.xlsx"


def process_one_row(app, row, setup_dict, prev_inspection_dict, template_path):
    """
    한 행 처리: 검증 → 방문일/고객사/폴더 → 딕셔너리에서 SET UP·이전 점검일 조회 → 템플릿 복사 → 셀 채우기 → 저장.
    성공 시 True, 실패 시 오류 메시지 반환(문자열).
    row: schedule.db에서 읽은 dict(row).
    """
    row_id = row.get("row_id")

    # DB에서 읽은 값 매핑
    val_ab = row.get("visit_done")
    val_ac = row.get("process_done")
    val_l = row.get("visit_1")
    val_m = row.get("visit_2")
    val_n = row.get("visit_3")
    customer = row.get("customer")
    responsible = row.get("manager")
    mat_worker = row.get("mat_worker")
    model = row.get("model")
    serial = row.get("sn")
    process = row.get("process")
    unit = row.get("unit")
    work = row.get("work")

    # 방문완료/기존완료 여부는 schedule.db 생성 시 필터링되지만, 방어적으로 한 번 더 체크
    if not is_visit_done_o(val_ab):
        return None  # 처리 대상 아님
    if val_ac is not None and str(val_ac).strip():
        return None  # 이미 처리됨 (기존완료 등)

    # 방문일정 없음
    visit_val, yyyymmdd = get_visit_date_and_yyyymmdd(val_l, val_m, val_n)
    if visit_val is None or yyyymmdd is None:
        return f"행 {row_id}: 방문일정 없음 오류"

    # 고객사명 누락
    if not customer or not str(customer).strip():
        return f"행 {row_id}: 고객사명 누락"

    # 필수값 누락 (MODEL=Y, S/N=AA, 업무내용=T)
    if not model or not str(model).strip():
        return f"행 {row_id}: MODEL(Y열) 누락"
    if not serial or not str(serial).strip():
        return f"행 {row_id}: S/N(AA열) 누락"
    if not work or not str(work).strip():
        return f"행 {row_id}: 업무내용(T열) 누락"

    # 고객사 폴더 > REPORT 찾기
    report_dir = find_customer_report_folder(customer)
    if report_dir is None:
        return f"행 {row_id}: 고객사 폴더 또는 REPORT 폴더를 찾을 수 없음 (고객사: {customer})"

    # 딕셔너리에서 SET UP(V10) 조회
    sn_str = str(serial).strip()
    setup_val = setup_dict.get(sn_str)
    if setup_val is None:
        return f"행 {row_id}: MASTER SHEET에 S/N이 존재하지 않습니다 ({serial})"

    # 딕셔너리에서 이전 점검일(V11) 조회 (없어도 진행, 빈칸 허용)
    previous_inspection_val = prev_inspection_dict.get(sn_str)

    # 파일명 생성 (고객사_날짜_MODEL_설비호기_(S/N)_작업내용)
    filename = make_report_filename(customer, yyyymmdd, model, unit, serial, work)
    save_path = report_dir / filename

    # 레포트 파일명 중복 체크
    if save_path.exists():
        return f"행 {row_id}: 같은 파일명의 레포트가 존재합니다"

    # 템플릿 파일 복사 후 열어서 셀 채우기
    try:
        shutil.copy2(template_path, save_path)
    except Exception as e:
        return f"행 {row_id}: 템플릿 복사 실패 - {e}"

    report_wb = app.books.open(str(save_path))
    try:
        if TEMPLATE_SHEET_NAME not in [s.name for s in report_wb.sheets]:
            report_wb.close()
            try:
                save_path.unlink(missing_ok=True)
            except Exception:
                pass
            return f"행 {row_num}: 템플릿에 시트 '{TEMPLATE_SHEET_NAME}' 없음"
        sheet = report_wb.sheets[TEMPLATE_SHEET_NAME]
        # 레포트 템플릿 셀 매핑 (.cursorrules 기준)
        sheet.range("E10").value = customer   # P열(고객사)
        sheet.range("E11").value = responsible  # Q열(담당자)
        sheet.range("E12").value = mat_worker  # S열(MAT 작업자)
        sheet.range("N10").value = model       # Y열(MODEL)
        sheet.range("N11").value = serial       # AA열(S/N)
        sheet.range("N12").value = process      # W열(공정)
        sheet.range("V8").value = visit_val     # 방문일정 N>M>L 우선
        sheet.range("V10").value = setup_val   # 마스터 SET UP
        sheet.range("V11").value = previous_inspection_val  # 이전 점검일
        # 경과일(V12 등)은 엑셀 수식 자동 계산이므로 코드에서 입력하지 않음
        report_wb.save()
    finally:
        report_wb.close()

    return True


def main():
    if not TEMPLATE_FOLDER.is_dir():
        print(f"오류: 템플릿 폴더를 찾을 수 없습니다. {TEMPLATE_FOLDER}")
        return
    if not DB_FOLDER.is_dir():
        print(f"오류: DB 폴더를 찾을 수 없습니다. {DB_FOLDER}")
        return
    if not MASTER_DB_PATH.is_file():
        print(f"오류: master.db를 찾을 수 없습니다. ({MASTER_DB_PATH})")
        return
    if not SCHEDULE_DB_PATH.is_file():
        print(f"오류: schedule.db를 찾을 수 없습니다. ({SCHEDULE_DB_PATH})")
        return

    template_files = list(TEMPLATE_FOLDER.glob("*.xlsx")) + list(TEMPLATE_FOLDER.glob("*.xlsm"))
    template_files = [p for p in template_files if "_backup_" not in p.name and not p.name.startswith("~$")]
    if not template_files:
        print(f"템플릿 파일이 없습니다. ({TEMPLATE_FOLDER})")
        return
    template_path = template_files[0]

    # master.db에서 SET UP / 이전 점검일 딕셔너리 생성
    setup_dict, prev_inspection_dict = load_master_dicts_from_db()
    print(f"마스터 로딩 완료: SET UP {len(setup_dict)}건, 이전 점검일 {len(prev_inspection_dict)}건")

    # schedule.db에서 status="미완료" 행 조회
    conn = sqlite3.connect(str(SCHEDULE_DB_PATH))
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute(
        """
        SELECT
            rowid AS row_id,
            receipt_date,
            visit_plan,
            visit_1,
            visit_2,
            visit_3,
            reason,
            customer,
            manager,
            receiver,
            mat_worker,
            work,
            charge_type,
            people,
            process,
            line,
            model,
            unit,
            sn,
            visit_done,
            process_done,
            status
        FROM schedule
        WHERE status = '미완료'
        """
    )
    rows = cur.fetchall()

    if not rows:
        print("status='미완료' 인 스케줄 행이 없습니다.")
        conn.close()
        return

    app = None
    success_count = 0
    skip_count = 0

    try:
        app = xw.App(visible=False)
        app.display_alerts = False

        for row in rows:
            row_id = row["row_id"]
            try:
                # 처리 시작 시 status="처리중"으로 변경 (동시 중복 차단)
                with conn:
                    cur.execute(
                        "UPDATE schedule SET status = '처리중' WHERE rowid = ? AND status = '미완료'",
                        (row_id,),
                    )
                    if cur.rowcount == 0:
                        # 다른 프로세스가 먼저 집어간 행
                        continue

                row_dict = dict(row)

                result = process_one_row(app, row_dict, setup_dict, prev_inspection_dict, template_path)
                if result is None:
                    # 처리 대상 아님
                    continue
                if result is True:
                    with conn:
                        cur.execute(
                            "UPDATE schedule SET status = '완료' WHERE rowid = ?",
                            (row_id,),
                        )
                    success_count += 1
                    print(f"행 {row_id}: 레포트 생성 완료")
                else:
                    # 오류 문자열 (검증 실패 등)
                    with conn:
                        cur.execute(
                            "UPDATE schedule SET status = '미완료' WHERE rowid = ?",
                            (row_id,),
                        )
                    skip_count += 1
                    print(result)

            except Exception:
                # 에러 발생 시: 레포트 파일 삭제 → status를 "미완료"로 되돌리기
                customer = row.get("customer")
                sn = row.get("sn")
                print(
                    f"[에러] 행 {row_id} (고객사: {customer}, S/N: {sn}) 레포트 생성 중지. "
                    f"스케줄 시트 작성 내용은 유지됨. 문제 확인 후 Step 1 재실행 필요."
                )

                # 레포트 파일 삭제 시도
                try:
                    visit_val, yyyymmdd = get_visit_date_and_yyyymmdd(
                        row.get("visit_1"), row.get("visit_2"), row.get("visit_3")
                    )
                    if visit_val is not None and yyyymmdd is not None:
                        report_dir = find_customer_report_folder(customer)
                        if report_dir is not None:
                            filename = make_report_filename(
                                customer,
                                yyyymmdd,
                                row.get("model"),
                                row.get("unit"),
                                sn,
                                row.get("work"),
                            )
                            save_path = report_dir / filename
                            try:
                                save_path.unlink(missing_ok=True)
                            except Exception:
                                pass
                except Exception:
                    pass

                with conn:
                    cur.execute(
                        "UPDATE schedule SET status = '미완료' WHERE rowid = ?",
                        (row_id,),
                    )
                skip_count += 1

        print("-" * 50)
        print(f"총 {success_count}건 레포트 생성, {skip_count}건 건너뜀")
    finally:
        if app is not None:
            app.quit()
        conn.close()


if __name__ == "__main__":
    main()
