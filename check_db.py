# -*- coding: utf-8 -*-
import sqlite3
import os
db_dir = r"c:\정동교\문서 자동화 TEST\커서 바이브코딩 자동화 관련\db"
master_path = os.path.join(db_dir, "master.db")
schedule_path = os.path.join(db_dir, "schedule.db")

def check_db(path, label):
    if not os.path.exists(path):
        print(f"=== {label} ===\n파일 없음")
        return
    conn = sqlite3.connect(path)
    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
    tables = cur.fetchall()
    print(f"=== {label} ===")
    print("테이블 목록:", [t[0] for t in tables])
    for t in tables:
        name = t[0]
        cur.execute(f"PRAGMA table_info({name})")
        cols = cur.fetchall()
        cur.execute(f"SELECT COUNT(*) FROM [{name}]")
        rows = cur.fetchone()[0]
        print(f"  테이블: {name} | 행 수: {rows} | 컬럼 수: {len(cols)}")
    conn.close()

check_db(master_path, "master.db")
check_db(schedule_path, "schedule.db")

# 3월 블록(183) 스케줄 확인
if os.path.exists(schedule_path):
    conn = sqlite3.connect(schedule_path)
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM schedule")
    total = cur.fetchone()[0]
    print(f"\nschedule 테이블 총 행 수: {total}")
    conn.close()
