# -*- coding: utf-8 -*-
import sqlite3
db_path = r"c:\정동교\문서 자동화 TEST\커서 바이브코딩 자동화 관련\db\schedule.db"
conn = sqlite3.connect(db_path)
cur = conn.cursor()
cur.execute("SELECT COUNT(*) FROM schedule WHERE status='미완료'")
print(cur.fetchone()[0])
conn.close()
