from flask import Flask, request, jsonify, render_template_string
import sqlite3
import pandas as pd
from datetime import datetime
import threading
import time
import os

app = Flask(__name__)

# -------------------------------
# DATABASE PATH
# -------------------------------
DB = r"C:\Users\Administrator\Desktop\Application Tracking\monitor.db"

# Tekla license file
SOURCE_FILE = "**********"
# Output duplicate systems
OUTPUT_FILE = "**************"

# -------------------------------
# CREATE DATABASE TABLE
# -------------------------------
def init_db():

    conn = sqlite3.connect(DB)
    c = conn.cursor()

    c.execute("""
    CREATE TABLE IF NOT EXISTS activity(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        computer TEXT,
        application TEXT,
        title TEXT,
        start_time TEXT,
        last_update TEXT
    )
    """)

    # speed improvement
    c.execute("CREATE INDEX IF NOT EXISTS idx_computer ON activity(computer)")

    conn.commit()
    conn.close()

def clean_old_records():

    conn = sqlite3.connect(DB)
    c = conn.cursor()

    c.execute("""
    DELETE FROM activity
    WHERE last_update < datetime('now','-30 seconds')
    """)

    conn.commit()
    conn.close()

@app.route("/update", methods=["POST"])
def update():

    data = request.json

    computer = data.get("computer")
    app_name = data.get("application")
    title = data.get("title")

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    conn = sqlite3.connect(DB)
    c = conn.cursor()

    c.execute("""
    INSERT INTO activity(computer,application,title,start_time,last_update)
    VALUES(?,?,?,?,?)
    """, (computer, app_name, title, now, now))

    conn.commit()
    conn.close()

    return jsonify({"status": "ok"})


# -------------------------------
# DASHBOARD
# -------------------------------
@app.route("/")
def dashboard():

    conn = sqlite3.connect(DB)
    c = conn.cursor()

    c.execute("""
    SELECT computer,application,title,start_time,last_update
    FROM activity
    ORDER BY last_update DESC
    """)

    rows = c.fetchall()
    conn.close()

    html = """
    <h2>LAN Application Monitor</h2>
    <table border=1 cellpadding=5>
    <tr>
    <th>Computer</th>
    <th>Application</th>
    <th>Window Title</th>
    <th>Start Time</th>
    <th>Last Active</th>
    </tr>

    {% for r in rows %}
    <tr>
    <td>{{r[0]}}</td>
    <td>{{r[1]}}</td>
    <td>{{r[2]}}</td>
    <td>{{r[3]}}</td>
    <td>{{r[4]}}</td>
    </tr>
    {% endfor %}
    </table>
    """

    return render_template_string(html, rows=rows)


# -------------------------------
# COMPARE SYSTEM IDs
# -------------------------------
def compare_system_ids():

    try:

        if not os.path.exists(SOURCE_FILE):
            print("License file not found")
            return

        # Read System IDs from Tekla Excel
        df = pd.read_excel(
            SOURCE_FILE,
            sheet_name="In Use Details",
            usecols="C"
        )

        # remove header row
        license_ids = set(df.iloc[1:, 0].dropna().astype(str))

        # read monitored computers
        conn = sqlite3.connect(DB)

        sql_df = pd.read_sql_query(
            "SELECT DISTINCT computer FROM activity",
            conn
        )

        conn.close()

        monitored_ids = set(sql_df["computer"].astype(str))

        # find systems not present in license
        missing_ids = monitored_ids - license_ids

        # always rewrite the file
        if missing_ids:
            result = pd.DataFrame({"Computer": list(missing_ids)})
        else:
            result = pd.DataFrame(columns=["Computer"])

        result.to_excel(OUTPUT_FILE, index=False)

        print("Comparison completed:", datetime.now())

    except Exception as e:
        print("Comparison error:", e)


# -------------------------------
# LOOP EVERY MINUTE
# -------------------------------
def compare_loop():

    while True:

        compare_system_ids()

        clean_old_records()

        time.sleep(5)


# -------------------------------
# START BACKGROUND THREAD
# -------------------------------
def start_compare_thread():

    t = threading.Thread(target=compare_loop)
    t.daemon = True
    t.start()


# -------------------------------
# MAIN START
# -------------------------------
if __name__ == "__main__":

    init_db()

    start_compare_thread()

    print("Server started")

    app.run(host="0.0.0.0", port=5000)
