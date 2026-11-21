from flask import Flask, render_template, request, redirect, url_for, session, send_file, flash
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime, date, timedelta
import sqlite3
import os
import io
import json
import re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.environ.get("DATABASE_PATH", os.path.join(BASE_DIR, "marketing.db"))
UPLOAD_DIR = os.environ.get("UPLOAD_DIR", os.path.join(BASE_DIR, "uploads"))
db_dir = os.path.dirname(DB_PATH)
if db_dir:
    os.makedirs(db_dir, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "marketing-flow-ai-change-this-in-production")

STATUSES = ["Шинэ", "Шалгаж байна", "Хийгдэж байна", "Хяналтанд", "Дууссан", "Цуцлагдсан"]
WORK_TYPES = [
    "Ажлын захиалга", "Social Media", "График дизайн", "Видео продакшн",
    "Хэвлэлийн эх", "Урамшуулал", "Кампанит ажил", "Copywriting",
    "Брэнд менежмент", "Дижитал сурталчилгаа", "PR / хамтын ажиллагаа",
    "Салбарын маркетинг", "Судалгаа / тайлан", "Төсөв / гэрээ",
    "Эвент", "Веб / автоматжуулалт",
]


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'user',
            branch TEXT,
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS plans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_file TEXT NOT NULL,
            sheet_name TEXT NOT NULL,
            row_number INTEGER,
            month TEXT,
            title TEXT NOT NULL,
            details TEXT,
            branch TEXT,
            assignee TEXT,
            deadline TEXT,
            budget REAL DEFAULT 0,
            status TEXT DEFAULT 'Төлөвлөгөө',
            raw_data TEXT,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            title TEXT NOT NULL,
            work_type TEXT NOT NULL,
            branch TEXT,
            description TEXT,
            priority TEXT DEFAULT 'Энгийн',
            status TEXT DEFAULT 'Шинэ',
            assignee TEXT,
            deadline TEXT,
            estimated_hours REAL,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            completed_at TEXT,
            FOREIGN KEY(user_id) REFERENCES users(id)
        );
    """)
    if conn.execute("SELECT COUNT(*) FROM users").fetchone()[0] == 0:
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        conn.execute(
            "INSERT INTO users (name,email,password,role,branch,created_at) VALUES (?,?,?,?,?,?)",
            ("Администратор", "admin", generate_password_hash("admin123"), "admin", "Төв", now),
        )
        conn.execute(
            "INSERT INTO users (name,email,password,role,branch,created_at) VALUES (?,?,?,?,?,?)",
            ("Central Mall", "central@yoshinoya.mn", generate_password_hash("user123"), "user", "Central Mall", now),
        )
    conn.commit()
    conn.close()


init_db()


@app.context_processor
def inject_globals():
    return {"work_types": WORK_TYPES, "statuses": STATUSES}


def login_required(role=None):
    def decorator(fn):
        def wrapped(*args, **kwargs):
            if "user_id" not in session:
                return redirect(url_for("login"))
            if role and session.get("role") != role:
                return redirect(url_for("dashboard"))
            return fn(*args, **kwargs)
        wrapped.__name__ = fn.__name__
        return wrapped
    return decorator


def parse_date(value):
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.isoformat()
    if value:
        text = str(value).strip()
        for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%d.%m.%Y"):
            try:
                return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
            except ValueError:
                pass
    return ""


def estimate_hours(work_type, priority, description):
    base = {
        "Ажлын захиалга": 4, "Social Media": 8, "График дизайн": 12,
        "Видео продакшн": 24, "Хэвлэлийн эх": 16, "Урамшуулал": 12,
        "Кампанит ажил": 40, "Copywriting": 6, "Брэнд менежмент": 16,
        "Дижитал сурталчилгаа": 16, "PR / хамтын ажиллагаа": 20,
        "Салбарын маркетинг": 16, "Судалгаа / тайлан": 12,
        "Төсөв / гэрээ": 8, "Эвент": 32, "Веб / автоматжуулалт": 24,
    }.get(work_type, 8)
    if priority == "Яаралтай":
        base *= 0.8
    elif priority == "Өндөр":
        base *= 0.9
    extra = min(16, max(0, len(description or "") // 300 * 2))
    return round(base + extra, 1)


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        conn = get_db()
        user = conn.execute("SELECT * FROM users WHERE email=? AND active=1", (email,)).fetchone()
        conn.close()
        if user and check_password_hash(user["password"], password):
            session.clear()
            session.update(user_id=user["id"], role=user["role"], name=user["name"], email=user["email"])
            return redirect(url_for("dashboard"))
        flash("И-мэйл эсвэл нууц үг буруу байна.", "error")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
@login_required()
def dashboard():
    conn = get_db()
    if session["role"] == "admin":
        counts = {}
        for row in conn.execute("SELECT status, COUNT(*) AS count FROM orders GROUP BY status").fetchall():
            counts[row["status"]] = row["count"]
        total_orders = conn.execute("SELECT COUNT(*) FROM orders").fetchone()[0]
        done_orders = conn.execute("SELECT COUNT(*) FROM orders WHERE status='Дууссан'").fetchone()[0]
        overdue = conn.execute("""
            SELECT COUNT(*) FROM orders
            WHERE status != 'Дууссан' AND deadline != '' AND deadline < date('now')
        """).fetchone()[0]
        plan_count = conn.execute("SELECT COUNT(*) FROM plans").fetchone()[0]
        recent = conn.execute("""
            SELECT o.*, u.name AS customer_name
            FROM orders o LEFT JOIN users u ON u.id=o.user_id
            ORDER BY o.created_at DESC LIMIT 8
        """).fetchall()
        branch_rows = conn.execute("""
            SELECT COALESCE(NULLIF(branch,''),'Тодорхойгүй') AS label, COUNT(*) AS count
            FROM orders GROUP BY branch ORDER BY count DESC
        """).fetchall()
        data = dict(total=total_orders, done=done_orders, overdue=overdue, plan=plan_count,
                    counts=counts, recent=recent, branches=branch_rows)
    else:
        recent = conn.execute("""
            SELECT * FROM orders WHERE user_id=? ORDER BY created_at DESC
        """, (session["user_id"],)).fetchall()
        total = len(recent)
        done = sum(1 for row in recent if row["status"] == "Дууссан")
        active = total - done
        data = dict(recent=recent, total=total, done=done, active=active)
    conn.close()
    return render_template("dashboard.html", data=data)


@app.route("/orders", methods=["GET", "POST"])
@login_required()
def orders():
    conn = get_db()
    if request.method == "POST":
        title = request.form["title"].strip()
        work_type = request.form["work_type"]
        branch = request.form.get("branch", "").strip()
        description = request.form.get("description", "").strip()
        priority = request.form.get("priority", "Энгийн")
        assignee = request.form.get("assignee", "").strip()
        deadline = request.form.get("deadline", "")
        estimated = request.form.get("estimated_hours")
        if session["role"] == "user":
            user_id = session["user_id"]
            branch = conn.execute("SELECT branch FROM users WHERE id=?", (user_id,)).fetchone()["branch"] or branch
            assignee = ""
            status = "Шинэ"
        else:
            user_id = request.form.get("user_id") or None
            status = request.form.get("status", "Шинэ")
        if not estimated:
            estimated = estimate_hours(work_type, priority, description)
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        if not deadline:
            deadline = (datetime.now() + timedelta(days=max(2, round(float(estimated) / 8)))).strftime("%Y-%m-%d")
        conn.execute("""
            INSERT INTO orders (user_id,title,work_type,branch,description,priority,status,assignee,deadline,
                                estimated_hours,created_at,updated_at)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
        """, (user_id, title, work_type, branch, description, priority, status, assignee, deadline,
              float(estimated), now, now))
        conn.commit()
        flash("Захиалга амжилттай бүртгэгдлээ.", "success")
        conn.close()
        return redirect(url_for("orders"))

    where, params = [], []
    if session["role"] != "admin":
        where.append("o.user_id=?")
        params.append(session["user_id"])
    for field, column in (("branch", "o.branch"), ("status", "o.status"),
                          ("work_type", "o.work_type"), ("assignee", "o.assignee")):
        value = request.args.get(field)
        if value:
            where.append(f"{column}=?")
            params.append(value)
    search = request.args.get("q", "").strip()
    if search:
        where.append("(o.title LIKE ? OR o.description LIKE ?)")
        params += [f"%{search}%", f"%{search}%"]
    sql = """
        SELECT o.*, u.name AS customer_name, u.email AS customer_email
        FROM orders o LEFT JOIN users u ON u.id=o.user_id
    """
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY o.created_at DESC"
    rows = conn.execute(sql, params).fetchall()
    users = conn.execute("SELECT * FROM users WHERE role='user' AND active=1 ORDER BY name").fetchall()
    branches = conn.execute("SELECT DISTINCT branch FROM users WHERE branch IS NOT NULL AND branch!='' UNION "
                            "SELECT DISTINCT branch FROM orders WHERE branch IS NOT NULL AND branch!=''").fetchall()
    assignees = conn.execute("SELECT DISTINCT assignee FROM orders WHERE assignee IS NOT NULL AND assignee!=''").fetchall()
    conn.close()
    return render_template("orders.html", rows=rows, users=users, branches=branches, assignees=assignees)


@app.route("/orders/<int:order_id>/update", methods=["POST"])
@login_required(role="admin")
def update_order(order_id):
    status = request.form["status"]
    assignee = request.form.get("assignee", "").strip()
    deadline = request.form.get("deadline", "")
    estimated = request.form.get("estimated_hours", "")
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    completed = now if status == "Дууссан" else None
    conn = get_db()
    conn.execute("""
        UPDATE orders SET status=?, assignee=?, deadline=?,
        estimated_hours=COALESCE(?,estimated_hours), updated_at=?, completed_at=?
        WHERE id=?
    """, (status, assignee, deadline, float(estimated) if estimated else None, now, completed, order_id))
    conn.commit()
    conn.close()
    flash("Захиалгын төлөв шинэчлэгдлээ.", "success")
    return redirect(url_for("orders"))


@app.route("/plans", methods=["GET", "POST"])
@login_required(role="admin")
def plans():
    conn = get_db()
    if request.method == "POST":
        file = request.files.get("excel_file")
        if not file or not file.filename:
            flash("Excel файл сонгоно уу.", "error")
        else:
            filename = secure_filename(file.filename)
            path = os.path.join(UPLOAD_DIR, filename)
            file.save(path)
            imported = import_excel(path, filename)
            flash(f"{imported} төлөвлөгөөний мөр SQLite-д бүртгэгдлээ.", "success")
    rows = conn.execute("SELECT * FROM plans ORDER BY created_at DESC, id DESC").fetchall()
    conn.close()
    return render_template("plans.html", rows=rows)


def import_excel(path, filename):
    wb = load_workbook(path, data_only=True)
    conn = get_db()
    imported = 0
    month_match = re.search(r"(\d{4})?[.\- ]*(\d{1,2})[-рсар]+", filename)
    default_month = datetime.now().strftime("%Y-%m")
    for ws in wb.worksheets:
        header_row = None
        for r in range(1, min(ws.max_row, 30) + 1):
            values = [ws.cell(r, c).value for c in range(1, min(ws.max_column, 20) + 1)]
            text = " ".join(str(v).lower() for v in values if v is not None)
            if any(k in text for k in ("агуулга", "хийгдэх", "захиалга", "хэлбэр", "бичлэг")):
                header_row = r
                break
        if header_row:
            headers = []
            for c in range(1, ws.max_column + 1):
                val = ws.cell(header_row, c).value
                headers.append(str(val).strip() if val is not None else "")
            for r in range(header_row + 1, ws.max_row + 1):
                vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
                if not any(v is not None and str(v).strip() for v in vals):
                    continue
                row = dict(zip(headers, vals))
                title = next((row[h] for h in headers if h and any(k in h.lower() for k in ("агуулга", "бичлэг", "захиалга")) and row[h]), None)
                if title is None:
                    title = next((v for v in vals if v is not None and not isinstance(v, (int, float))), None)
                if title is None:
                    continue
                details = " | ".join(f"{h}: {v}" for h, v in zip(headers, vals) if h and v is not None)
                branch = next((row[h] for h in headers if h and "салбар" in h.lower() and row[h]), "")
                assignee = next((row[h] for h in headers if h and "хариу" in h.lower() and row[h]), "")
                budget = next((row[h] for h in headers if h and "төсөв" in h.lower() and row[h]), 0)
                deadline = ""
                for h in headers:
                    if h and any(k in h.lower() for k in ("өдөр", "хугацаа", "deadline")):
                        deadline = parse_date(row[h])
                        if deadline:
                            break
                try:
                    budget = float(budget or 0)
                except (TypeError, ValueError):
                    budget = 0
                conn.execute("""
                    INSERT INTO plans (source_file,sheet_name,row_number,month,title,details,branch,assignee,
                                       deadline,budget,status,raw_data,created_at)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (filename, ws.title, r, default_month, str(title), details, str(branch or ""),
                      str(assignee or ""), deadline, budget, "Төлөвлөгөө",
                      json.dumps([str(v) if v is not None else "" for v in vals], ensure_ascii=False),
                      datetime.now().strftime("%Y-%m-%d %H:%M")))
                imported += 1
        else:
            for r in range(1, ws.max_row + 1):
                vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
                texts = [str(v).strip() for v in vals if v is not None and str(v).strip()]
                if not texts:
                    continue
                title = " | ".join(texts[:4])
                if len(title) < 3:
                    continue
                conn.execute("""
                    INSERT INTO plans (source_file,sheet_name,row_number,month,title,details,branch,assignee,
                                       deadline,budget,status,raw_data,created_at)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (filename, ws.title, r, default_month, title[:300], " | ".join(texts),
                      "", "", "", 0, "Төлөвлөгөө",
                      json.dumps([str(v) if v is not None else "" for v in vals], ensure_ascii=False),
                      datetime.now().strftime("%Y-%m-%d %H:%M")))
                imported += 1
    conn.commit()
    conn.close()
    return imported


@app.route("/customers", methods=["GET", "POST"])
@login_required(role="admin")
def customers():
    conn = get_db()
    if request.method == "POST":
        name = request.form["name"].strip()
        email = request.form["email"].strip().lower()
        branch = request.form.get("branch", "").strip()
        password = request.form.get("password", "user123")
        try:
            conn.execute(
                "INSERT INTO users (name,email,password,role,branch,created_at) VALUES (?,?,?,?,?,?)",
                (name, email, generate_password_hash(password), "user", branch,
                 datetime.now().strftime("%Y-%m-%d %H:%M")),
            )
            conn.commit()
            flash("Захиалагч бүртгэгдлээ.", "success")
        except sqlite3.IntegrityError:
            flash("Энэ company email аль хэдийн бүртгэлтэй байна.", "error")
    users = conn.execute("SELECT * FROM users WHERE role='user' ORDER BY created_at DESC").fetchall()
    conn.close()
    return render_template("customers.html", users=users)


@app.route("/report")
@login_required()
def report():
    conn = get_db()
    orders = conn.execute("""
        SELECT o.*, u.name AS customer_name, u.email AS customer_email
        FROM orders o LEFT JOIN users u ON u.id=o.user_id
        ORDER BY o.created_at DESC
    """).fetchall()
    plans = conn.execute("SELECT * FROM plans ORDER BY id DESC").fetchall()
    conn.close()
    wb = Workbook()
    ws = wb.active
    ws.title = "Төлөвлөгөө ба захиалга"
    headers = ["Төрөл", "№", "Нэр", "Ангилал / Салбар", "Хариуцагч", "Төлөв",
               "Хугацаа", "Төсөв / цаг", "Тайлбар", "Үүсгэсэн"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="5D3FD3")
        cell.alignment = Alignment(horizontal="center")
    for p in plans:
        ws.append(["Төлөвлөгөө", p["id"], p["title"], p["branch"] or p["sheet_name"],
                   p["assignee"], p["status"], p["deadline"], p["budget"], p["details"], p["created_at"]])
    for o in orders:
        ws.append(["Захиалга", o["id"], o["title"], o["branch"] or o["work_type"],
                   o["assignee"] or o["customer_name"], o["status"], o["deadline"],
                   o["estimated_hours"], o["description"], o["created_at"]])
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 22
    ws.column_dimensions["I"].width = 55
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True,
                     download_name=f"marketing_report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/report/view")
@login_required(role="admin")
def report_view():
    conn = get_db()
    plans = conn.execute("SELECT * FROM plans ORDER BY id DESC").fetchall()
    orders = conn.execute("""
        SELECT o.*, u.name AS customer_name FROM orders o
        LEFT JOIN users u ON u.id=o.user_id ORDER BY o.created_at DESC
    """).fetchall()
    conn.close()
    return render_template("report.html", plans=plans, orders=orders)


conn = get_db()
plan_count = conn.execute("SELECT COUNT(*) FROM plans").fetchone()[0]
conn.close()
if plan_count == 0:
    default_excel = os.path.join(BASE_DIR, "Маркетинг 11-р сарын төлөвлөгөө.xlsx")
    if os.path.exists(default_excel):
        import_excel(default_excel, os.path.basename(default_excel))

if __name__ == "__main__":
    app.run(
        debug=os.environ.get("FLASK_DEBUG", "1") == "1",
        host="0.0.0.0",
        port=int(os.environ.get("PORT", "5000")),
    )
