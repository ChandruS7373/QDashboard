import sqlite3
import hashlib
import os
import secrets
from datetime import datetime, date, timedelta

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "qualesce.db")

ROLES = ["admin", "lead", "manager", "employee", "sales"]
TASK_STATUSES = ["Not Started", "In Progress", "Completed", "On Hold"]


def get_conn():
    return sqlite3.connect(DB_PATH)


def init_db():
    conn = get_conn()
    c = conn.cursor()
    c.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'employee',
            is_active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            description TEXT DEFAULT '',
            assigned_to_id INTEGER NOT NULL,
            assigned_by_id INTEGER NOT NULL,
            status TEXT NOT NULL DEFAULT 'Not Started',
            progress INTEGER NOT NULL DEFAULT 0,
            due_date TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY (assigned_to_id) REFERENCES users(id),
            FOREIGN KEY (assigned_by_id) REFERENCES users(id)
        );
        CREATE TABLE IF NOT EXISTS licenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tool_name TEXT NOT NULL,
            no_of_licenses INTEGER NOT NULL DEFAULT 1,
            start_date TEXT DEFAULT '',
            end_date TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );
    """)
    conn.commit()
    try:
        c.executescript("""
            CREATE TABLE IF NOT EXISTS task_comments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                task_id INTEGER NOT NULL,
                user_id INTEGER NOT NULL,
                comment TEXT NOT NULL,
                week_start TEXT NOT NULL,
                created_at TEXT NOT NULL,
                FOREIGN KEY (task_id) REFERENCES tasks(id),
                FOREIGN KEY (user_id) REFERENCES users(id),
                UNIQUE (task_id, user_id, week_start)
            );
        """)
        conn.commit()
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE tasks ADD COLUMN start_date TEXT DEFAULT ''")
        conn.commit()
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE tasks ADD COLUMN comment TEXT DEFAULT ''")
        conn.commit()
    except Exception:
        pass
    try:
        c.executescript("""
            CREATE TABLE IF NOT EXISTS email_settings (
                id INTEGER PRIMARY KEY DEFAULT 1,
                outlook_email TEXT NOT NULL DEFAULT '',
                outlook_password TEXT NOT NULL DEFAULT '',
                updated_at TEXT NOT NULL DEFAULT ''
            );
            INSERT OR IGNORE INTO email_settings (id, outlook_email, outlook_password, updated_at)
            VALUES (1, '', '', '');
        """)
        conn.commit()
    except Exception:
        pass
    try:
        c.executescript("""
            CREATE TABLE IF NOT EXISTS projects (
                id INTEGER PRIMARY KEY,
                name TEXT DEFAULT '',
                client TEXT DEFAULT '',
                lead TEXT DEFAULT '',
                employee TEXT DEFAULT '',
                status TEXT DEFAULT '',
                proj_type TEXT DEFAULT '',
                start TEXT DEFAULT '',
                end TEXT DEFAULT '',
                due_date TEXT DEFAULT '',
                po TEXT DEFAULT '',
                desc TEXT DEFAULT '',
                manual_hrs TEXT DEFAULT '',
                auto_hrs TEXT DEFAULT '',
                cost_per_hr TEXT DEFAULT '',
                hours_saved TEXT DEFAULT '',
                cost_saved TEXT DEFAULT '',
                roi_pct TEXT DEFAULT '',
                is_new INTEGER DEFAULT 0,
                is_active INTEGER DEFAULT 1,
                ckpt_pdd_sdd_start TEXT DEFAULT '',
                ckpt_pdd_sdd_end TEXT DEFAULT '',
                ckpt_development_start TEXT DEFAULT '',
                ckpt_development_end TEXT DEFAULT '',
                ckpt_uat_start TEXT DEFAULT '',
                ckpt_uat_end TEXT DEFAULT '',
                ckpt_deployment_start TEXT DEFAULT '',
                ckpt_deployment_end TEXT DEFAULT '',
                created_at TEXT DEFAULT '',
                updated_at TEXT DEFAULT ''
            );
        """)
        conn.commit()
    except Exception:
        pass
    c.execute("SELECT COUNT(*) FROM users")
    if c.fetchone()[0] == 0:
        _seed_admin(c)
        conn.commit()
    conn.close()


_PROJECT_COLS = [
    "id","name","client","lead","employee","status","proj_type","start","end","due_date","po","desc",
    "manual_hrs","auto_hrs","cost_per_hr","hours_saved","cost_saved","roi_pct","is_new","is_active",
    "ckpt_pdd_sdd_start","ckpt_pdd_sdd_end","ckpt_development_start","ckpt_development_end",
    "ckpt_uat_start","ckpt_uat_end","ckpt_deployment_start","ckpt_deployment_end",
]


def get_all_projects() -> list:
    conn = get_conn()
    c = conn.cursor()
    c.execute(f"SELECT {','.join(_PROJECT_COLS)} FROM projects ORDER BY id")
    rows = c.fetchall()
    conn.close()
    return [dict(zip(_PROJECT_COLS, r)) for r in rows]


def upsert_projects(records: list):
    """Persist the full project list to SQLite (upsert by id, delete removed rows)."""
    if not records:
        return
    conn = get_conn()
    c = conn.cursor()
    now = _now()
    incoming_ids = []
    rows_to_save = []
    for r in records:
        try:
            pid = int(float(str(r.get("id", "") or 0)))
        except Exception:
            pid = 0
        if not pid:
            continue
        incoming_ids.append(pid)
        vals = [pid] + [str(r.get(col, "") or "") for col in _PROJECT_COLS[1:]]
        vals[_PROJECT_COLS.index("is_new")] = 1 if r.get("is_new") else 0
        vals[_PROJECT_COLS.index("is_active")] = 1 if r.get("is_active", True) else 0
        rows_to_save.append(vals)
    if incoming_ids:
        placeholders = ",".join("?" * len(incoming_ids))
        c.execute(f"DELETE FROM projects WHERE id NOT IN ({placeholders})", incoming_ids)
    for vals in rows_to_save:
        pid = vals[0]
        c.execute("SELECT id FROM projects WHERE id=?", (pid,))
        if c.fetchone():
            set_clause = ", ".join(f"{col}=?" for col in _PROJECT_COLS[1:]) + ", updated_at=?"
            c.execute(f"UPDATE projects SET {set_clause} WHERE id=?", vals[1:] + [now, pid])
        else:
            full_cols = _PROJECT_COLS + ["created_at", "updated_at"]
            ph = ",".join(["?"] * len(full_cols))
            c.execute(f"INSERT INTO projects ({','.join(full_cols)}) VALUES ({ph})", vals + [now, now])
    conn.commit()
    conn.close()


def _seed_admin(cur):
    cur.execute(
        "INSERT INTO users (name, email, password_hash, role, is_active, created_at) VALUES (?,?,?,?,?,?)",
        ("Admin", "admin@qualesce.com", _hash("Admin@123"), "admin", 1, _now()),
    )


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _hash(password: str) -> str:
    salt = secrets.token_hex(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 200_000)
    return f"{salt}:{dk.hex()}"


def verify_password(password: str, stored: str) -> bool:
    try:
        salt, dk_hex = stored.split(":", 1)
        dk = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 200_000)
        return secrets.compare_digest(dk.hex(), dk_hex)
    except Exception:
        return False


def authenticate(email: str, password: str):
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT id, name, email, password_hash, role, is_active FROM users WHERE email=?",
        (email.strip().lower(),),
    )
    row = c.fetchone()
    conn.close()
    if row and row[5] == 1 and verify_password(password, row[3]):
        return {"id": row[0], "name": row[1], "email": row[2], "role": row[4]}
    return None


def get_user_by_email(email: str):
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT id, name, email, role, is_active FROM users WHERE email=?",
        (email.strip().lower(),),
    )
    row = c.fetchone()
    conn.close()
    if row and row[4] == 1:
        return {"id": row[0], "name": row[1], "email": row[2], "role": row[3]}
    return None


# ── USER CRUD ──────────────────────────────────────────────────────────────────

def create_user(name: str, email: str, password: str, role: str) -> int:
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "INSERT INTO users (name, email, password_hash, role, is_active, created_at) VALUES (?,?,?,?,?,?)",
        (name.strip(), email.strip().lower(), _hash(password), role, 1, _now()),
    )
    conn.commit()
    uid = c.lastrowid
    conn.close()
    return uid


def get_all_users() -> list:
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT id, name, email, role, is_active, created_at FROM users ORDER BY id")
    rows = c.fetchall()
    conn.close()
    return [
        {"id": r[0], "name": r[1], "email": r[2], "role": r[3],
         "is_active": bool(r[4]), "created_at": r[5]}
        for r in rows
    ]


def get_employees() -> list:
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT id, name, email FROM users WHERE role='employee' AND is_active=1 ORDER BY name"
    )
    rows = c.fetchall()
    conn.close()
    return [{"id": r[0], "name": r[1], "email": r[2]} for r in rows]


def get_leads() -> list:
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT id, name, email FROM users WHERE role='lead' AND is_active=1 ORDER BY name"
    )
    rows = c.fetchall()
    conn.close()
    return [{"id": r[0], "name": r[1], "email": r[2]} for r in rows]


def get_employees_and_leads() -> list:
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT id, name, email, role FROM users WHERE role IN ('employee','lead') AND is_active=1 ORDER BY role, name"
    )
    rows = c.fetchall()
    conn.close()
    return [{"id": r[0], "name": r[1], "email": r[2], "role": r[3]} for r in rows]


def update_user(user_id: int, name: str, email: str, role: str):
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "UPDATE users SET name=?, email=?, role=? WHERE id=?",
        (name.strip(), email.strip().lower(), role, user_id),
    )
    conn.commit()
    conn.close()


def reset_password(user_id: int, new_password: str):
    conn = get_conn()
    c = conn.cursor()
    c.execute("UPDATE users SET password_hash=? WHERE id=?", (_hash(new_password), user_id))
    conn.commit()
    conn.close()


# ── EMAIL SETTINGS ─────────────────────────────────────────────────────────────

def _ensure_email_settings(c):
    c.execute("""
        CREATE TABLE IF NOT EXISTS email_settings (
            id INTEGER PRIMARY KEY DEFAULT 1,
            outlook_email TEXT NOT NULL DEFAULT '',
            outlook_password TEXT NOT NULL DEFAULT '',
            updated_at TEXT NOT NULL DEFAULT ''
        )
    """)
    c.execute("INSERT OR IGNORE INTO email_settings (id, outlook_email, outlook_password, updated_at) VALUES (1,'','','')")


def get_email_settings() -> dict:
    conn = get_conn()
    c = conn.cursor()
    _ensure_email_settings(c)
    conn.commit()
    c.execute("SELECT outlook_email, outlook_password, updated_at FROM email_settings WHERE id=1")
    row = c.fetchone()
    conn.close()
    if row:
        return {"outlook_email": row[0], "outlook_password": row[1], "updated_at": row[2]}
    return {"outlook_email": "", "outlook_password": "", "updated_at": ""}


def save_email_settings(outlook_email: str, outlook_password: str):
    conn = get_conn()
    c = conn.cursor()
    _ensure_email_settings(c)
    c.execute(
        "UPDATE email_settings SET outlook_email=?, outlook_password=?, updated_at=? WHERE id=1",
        (outlook_email.strip().lower(), outlook_password, _now()),
    )
    conn.commit()
    conn.close()


# ── PER-USER EMAIL SETTINGS ────────────────────────────────────────────────────

def _ensure_user_email_settings(c):
    c.execute("""
        CREATE TABLE IF NOT EXISTS user_email_settings (
            user_id INTEGER PRIMARY KEY,
            outlook_email TEXT NOT NULL DEFAULT '',
            outlook_password TEXT NOT NULL DEFAULT '',
            updated_at TEXT NOT NULL DEFAULT ''
        )
    """)


def get_user_email_settings(user_id: int) -> dict:
    conn = get_conn()
    c = conn.cursor()
    _ensure_user_email_settings(c)
    conn.commit()
    c.execute("SELECT outlook_email, outlook_password, updated_at FROM user_email_settings WHERE user_id=?", (user_id,))
    row = c.fetchone()
    conn.close()
    if row:
        return {"outlook_email": row[0], "outlook_password": row[1], "updated_at": row[2]}
    return {"outlook_email": "", "outlook_password": "", "updated_at": ""}


def save_user_email_settings(user_id: int, outlook_email: str, outlook_password: str):
    conn = get_conn()
    c = conn.cursor()
    _ensure_user_email_settings(c)
    c.execute(
        "INSERT OR REPLACE INTO user_email_settings (user_id, outlook_email, outlook_password, updated_at) VALUES (?,?,?,?)",
        (user_id, outlook_email.strip().lower(), outlook_password, _now()),
    )
    conn.commit()
    conn.close()


def set_active(user_id: int, active: bool):
    conn = get_conn()
    c = conn.cursor()
    c.execute("UPDATE users SET is_active=? WHERE id=?", (1 if active else 0, user_id))
    conn.commit()
    conn.close()


def delete_user(user_id: int):
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM users WHERE id=?", (user_id,))
    conn.commit()
    conn.close()


# ── TASK CRUD ──────────────────────────────────────────────────────────────────

_TASK_SQL = """
    SELECT t.id, t.title, t.description, t.status, t.progress,
           t.due_date, t.start_date, t.created_at, t.updated_at, t.comment,
           u1.name, u1.email, u2.name, u2.email
    FROM tasks t
    JOIN users u1 ON t.assigned_to_id = u1.id
    JOIN users u2 ON t.assigned_by_id = u2.id
"""


def _task(r) -> dict:
    return {
        "id": r[0], "title": r[1], "description": r[2],
        "status": r[3], "progress": r[4], "due_date": r[5],
        "start_date": r[6], "created_at": r[7], "updated_at": r[8],
        "comment": r[9], "assigned_to": r[10], "assigned_to_email": r[11],
        "assigned_by": r[12], "assigned_by_email": r[13],
    }


def create_task(title: str, description: str, assigned_to_id: int,
                assigned_by_id: int, due_date: str, start_date: str = "") -> int:
    conn = get_conn()
    c = conn.cursor()
    now = _now()
    c.execute(
        "INSERT INTO tasks (title, description, assigned_to_id, assigned_by_id, "
        "status, progress, start_date, due_date, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?)",
        (title.strip(), description.strip(), assigned_to_id, assigned_by_id,
         "Not Started", 0, start_date, due_date, now, now),
    )
    conn.commit()
    tid = c.lastrowid
    conn.close()
    return tid


def get_all_tasks() -> list:
    conn = get_conn()
    c = conn.cursor()
    c.execute(_TASK_SQL + " ORDER BY t.id DESC")
    rows = c.fetchall()
    conn.close()
    return [_task(r) for r in rows]


def get_all_tasks_asc() -> list:
    conn = get_conn()
    c = conn.cursor()
    c.execute(_TASK_SQL + " ORDER BY t.id ASC")
    rows = c.fetchall()
    conn.close()
    return [_task(r) for r in rows]


def sync_tasks_from_excel(excel_path: str):
    """Import tasks from the Excel Tasks sheet into SQLite on startup.
    Skips rows where assigned_to or assigned_by email doesn't match a DB user.
    Existing task IDs are updated; new rows are inserted."""
    if not excel_path:
        return
    try:
        import pandas as pd
        df = pd.read_excel(excel_path, sheet_name="Tasks", dtype=str, engine="openpyxl").fillna("")
    except Exception:
        return
    if df.empty:
        return
    conn = get_conn()
    c = conn.cursor()
    # Build email→id lookup for users
    c.execute("SELECT id, email FROM users")
    email_map = {row[1].strip().lower(): row[0] for row in c.fetchall()}
    for _, row in df.iterrows():
        try:
            task_id  = int(float(row.get("id", 0) or 0))
            title    = str(row.get("title", "")).strip()
            if not task_id or not title:
                continue
            desc     = str(row.get("description", "")).strip()
            status   = str(row.get("status", "Not Started")).strip()
            progress = int(float(row.get("progress", 0) or 0))
            due_date = str(row.get("due_date", "")).strip()
            start_date = str(row.get("start_date", "")).strip()
            comment  = str(row.get("comment", "")).strip()
            created_at = str(row.get("created_at", _now())).strip() or _now()
            updated_at = str(row.get("updated_at", _now())).strip() or _now()
            to_email  = str(row.get("assigned_to_email", "")).strip().lower()
            by_email  = str(row.get("assigned_by_email", "")).strip().lower()
            to_id  = email_map.get(to_email)
            by_id  = email_map.get(by_email)
            if not to_id or not by_id:
                continue
            c.execute("SELECT id FROM tasks WHERE id=?", (task_id,))
            if c.fetchone():
                c.execute(
                    "UPDATE tasks SET title=?,description=?,assigned_to_id=?,assigned_by_id=?,"
                    "status=?,progress=?,start_date=?,due_date=?,comment=?,updated_at=? WHERE id=?",
                    (title, desc, to_id, by_id, status, progress, start_date, due_date, comment, updated_at, task_id))
            else:
                c.execute(
                    "INSERT INTO tasks (id,title,description,assigned_to_id,assigned_by_id,"
                    "status,progress,start_date,due_date,comment,created_at,updated_at) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                    (task_id, title, desc, to_id, by_id, status, progress, start_date, due_date, comment, created_at, updated_at))
        except Exception:
            continue
    conn.commit()
    conn.close()


def sync_users_from_excel(users_excel_path: str):
    """Import users from users.xlsx into SQLite so they can log in after a restart."""
    if not users_excel_path or not os.path.exists(users_excel_path):
        return
    try:
        import pandas as pd
        df = pd.read_excel(users_excel_path, dtype=str, engine="openpyxl").fillna("")
    except Exception:
        return
    if df.empty:
        return
    conn = get_conn()
    c = conn.cursor()
    for _, row in df.iterrows():
        try:
            name = str(row.get("Name", "")).strip()
            email = str(row.get("Email", "")).strip().lower()
            role = str(row.get("Role", "employee")).strip()
            password = str(row.get("Password", "")).strip()
            active_str = str(row.get("Active", "Yes")).strip().lower()
            is_active = 1 if active_str in ("yes", "1", "true") else 0
            if not name or not email or not password:
                continue
            if role not in ROLES:
                role = "employee"
            c.execute("SELECT id FROM users WHERE email=?", (email,))
            if not c.fetchone():
                c.execute(
                    "INSERT INTO users (name, email, password_hash, role, is_active, created_at) VALUES (?,?,?,?,?,?)",
                    (name, email, _hash(password), role, is_active, _now()),
                )
        except Exception:
            continue
    conn.commit()
    conn.close()


def sync_comments_from_excel(excel_path: str):
    """Bidirectional sync: import missing comments from Excel and remove ones deleted from Excel."""
    if not excel_path or not os.path.exists(excel_path):
        return
    try:
        import pandas as pd
        df = pd.read_excel(excel_path, sheet_name="Comments", dtype=str, engine="openpyxl").fillna("")
    except Exception:
        return

    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT id, email FROM users")
    email_map = {row[1].strip().lower(): row[0] for row in c.fetchall()}

    # Build set of (task_id, user_id, week_start) from Excel
    excel_keys = set()
    for _, row in df.iterrows():
        try:
            task_id = int(float(str(row.get("task_id", 0) or 0)))
            user_email = str(row.get("employee_email", "")).strip().lower()
            comment_text = str(row.get("comment", "")).strip()
            week_start = str(row.get("week_start", "")).strip()
            created_at = str(row.get("created_at", "")).strip() or _now()
            if not task_id or not user_email or not comment_text or not week_start:
                continue
            user_id = email_map.get(user_email)
            if not user_id:
                continue
            c.execute("SELECT id FROM tasks WHERE id=?", (task_id,))
            if not c.fetchone():
                continue
            excel_keys.add((task_id, user_id, week_start))
            c.execute(
                "INSERT OR IGNORE INTO task_comments (task_id, user_id, comment, week_start, created_at) VALUES (?,?,?,?,?)",
                (task_id, user_id, comment_text, week_start, created_at),
            )
        except Exception:
            continue

    # Remove SQLite comments not present in Excel (deleted from Excel)
    c.execute("SELECT id, task_id, user_id, week_start FROM task_comments")
    for row in c.fetchall():
        key = (row[1], row[2], row[3])
        if key not in excel_keys:
            c.execute("DELETE FROM task_comments WHERE id=?", (row[0],))

    conn.commit()
    conn.close()


def sync_tasks_from_df(df):
    """Import tasks from a DataFrame into SQLite (same logic as sync_tasks_from_excel)."""
    import pandas as pd
    if df is None or (hasattr(df, 'empty') and df.empty):
        return
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT id, email FROM users")
    email_map = {row[1].strip().lower(): row[0] for row in c.fetchall()}
    for _, row in df.iterrows():
        try:
            task_id = int(float(str(row.get("id", 0) or 0)))
            title = str(row.get("title", "")).strip()
            if not task_id or not title:
                continue
            desc = str(row.get("description", "")).strip()
            status = str(row.get("status", "Not Started")).strip()
            progress = int(float(str(row.get("progress", 0) or 0)))
            due_date = str(row.get("due_date", "")).strip()
            start_date = str(row.get("start_date", "")).strip()
            comment = str(row.get("comment", "")).strip()
            created_at = str(row.get("created_at", _now())).strip() or _now()
            updated_at = str(row.get("updated_at", _now())).strip() or _now()
            to_email = str(row.get("assigned_to_email", "")).strip().lower()
            by_email = str(row.get("assigned_by_email", "")).strip().lower()
            to_id = email_map.get(to_email)
            by_id = email_map.get(by_email)
            if not to_id or not by_id:
                continue
            c.execute("SELECT id FROM tasks WHERE id=?", (task_id,))
            if c.fetchone():
                c.execute(
                    "UPDATE tasks SET title=?,description=?,assigned_to_id=?,assigned_by_id=?,"
                    "status=?,progress=?,start_date=?,due_date=?,comment=?,updated_at=? WHERE id=?",
                    (title, desc, to_id, by_id, status, progress, start_date, due_date, comment, updated_at, task_id))
            else:
                c.execute(
                    "INSERT INTO tasks (id,title,description,assigned_to_id,assigned_by_id,"
                    "status,progress,start_date,due_date,comment,created_at,updated_at) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                    (task_id, title, desc, to_id, by_id, status, progress, start_date, due_date, comment, created_at, updated_at))
        except Exception:
            continue
    conn.commit()
    conn.close()


def sync_users_from_df(df):
    """Import users from a DataFrame into SQLite (same logic as sync_users_from_excel)."""
    if df is None or (hasattr(df, 'empty') and df.empty):
        return
    conn = get_conn()
    c = conn.cursor()
    for _, row in df.iterrows():
        try:
            name = str(row.get("Name", row.get("name", ""))).strip()
            email = str(row.get("Email", row.get("email", ""))).strip().lower()
            role = str(row.get("Role", row.get("role", "employee"))).strip()
            password = str(row.get("Password", row.get("password", ""))).strip()
            active_str = str(row.get("Active", row.get("is_active", "Yes"))).strip().lower()
            is_active = 1 if active_str in ("yes", "1", "true") else 0
            if not name or not email or not password:
                continue
            if role not in ROLES:
                role = "employee"
            c.execute("SELECT id FROM users WHERE email=?", (email,))
            if not c.fetchone():
                c.execute(
                    "INSERT INTO users (name, email, password_hash, role, is_active, created_at) VALUES (?,?,?,?,?,?)",
                    (name, email, _hash(password), role, is_active, _now()),
                )
        except Exception:
            continue
    conn.commit()
    conn.close()


def sync_comments_from_df(df):
    """Bidirectional sync from a DataFrame: import missing, delete removed."""
    if df is None or (hasattr(df, 'empty') and df.empty):
        return
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT id, email FROM users")
    email_map = {row[1].strip().lower(): row[0] for row in c.fetchall()}
    excel_keys = set()
    for _, row in df.iterrows():
        try:
            task_id = int(float(str(row.get("task_id", 0) or 0)))
            user_email = str(row.get("employee_email", "")).strip().lower()
            comment_text = str(row.get("comment", "")).strip()
            week_start = str(row.get("week_start", "")).strip()
            created_at = str(row.get("created_at", "")).strip() or _now()
            if not task_id or not user_email or not comment_text or not week_start:
                continue
            user_id = email_map.get(user_email)
            if not user_id:
                continue
            c.execute("SELECT id FROM tasks WHERE id=?", (task_id,))
            if not c.fetchone():
                continue
            excel_keys.add((task_id, user_id, week_start))
            c.execute(
                "INSERT OR IGNORE INTO task_comments (task_id, user_id, comment, week_start, created_at) VALUES (?,?,?,?,?)",
                (task_id, user_id, comment_text, week_start, created_at),
            )
        except Exception:
            continue
    c.execute("SELECT id, task_id, user_id, week_start FROM task_comments")
    for row in c.fetchall():
        if (row[1], row[2], row[3]) not in excel_keys:
            c.execute("DELETE FROM task_comments WHERE id=?", (row[0],))
    conn.commit()
    conn.close()


def get_user_tasks(user_id: int) -> list:
    conn = get_conn()
    c = conn.cursor()
    c.execute(_TASK_SQL + " WHERE t.assigned_to_id=? ORDER BY t.id DESC", (user_id,))
    rows = c.fetchall()
    conn.close()
    return [_task(r) for r in rows]


def update_task_progress(task_id: int, progress: int, status: str, comment: str = ""):
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "UPDATE tasks SET progress=?, status=?, comment=?, updated_at=? WHERE id=?",
        (progress, status, comment, _now(), task_id),
    )
    conn.commit()
    conn.close()


def delete_task(task_id: int):
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM task_comments WHERE task_id=?", (task_id,))
    c.execute("DELETE FROM tasks WHERE id=?", (task_id,))
    conn.commit()
    conn.close()


def update_task_meta(task_id: int, title: str, description: str, start_date: str, due_date: str):
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "UPDATE tasks SET title=?, description=?, start_date=?, due_date=?, updated_at=? WHERE id=?",
        (title.strip(), description.strip(), start_date, due_date, _now(), task_id),
    )
    conn.commit()
    conn.close()


# ── LICENSE CRUD ───────────────────────────────────────────────────────────────

def _ensure_license_extras():
    """Migrate: add client_email to licenses; create notification_log table."""
    conn = get_conn()
    c = conn.cursor()
    try:
        c.execute("ALTER TABLE licenses ADD COLUMN client_email TEXT DEFAULT ''")
        conn.commit()
    except Exception:
        pass
    try:
        c.execute("""
            CREATE TABLE IF NOT EXISTS license_notification_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                license_id INTEGER NOT NULL,
                license_type TEXT NOT NULL,
                threshold TEXT NOT NULL,
                sent_at TEXT NOT NULL,
                UNIQUE(license_id, license_type, threshold)
            )
        """)
        conn.commit()
    except Exception:
        pass
    conn.close()


def has_notification_been_sent(license_id: int, license_type: str, threshold: str) -> bool:
    _ensure_license_extras()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT 1 FROM license_notification_log WHERE license_id=? AND license_type=? AND threshold=?",
        (license_id, license_type, threshold)
    )
    found = c.fetchone() is not None
    conn.close()
    return found


def mark_notification_sent(license_id: int, license_type: str, threshold: str):
    _ensure_license_extras()
    conn = get_conn()
    c = conn.cursor()
    try:
        c.execute(
            "INSERT OR IGNORE INTO license_notification_log (license_id, license_type, threshold, sent_at) VALUES (?,?,?,?)",
            (license_id, license_type, threshold, _now())
        )
        conn.commit()
    except Exception:
        pass
    conn.close()


def create_license(tool_name: str, no_of_licenses: int, start_date: str, end_date: str,
                   client_email: str = "") -> int:
    _ensure_license_extras()
    conn = get_conn()
    c = conn.cursor()
    now = _now()
    c.execute(
        "INSERT INTO licenses (tool_name, no_of_licenses, start_date, end_date, client_email, created_at, updated_at) "
        "VALUES (?,?,?,?,?,?,?)",
        (tool_name.strip(), no_of_licenses, start_date.strip(), end_date.strip(),
         client_email.strip(), now, now),
    )
    conn.commit()
    lid = c.lastrowid
    conn.close()
    return lid


def get_all_licenses() -> list:
    _ensure_license_extras()
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT id, tool_name, no_of_licenses, start_date, end_date, client_email, created_at FROM licenses ORDER BY id DESC")
    rows = c.fetchall()
    conn.close()
    return [
        {"id": r[0], "tool_name": r[1], "no_of_licenses": r[2],
         "start_date": r[3], "end_date": r[4],
         "client_email": r[5] or "", "created_at": r[6]}
        for r in rows
    ]


def update_license(license_id: int, tool_name: str, no_of_licenses: int,
                   start_date: str, end_date: str, client_email: str = ""):
    _ensure_license_extras()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "UPDATE licenses SET tool_name=?, no_of_licenses=?, start_date=?, end_date=?, client_email=?, updated_at=? WHERE id=?",
        (tool_name.strip(), no_of_licenses, start_date.strip(), end_date.strip(),
         client_email.strip(), _now(), license_id),
    )
    conn.commit()
    conn.close()


def delete_license(license_id: int):
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM licenses WHERE id=?", (license_id,))
    conn.commit()
    conn.close()


# ── SOLD LICENSE CRUD ──────────────────────────────────────────────────────────

def _ensure_sold_licenses_table():
    conn = get_conn()
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS sold_licenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tool_name TEXT NOT NULL,
            client TEXT NOT NULL,
            no_of_licenses INTEGER NOT NULL DEFAULT 1,
            start_date TEXT DEFAULT '',
            end_date TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            client_email TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
    """)
    # Migration: add client_email if missing from older DB
    try:
        c.execute("ALTER TABLE sold_licenses ADD COLUMN client_email TEXT DEFAULT ''")
        conn.commit()
    except Exception:
        pass
    conn.commit()
    conn.close()


def create_sold_license(tool_name: str, client: str, no_of_licenses: int,
                        start_date: str, end_date: str, notes: str = "",
                        client_email: str = "") -> int:
    _ensure_sold_licenses_table()
    conn = get_conn()
    c = conn.cursor()
    now = _now()
    c.execute(
        "INSERT INTO sold_licenses (tool_name, client, no_of_licenses, start_date, end_date, notes, client_email, created_at, updated_at) "
        "VALUES (?,?,?,?,?,?,?,?,?)",
        (tool_name.strip(), client.strip(), no_of_licenses,
         start_date.strip(), end_date.strip(), notes.strip(),
         client_email.strip(), now, now),
    )
    conn.commit()
    lid = c.lastrowid
    conn.close()
    return lid


def get_all_sold_licenses() -> list:
    _ensure_sold_licenses_table()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT id, tool_name, client, no_of_licenses, start_date, end_date, notes, client_email, created_at "
        "FROM sold_licenses ORDER BY id DESC"
    )
    rows = c.fetchall()
    conn.close()
    return [
        {"id": r[0], "tool_name": r[1], "client": r[2], "no_of_licenses": r[3],
         "start_date": r[4], "end_date": r[5], "notes": r[6],
         "client_email": r[7] or "", "created_at": r[8]}
        for r in rows
    ]


def update_sold_license(lid: int, tool_name: str, client: str, no_of_licenses: int,
                        start_date: str, end_date: str, notes: str = "",
                        client_email: str = ""):
    _ensure_sold_licenses_table()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "UPDATE sold_licenses SET tool_name=?, client=?, no_of_licenses=?, "
        "start_date=?, end_date=?, notes=?, client_email=?, updated_at=? WHERE id=?",
        (tool_name.strip(), client.strip(), no_of_licenses,
         start_date.strip(), end_date.strip(), notes.strip(),
         client_email.strip(), _now(), lid),
    )
    conn.commit()
    conn.close()


def delete_sold_license(lid: int):
    _ensure_sold_licenses_table()
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM sold_licenses WHERE id=?", (lid,))
    conn.commit()
    conn.close()


def get_week_start(dt=None) -> str:
    """Returns the Monday of the week for dt (or today) as YYYY-MM-DD."""
    d = dt or date.today()
    return (d - timedelta(days=d.weekday())).isoformat()


# ── CRM TABLES ─────────────────────────────────────────────────────────────────

CRM_LEAD_STATUSES  = ["New", "Contacted", "Qualified", "Proposal", "Negotiation", "Won", "Lost"]
CRM_LEAD_SOURCES   = ["Direct", "Referral", "Website", "LinkedIn", "Email Campaign", "Cold Call", "Other"]
CRM_OPP_STAGES     = ["Prospecting", "Qualification", "Proposal", "Negotiation", "Closed Won", "Closed Lost"]
CRM_ACTIVITY_TYPES = ["Call", "Email", "Meeting", "Demo", "Follow-up", "Other"]


def _ensure_crm_tables():
    conn = get_conn()
    c = conn.cursor()
    c.executescript("""
        CREATE TABLE IF NOT EXISTS crm_leads (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_name TEXT NOT NULL,
            contact_name TEXT DEFAULT '',
            email TEXT DEFAULT '',
            phone TEXT DEFAULT '',
            source TEXT DEFAULT '',
            status TEXT NOT NULL DEFAULT 'New',
            notes TEXT DEFAULT '',
            assigned_to_id INTEGER,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY (assigned_to_id) REFERENCES users(id)
        );
        CREATE TABLE IF NOT EXISTS crm_opportunities (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            lead_id INTEGER,
            title TEXT NOT NULL,
            value REAL DEFAULT 0,
            stage TEXT NOT NULL DEFAULT 'Prospecting',
            probability INTEGER DEFAULT 0,
            expected_close TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            assigned_to_id INTEGER,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY (lead_id) REFERENCES crm_leads(id),
            FOREIGN KEY (assigned_to_id) REFERENCES users(id)
        );
        CREATE TABLE IF NOT EXISTS crm_activities (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            lead_id INTEGER,
            opportunity_id INTEGER,
            type TEXT DEFAULT 'Call',
            subject TEXT NOT NULL,
            notes TEXT DEFAULT '',
            activity_date TEXT DEFAULT '',
            is_done INTEGER DEFAULT 0,
            created_by_id INTEGER,
            created_at TEXT NOT NULL,
            FOREIGN KEY (lead_id) REFERENCES crm_leads(id),
            FOREIGN KEY (opportunity_id) REFERENCES crm_opportunities(id),
            FOREIGN KEY (created_by_id) REFERENCES users(id)
        );
    """)
    conn.commit()
    conn.close()


# ── CRM LEAD CRUD ──────────────────────────────────────────────────────────────

def create_lead(company_name: str, contact_name: str, email: str, phone: str,
                source: str, status: str, notes: str, assigned_to_id=None) -> int:
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    now = _now()
    c.execute(
        "INSERT INTO crm_leads (company_name, contact_name, email, phone, source, status, notes, assigned_to_id, created_at, updated_at) "
        "VALUES (?,?,?,?,?,?,?,?,?,?)",
        (company_name.strip(), contact_name.strip(), email.strip().lower(), phone.strip(),
         source, status, notes.strip(), assigned_to_id, now, now),
    )
    conn.commit()
    lid = c.lastrowid
    conn.close()
    return lid


def get_all_leads() -> list:
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute("""
        SELECT l.id, l.company_name, l.contact_name, l.email, l.phone,
               l.source, l.status, l.notes, l.assigned_to_id,
               u.name, l.created_at, l.updated_at
        FROM crm_leads l
        LEFT JOIN users u ON l.assigned_to_id = u.id
        ORDER BY l.id DESC
    """)
    rows = c.fetchall()
    conn.close()
    return [
        {"id": r[0], "company_name": r[1], "contact_name": r[2], "email": r[3],
         "phone": r[4], "source": r[5], "status": r[6], "notes": r[7],
         "assigned_to_id": r[8], "assigned_to": r[9] or "",
         "created_at": r[10], "updated_at": r[11]}
        for r in rows
    ]


def update_lead(lead_id: int, company_name: str, contact_name: str, email: str,
                phone: str, source: str, status: str, notes: str, assigned_to_id=None):
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "UPDATE crm_leads SET company_name=?, contact_name=?, email=?, phone=?, "
        "source=?, status=?, notes=?, assigned_to_id=?, updated_at=? WHERE id=?",
        (company_name.strip(), contact_name.strip(), email.strip().lower(), phone.strip(),
         source, status, notes.strip(), assigned_to_id, _now(), lead_id),
    )
    conn.commit()
    conn.close()


def delete_lead(lead_id: int):
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM crm_activities WHERE lead_id=?", (lead_id,))
    c.execute("DELETE FROM crm_opportunities WHERE lead_id=?", (lead_id,))
    c.execute("DELETE FROM crm_leads WHERE id=?", (lead_id,))
    conn.commit()
    conn.close()


# ── CRM OPPORTUNITY CRUD ───────────────────────────────────────────────────────

def create_opportunity(lead_id, title: str, value: float, stage: str,
                       probability: int, expected_close: str, notes: str,
                       assigned_to_id=None) -> int:
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    now = _now()
    c.execute(
        "INSERT INTO crm_opportunities (lead_id, title, value, stage, probability, expected_close, notes, assigned_to_id, created_at, updated_at) "
        "VALUES (?,?,?,?,?,?,?,?,?,?)",
        (lead_id, title.strip(), value, stage, probability, expected_close, notes.strip(), assigned_to_id, now, now),
    )
    conn.commit()
    oid = c.lastrowid
    conn.close()
    return oid


def get_all_opportunities() -> list:
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute("""
        SELECT o.id, o.lead_id, l.company_name, o.title, o.value, o.stage,
               o.probability, o.expected_close, o.notes, o.assigned_to_id,
               u.name, o.created_at, o.updated_at
        FROM crm_opportunities o
        LEFT JOIN crm_leads l ON o.lead_id = l.id
        LEFT JOIN users u ON o.assigned_to_id = u.id
        ORDER BY o.id DESC
    """)
    rows = c.fetchall()
    conn.close()
    return [
        {"id": r[0], "lead_id": r[1], "company_name": r[2] or "", "title": r[3],
         "value": r[4] or 0, "stage": r[5], "probability": r[6] or 0,
         "expected_close": r[7] or "", "notes": r[8] or "",
         "assigned_to_id": r[9], "assigned_to": r[10] or "",
         "created_at": r[11], "updated_at": r[12]}
        for r in rows
    ]


def update_opportunity(opp_id: int, lead_id, title: str, value: float, stage: str,
                       probability: int, expected_close: str, notes: str, assigned_to_id=None):
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "UPDATE crm_opportunities SET lead_id=?, title=?, value=?, stage=?, probability=?, "
        "expected_close=?, notes=?, assigned_to_id=?, updated_at=? WHERE id=?",
        (lead_id, title.strip(), value, stage, probability, expected_close,
         notes.strip(), assigned_to_id, _now(), opp_id),
    )
    conn.commit()
    conn.close()


def delete_opportunity(opp_id: int):
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM crm_activities WHERE opportunity_id=?", (opp_id,))
    c.execute("DELETE FROM crm_opportunities WHERE id=?", (opp_id,))
    conn.commit()
    conn.close()


# ── CRM ACTIVITY CRUD ──────────────────────────────────────────────────────────

def create_activity(lead_id, opportunity_id, act_type: str, subject: str,
                    notes: str, activity_date: str, created_by_id=None) -> int:
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "INSERT INTO crm_activities (lead_id, opportunity_id, type, subject, notes, activity_date, is_done, created_by_id, created_at) "
        "VALUES (?,?,?,?,?,?,?,?,?)",
        (lead_id, opportunity_id, act_type, subject.strip(), notes.strip(),
         activity_date, 0, created_by_id, _now()),
    )
    conn.commit()
    aid = c.lastrowid
    conn.close()
    return aid


def get_all_activities() -> list:
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute("""
        SELECT a.id, a.lead_id, l.company_name, a.opportunity_id, o.title,
               a.type, a.subject, a.notes, a.activity_date, a.is_done,
               a.created_by_id, u.name, a.created_at
        FROM crm_activities a
        LEFT JOIN crm_leads l ON a.lead_id = l.id
        LEFT JOIN crm_opportunities o ON a.opportunity_id = o.id
        LEFT JOIN users u ON a.created_by_id = u.id
        ORDER BY a.activity_date DESC, a.id DESC
    """)
    rows = c.fetchall()
    conn.close()
    return [
        {"id": r[0], "lead_id": r[1], "company_name": r[2] or "", "opportunity_id": r[3],
         "opportunity_title": r[4] or "", "type": r[5], "subject": r[6],
         "notes": r[7] or "", "activity_date": r[8] or "", "is_done": bool(r[9]),
         "created_by_id": r[10], "created_by": r[11] or "", "created_at": r[12]}
        for r in rows
    ]


def update_activity_done(activity_id: int, is_done: bool):
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute("UPDATE crm_activities SET is_done=? WHERE id=?", (1 if is_done else 0, activity_id))
    conn.commit()
    conn.close()


def delete_activity(activity_id: int):
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM crm_activities WHERE id=?", (activity_id,))
    conn.commit()
    conn.close()


# ── TASK COMMENT CRUD ──────────────────────────────────────────────────────────

def add_task_comment(task_id: int, user_id: int, comment: str, week_start: str) -> bool:
    """Insert weekly comment. Returns False if already exists or any DB error occurs."""
    conn = get_conn()
    c = conn.cursor()
    try:
        c.execute(
            "INSERT INTO task_comments (task_id, user_id, comment, week_start, created_at) VALUES (?,?,?,?,?)",
            (task_id, user_id, comment.strip(), week_start, _now()),
        )
        conn.commit()
        return True
    except Exception:
        return False
    finally:
        conn.close()


def get_user_week_comment(task_id: int, user_id: int, week_start: str):
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT id, comment, created_at FROM task_comments WHERE task_id=? AND user_id=? AND week_start=?",
        (task_id, user_id, week_start),
    )
    row = c.fetchone()
    conn.close()
    return {"id": row[0], "comment": row[1], "created_at": row[2]} if row else None


def get_task_comments_with_users(task_id: int = None, from_date: str = None, to_date: str = None) -> list:
    conn = get_conn()
    c = conn.cursor()
    sql = """
        SELECT tc.id, tc.task_id, t.title, tc.week_start, tc.comment,
               tc.created_at, u.name, u.email
        FROM task_comments tc
        JOIN tasks t ON tc.task_id = t.id
        JOIN users u ON tc.user_id = u.id
        WHERE 1=1
    """
    params = []
    if task_id is not None:
        sql += " AND tc.task_id=?"
        params.append(task_id)
    if from_date:
        sql += " AND tc.week_start>=?"
        params.append(from_date)
    if to_date:
        sql += " AND tc.week_start<=?"
        params.append(to_date)
    sql += " ORDER BY tc.week_start DESC, tc.task_id, u.name"
    c.execute(sql, params)
    rows = c.fetchall()
    conn.close()
    return [
        {"id": r[0], "task_id": r[1], "task_title": r[2], "week_start": r[3],
         "comment": r[4], "created_at": r[5], "user_name": r[6], "user_email": r[7]}
        for r in rows
    ]


def get_all_comments_for_excel() -> list:
    """Return all task comments with task title, employee info, and timestamps for Excel export."""
    conn = get_conn()
    c = conn.cursor()
    c.execute("""
        SELECT tc.id, t.id, t.title, u.name, u.email,
               tc.comment, tc.week_start, tc.created_at
        FROM task_comments tc
        JOIN tasks t ON tc.task_id = t.id
        JOIN users u ON tc.user_id = u.id
        ORDER BY tc.created_at DESC
    """)
    rows = c.fetchall()
    conn.close()
    return [
        {"id": r[0], "task_id": r[1], "task_title": r[2],
         "employee_name": r[3], "employee_email": r[4],
         "comment": r[5], "week_start": r[6], "created_at": r[7]}
        for r in rows
    ]
