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
    c.execute("SELECT COUNT(*) FROM users")
    if c.fetchone()[0] == 0:
        _seed_admin(c)
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

def create_license(tool_name: str, no_of_licenses: int, start_date: str, end_date: str) -> int:
    conn = get_conn()
    c = conn.cursor()
    now = _now()
    c.execute(
        "INSERT INTO licenses (tool_name, no_of_licenses, start_date, end_date, created_at, updated_at) "
        "VALUES (?,?,?,?,?,?)",
        (tool_name.strip(), no_of_licenses, start_date.strip(), end_date.strip(), now, now),
    )
    conn.commit()
    lid = c.lastrowid
    conn.close()
    return lid


def get_all_licenses() -> list:
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT id, tool_name, no_of_licenses, start_date, end_date, created_at FROM licenses ORDER BY id DESC")
    rows = c.fetchall()
    conn.close()
    return [
        {"id": r[0], "tool_name": r[1], "no_of_licenses": r[2],
         "start_date": r[3], "end_date": r[4], "created_at": r[5]}
        for r in rows
    ]


def update_license(license_id: int, tool_name: str, no_of_licenses: int,
                   start_date: str, end_date: str):
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "UPDATE licenses SET tool_name=?, no_of_licenses=?, start_date=?, end_date=?, updated_at=? WHERE id=?",
        (tool_name.strip(), no_of_licenses, start_date.strip(), end_date.strip(), _now(), license_id),
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
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
    """)
    conn.commit()
    conn.close()


def create_sold_license(tool_name: str, client: str, no_of_licenses: int,
                        start_date: str, end_date: str, notes: str = "") -> int:
    _ensure_sold_licenses_table()
    conn = get_conn()
    c = conn.cursor()
    now = _now()
    c.execute(
        "INSERT INTO sold_licenses (tool_name, client, no_of_licenses, start_date, end_date, notes, created_at, updated_at) "
        "VALUES (?,?,?,?,?,?,?,?)",
        (tool_name.strip(), client.strip(), no_of_licenses,
         start_date.strip(), end_date.strip(), notes.strip(), now, now),
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
        "SELECT id, tool_name, client, no_of_licenses, start_date, end_date, notes, created_at "
        "FROM sold_licenses ORDER BY id DESC"
    )
    rows = c.fetchall()
    conn.close()
    return [
        {"id": r[0], "tool_name": r[1], "client": r[2], "no_of_licenses": r[3],
         "start_date": r[4], "end_date": r[5], "notes": r[6], "created_at": r[7]}
        for r in rows
    ]


def update_sold_license(lid: int, tool_name: str, client: str, no_of_licenses: int,
                        start_date: str, end_date: str, notes: str = ""):
    _ensure_sold_licenses_table()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "UPDATE sold_licenses SET tool_name=?, client=?, no_of_licenses=?, "
        "start_date=?, end_date=?, notes=?, updated_at=? WHERE id=?",
        (tool_name.strip(), client.strip(), no_of_licenses,
         start_date.strip(), end_date.strip(), notes.strip(), _now(), lid),
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


# ── TASK COMMENT CRUD ──────────────────────────────────────────────────────────

def add_task_comment(task_id: int, user_id: int, comment: str, week_start: str) -> bool:
    """Insert weekly comment. Returns False if already exists for this week."""
    conn = get_conn()
    c = conn.cursor()
    try:
        c.execute(
            "INSERT INTO task_comments (task_id, user_id, comment, week_start, created_at) VALUES (?,?,?,?,?)",
            (task_id, user_id, comment.strip(), week_start, _now()),
        )
        conn.commit()
        return True
    except sqlite3.IntegrityError:
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
