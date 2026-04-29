import sqlite3
import hashlib
import os
import secrets
from datetime import datetime

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
           t.due_date, t.created_at, t.updated_at,
           u1.name, u1.email, u2.name
    FROM tasks t
    JOIN users u1 ON t.assigned_to_id = u1.id
    JOIN users u2 ON t.assigned_by_id = u2.id
"""


def _task(r) -> dict:
    return {
        "id": r[0], "title": r[1], "description": r[2],
        "status": r[3], "progress": r[4], "due_date": r[5],
        "created_at": r[6], "updated_at": r[7],
        "assigned_to": r[8], "assigned_to_email": r[9], "assigned_by": r[10],
    }


def create_task(title: str, description: str, assigned_to_id: int,
                assigned_by_id: int, due_date: str) -> int:
    conn = get_conn()
    c = conn.cursor()
    now = _now()
    c.execute(
        "INSERT INTO tasks (title, description, assigned_to_id, assigned_by_id, "
        "status, progress, due_date, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?)",
        (title.strip(), description.strip(), assigned_to_id, assigned_by_id,
         "Not Started", 0, due_date, now, now),
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


def update_task_progress(task_id: int, progress: int, status: str):
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "UPDATE tasks SET progress=?, status=?, updated_at=? WHERE id=?",
        (progress, status, _now(), task_id),
    )
    conn.commit()
    conn.close()


def delete_task(task_id: int):
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM tasks WHERE id=?", (task_id,))
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
