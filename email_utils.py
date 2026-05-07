import smtplib
import random
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

PORTAL_URL = "https://q-dashboard.streamlit.app/"


def _smtp_creds():
    # 1. Global DB settings (configured via admin Users tab)
    try:
        import auth as _auth
        cfg = _auth.get_email_settings()
        if cfg["outlook_email"] and cfg["outlook_password"]:
            return cfg["outlook_email"], cfg["outlook_password"]
    except Exception:
        pass
    # 2. Global secrets.toml / env var
    try:
        import streamlit as st
        email = st.secrets.get("OUTLOOK_EMAIL", "")
        pwd   = st.secrets.get("OUTLOOK_PASSWORD", "")
        if email and pwd:
            return email, pwd
    except Exception:
        pass
    return os.environ.get("OUTLOOK_EMAIL", ""), os.environ.get("OUTLOOK_PASSWORD", "")


def _smtp_creds_for_role(role: str, user_id: int = 0) -> tuple[str, str]:
    """Return (email, password) for a specific role/user, with fallback chain:
    1. Per-user DB  →  2. Role secrets  →  3. Global DB  →  4. Global secrets"""
    # 1. Per-user DB settings
    try:
        import auth as _auth
        cfg = _auth.get_user_email_settings(user_id)
        if cfg["outlook_email"] and cfg["outlook_password"]:
            return cfg["outlook_email"], cfg["outlook_password"]
    except Exception:
        pass
    # 2. Role-specific secrets (LEAD_OUTLOOK_EMAIL / MANAGER_OUTLOOK_EMAIL)
    try:
        import streamlit as st
        key = role.upper()
        r_email = st.secrets.get(f"{key}_OUTLOOK_EMAIL", "")
        r_pwd   = st.secrets.get(f"{key}_OUTLOOK_PASSWORD", "")
        if r_email and r_pwd:
            return r_email, r_pwd
    except Exception:
        pass
    # 3. Role-specific env vars
    key = role.upper()
    r_email = os.environ.get(f"{key}_OUTLOOK_EMAIL", "")
    r_pwd   = os.environ.get(f"{key}_OUTLOOK_PASSWORD", "")
    if r_email and r_pwd:
        return r_email, r_pwd
    # 4. Fall back to global settings
    return _smtp_creds()


def _smtp_send(sender: str, password: str, to_email: str, subject: str, html_body: str,
               reply_to: str = "", display_name: str = "") -> tuple[bool, str]:
    msg = MIMEMultipart("alternative")
    msg["From"] = f"{display_name} <{sender}>" if display_name else sender
    msg["To"] = to_email
    msg["Subject"] = subject
    if reply_to:
        msg["Reply-To"] = reply_to
    msg.attach(MIMEText(html_body, "html"))
    try:
        with smtplib.SMTP("smtp.office365.com", 587) as server:
            server.ehlo()
            server.starttls()
            server.login(sender, password)
            server.sendmail(sender, to_email, msg.as_string())
        return True, ""
    except Exception as e:
        return False, str(e)


def send_email(to_email: str, subject: str, html_body: str) -> tuple[bool, str]:
    sender, password = _smtp_creds()
    if not sender or not password:
        return False, "SMTP credentials not configured in secrets.toml"
    return _smtp_send(sender, password, to_email, subject, html_body)


def generate_otp() -> str:
    return str(random.randint(100000, 999999))


def send_otp_email(to_email: str, user_name: str, otp: str) -> tuple[bool, str]:
    subject = "Qualesce – Password Reset Code"
    body = f"""
<div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
     max-width:480px;margin:0 auto;padding:32px;background:#F8FAFC;border-radius:12px">
  <div style="font-size:28px;font-weight:900;color:#3B82F6;letter-spacing:-1px;margin-bottom:4px">Q</div>
  <h2 style="color:#0F172A;margin:0 0 12px">Password Reset Request</h2>
  <p style="color:#475569;margin:0 0 20px">Hi <b>{user_name}</b>,<br>
  Use the code below to reset your Qualesce password. It expires in <b>10 minutes</b>.</p>
  <div style="font-size:40px;font-weight:900;letter-spacing:12px;color:#3B82F6;text-align:center;
       background:#EFF6FF;border:2px dashed #BFDBFE;border-radius:10px;padding:24px;margin:0 0 20px">
    {otp}
  </div>
  <p style="color:#94A3B8;font-size:12px;margin:0">
    If you did not request a password reset, you can safely ignore this email.
  </p>
  <p style="color:#94A3B8;font-size:12px;margin:8px 0 0">
    <a href="{PORTAL_URL}" style="color:#3B82F6">Back to Qualesce Portal</a>
  </p>
</div>
"""
    return send_email(to_email, subject, body)


def send_task_assigned_email(emp_email: str, emp_name: str, task_title: str,
                              assigned_by: str, due_date: str,
                              sender_email: str = "", sender_password: str = "") -> tuple[bool, str]:
    subject = f"Qualesce – New Task Assigned: {task_title}"
    due_line = f"Due: <b>{due_date}</b>" if due_date else "No due date set"
    body = f"""
<div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
     max-width:480px;margin:0 auto;padding:32px;background:#F8FAFC;border-radius:12px">
  <div style="font-size:28px;font-weight:900;color:#3B82F6;letter-spacing:-1px;margin-bottom:4px">Q</div>
  <h2 style="color:#0F172A;margin:0 0 12px">New Task Assigned to You</h2>
  <p style="color:#475569;margin:0 0 16px">Hi <b>{emp_name}</b>, a new task has been assigned to you.</p>
  <div style="background:#fff;border:1px solid #E2E8F0;border-radius:10px;padding:18px 22px;margin:0 0 16px">
    <div style="font-size:17px;font-weight:700;color:#0F172A;margin-bottom:8px">{task_title}</div>
    <div style="color:#64748B;font-size:13px;margin-bottom:4px">Assigned by: <b>{assigned_by}</b></div>
    <div style="color:#64748B;font-size:13px">{due_line}</div>
  </div>
  <p style="color:#94A3B8;font-size:12px;margin:0">
    Log in to <a href="{PORTAL_URL}" style="color:#3B82F6;font-weight:600">Qualesce Portal</a> to view your task details and update your progress.
  </p>
</div>
"""
    if sender_email and sender_password:
        ok, err = _smtp_send(sender_email, sender_password, emp_email, subject, body)
        if ok:
            return ok, err
        # SMTP AUTH disabled for this account — fall back to global credentials
        # Reply-To is set to the intended sender so employee replies reach them directly
        global_sender, global_pwd = _smtp_creds()
        if global_sender and global_pwd and global_sender != sender_email:
            return _smtp_send(global_sender, global_pwd, emp_email, subject, body,
                              reply_to=sender_email,
                              display_name=f"{assigned_by}")
        return ok, err
    return send_email(emp_email, subject, body)


def send_task_updated_email(assigner_email: str, assigner_name: str, emp_name: str,
                             task_title: str, new_status: str, progress: int,
                             comment: str) -> tuple[bool, str]:
    subject = f"Qualesce – Task Update: {task_title}"
    comment_block = (
        f'<div style="background:#F8FAFC;border-left:3px solid #CBD5E1;padding:10px 14px;'
        f'border-radius:0 6px 6px 0;color:#475569;font-size:13px;font-style:italic;margin-top:10px">'
        f'"{comment}"</div>'
    ) if comment.strip() else ""
    body = f"""
<div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
     max-width:480px;margin:0 auto;padding:32px;background:#F8FAFC;border-radius:12px">
  <div style="font-size:28px;font-weight:900;color:#3B82F6;letter-spacing:-1px;margin-bottom:4px">Q</div>
  <h2 style="color:#0F172A;margin:0 0 12px">Task Updated</h2>
  <p style="color:#475569;margin:0 0 16px">Hi <b>{assigner_name}</b>,
  <b>{emp_name}</b> has updated a task you assigned.</p>
  <div style="background:#fff;border:1px solid #E2E8F0;border-radius:10px;padding:18px 22px;margin:0 0 16px">
    <div style="font-size:17px;font-weight:700;color:#0F172A;margin-bottom:8px">{task_title}</div>
    <div style="display:flex;gap:16px;flex-wrap:wrap">
      <span style="background:#DBEAFE;color:#1D4ED8;padding:3px 10px;border-radius:99px;
            font-size:12px;font-weight:600">{new_status}</span>
      <span style="background:#DCFCE7;color:#16A34A;padding:3px 10px;border-radius:99px;
            font-size:12px;font-weight:600">{progress}% complete</span>
    </div>
    {comment_block}
  </div>
  <p style="color:#94A3B8;font-size:12px;margin:0">
    Log in to <a href="{PORTAL_URL}" style="color:#3B82F6;font-weight:600">Qualesce Portal</a> to view the full update and task history.
  </p>
</div>
"""
    return send_email(assigner_email, subject, body)
