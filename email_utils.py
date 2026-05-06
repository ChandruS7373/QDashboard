import smtplib
import random
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def _smtp_creds():
    # 1. Try database settings (configured via the admin UI)
    try:
        import auth as _auth
        cfg = _auth.get_email_settings()
        if cfg["outlook_email"] and cfg["outlook_password"]:
            return cfg["outlook_email"], cfg["outlook_password"]
    except Exception:
        pass
    # 2. Fall back to secrets.toml / environment variable
    try:
        import streamlit as st
        email = st.secrets.get("OUTLOOK_EMAIL", "")
        pwd   = st.secrets.get("OUTLOOK_PASSWORD", "")
        if email and pwd:
            return email, pwd
    except Exception:
        pass
    return os.environ.get("OUTLOOK_EMAIL", ""), os.environ.get("OUTLOOK_PASSWORD", "")


def _smtp_send(sender: str, password: str, to_email: str, subject: str, html_body: str) -> tuple[bool, str]:
    msg = MIMEMultipart("alternative")
    msg["From"] = sender
    msg["To"] = to_email
    msg["Subject"] = subject
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
    Log in to Qualesce to view your task details and update your progress.
  </p>
</div>
"""
    if sender_email and sender_password:
        return _smtp_send(sender_email, sender_password, emp_email, subject, body)
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
    Log in to Qualesce to view the full update and task history.
  </p>
</div>
"""
    return send_email(assigner_email, subject, body)
