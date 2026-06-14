"""Authentication, password management, and login UI."""

from __future__ import annotations

import json
import os
from typing import Tuple

import bcrypt
import streamlit as st

from .settings import T, USERS_FILE

def hash_password(plain: str) -> str:
    return bcrypt.hashpw(plain.encode("utf-8"), bcrypt.gensalt(rounds=12)).decode("utf-8")

def verify_password(plain: str, hashed: str) -> bool:
    try:
        return bcrypt.checkpw(plain.encode("utf-8"), hashed.encode("utf-8"))
    except Exception:
        return False

def load_users() -> dict:
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    # First run: create hashed default
    default_hash = hash_password("databridge2026")
    users = {"admin": default_hash}
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)
    return users

def save_users(users: dict) -> None:
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)

def check_login(username: str, password: str) -> bool:
    users = load_users()
    hashed = users.get(username.strip(), "")
    if not hashed:
        return False
    # Support plain-text legacy passwords (auto-upgrade on login)
    if hashed.startswith("$2b$") or hashed.startswith("$2a$"):
        return verify_password(password, hashed)
    else:
        # Legacy plain-text — upgrade to bcrypt on successful login
        if hashed == password:
            users[username] = hash_password(password)
            save_users(users)
            return True
        return False

def change_password(username: str, old_pw: str, new_pw: str) -> Tuple[bool, str]:
    """Returns (success, message)"""
    if not check_login(username, old_pw):
        return False, "❌ كلمة المرور الحالية غلط"
    if len(new_pw) < 8:
        return False, "❌ كلمة المرور الجديدة لازم تكون 8 حروف على الأقل"
    if new_pw == old_pw:
        return False, "❌ كلمة المرور الجديدة لازم تختلف عن الحالية"
    users = load_users()
    users[username] = hash_password(new_pw)
    save_users(users)
    return True, "✅ تم تغيير كلمة المرور بنجاح"

def render_login(lang: str) -> bool:
    t = T[lang]
    st.markdown(f"""
    <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;
                min-height:70vh;gap:0.5rem;">
      <div style="text-align:center;margin-bottom:2rem;">
        <div style="font-size:2.5rem;font-weight:800;color:#fff;letter-spacing:-1px;">
          Data<span style="color:#7c6aff">Bridge</span>
        </div>
        <div style="font-size:0.9rem;color:#7c6aff;font-weight:500;">{t['app_sub']}</div>
        <div style="font-size:0.75rem;color:#444;margin-top:0.3rem;">{t['app_slogan']}</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    col_l, col_m, col_r = st.columns([1, 1.2, 1])
    with col_m:
        st.markdown(f"""
        <div style="background:#0d0d1a;border:1px solid #1e1e2e;border-radius:16px;
                    padding:2rem 2rem 1.5rem 2rem;margin-bottom:1rem;">
          <div style="font-size:1rem;font-weight:600;color:#e0e0f0;margin-bottom:1.5rem;
                      text-align:center;">{t['login_title']}</div>
        </div>
        """, unsafe_allow_html=True)
        username = st.text_input(t["username"], placeholder="admin", key="login_user")
        password = st.text_input(t["password"], type="password", placeholder="••••••••", key="login_pass")
        if st.button(t["login_btn"], use_container_width=True, key="login_submit"):
            if check_login(username, password):
                st.session_state["authenticated"] = True
                st.session_state["current_user"]  = username
                st.rerun()
            else:
                st.error(t["login_error"])
    return st.session_state.get("authenticated", False)


# ════════════════════════════════════════
