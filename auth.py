"""
auth.py
ユーザー認証・管理・プロジェクト閲覧権限。

テーブル定義（Supabaseのコンソールで1度だけ実行）:

    CREATE TABLE users (
        id            uuid        DEFAULT gen_random_uuid() PRIMARY KEY,
        username      text        UNIQUE NOT NULL,
        display_name  text        NOT NULL DEFAULT '',
        password_hash text        NOT NULL,
        role          text        NOT NULL DEFAULT 'viewer'
                                  CHECK (role IN ('admin', 'editor', 'viewer')),
        created_at    timestamptz DEFAULT now(),
        updated_at    timestamptz DEFAULT now()
    );

    CREATE TABLE project_permissions (
        id          uuid        DEFAULT gen_random_uuid() PRIMARY KEY,
        user_id     uuid        NOT NULL REFERENCES users(id) ON DELETE CASCADE,
        project_id  uuid        NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
        created_at  timestamptz DEFAULT now(),
        UNIQUE(user_id, project_id)
    );

初期管理者の作成:
    python3 -c "from auth import create_user; create_user('admin', '管理者', 'changeme', 'admin')"
"""

from __future__ import annotations

import hashlib
import hmac
import os

import streamlit as st


# ---------------------------------------------------------------------------
# Supabase クライアント（db.py と同じパターン）
# ---------------------------------------------------------------------------

def _client():
    try:
        from supabase import create_client
        url = st.secrets["supabase"]["url"]
        key = st.secrets["supabase"]["key"]
        return create_client(url, key)
    except Exception:
        return None


# ---------------------------------------------------------------------------
# パスワードハッシュ（bcrypt不要 — PBKDF2-SHA256で十分）
# ---------------------------------------------------------------------------

_ITERATIONS = 260_000  # OWASP推奨


def _hash_password(plain: str) -> str:
    """PBKDF2-SHA256でハッシュ化。salt:hash の形式で返す。"""
    salt = os.urandom(16).hex()
    dk = hashlib.pbkdf2_hmac("sha256", plain.encode(), salt.encode(), _ITERATIONS)
    return f"{salt}:{dk.hex()}"


def _verify_password(plain: str, stored: str) -> bool:
    """保存済みハッシュと照合する。"""
    try:
        salt, hash_hex = stored.split(":", 1)
        dk = hashlib.pbkdf2_hmac("sha256", plain.encode(), salt.encode(), _ITERATIONS)
        return hmac.compare_digest(dk.hex(), hash_hex)
    except Exception:
        return False


# ---------------------------------------------------------------------------
# 認証
# ---------------------------------------------------------------------------

def authenticate(username: str, password: str) -> dict | None:
    """
    ユーザー認証。成功時は {id, username, display_name, role} を返す。
    """
    client = _client()
    if client is None:
        return None
    try:
        res = (
            client.table("users")
            .select("id, username, display_name, password_hash, role")
            .eq("username", username)
            .limit(1)
            .execute()
        )
        if not res.data:
            return None
        user = res.data[0]
        if not _verify_password(password, user["password_hash"]):
            return None
        return {
            "id": user["id"],
            "username": user["username"],
            "display_name": user["display_name"],
            "role": user["role"],
        }
    except Exception:
        return None


# ---------------------------------------------------------------------------
# ユーザー CRUD
# ---------------------------------------------------------------------------

def list_users() -> list[dict]:
    """全ユーザー一覧（password_hash除外）。"""
    client = _client()
    if client is None:
        return []
    try:
        res = (
            client.table("users")
            .select("id, username, display_name, role, created_at")
            .order("created_at")
            .execute()
        )
        return res.data
    except Exception:
        return []


def create_user(username: str, display_name: str, password: str, role: str) -> bool:
    """ユーザー作成。成功でTrue。"""
    client = _client()
    if client is None:
        return False
    if role not in ("admin", "editor", "viewer"):
        return False
    payload = {
        "username": username.strip(),
        "display_name": display_name.strip(),
        "password_hash": _hash_password(password),
        "role": role,
    }
    try:
        client.table("users").insert(payload).execute()
        return True
    except Exception:
        return False


def update_user(
    user_id: str,
    display_name: str | None = None,
    role: str | None = None,
    password: str | None = None,
) -> bool:
    """ユーザー情報更新。指定したフィールドのみ更新。"""
    client = _client()
    if client is None:
        return False
    payload = {}
    if display_name is not None:
        payload["display_name"] = display_name.strip()
    if role is not None and role in ("admin", "editor", "viewer"):
        payload["role"] = role
    if password is not None and password.strip():
        payload["password_hash"] = _hash_password(password.strip())
    if not payload:
        return False
    try:
        client.table("users").update(payload).eq("id", user_id).execute()
        return True
    except Exception:
        return False


def delete_user(user_id: str) -> bool:
    """ユーザー削除（CASCADE で権限も削除される）。"""
    client = _client()
    if client is None:
        return False
    try:
        client.table("users").delete().eq("id", user_id).execute()
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# プロジェクト権限
# ---------------------------------------------------------------------------

def get_accessible_project_ids(user_id: str) -> list[str]:
    """ユーザーがアクセス可能な project_id 一覧を返す。"""
    client = _client()
    if client is None:
        return []
    try:
        res = (
            client.table("project_permissions")
            .select("project_id")
            .eq("user_id", user_id)
            .execute()
        )
        return [r["project_id"] for r in res.data]
    except Exception:
        return []


def get_project_permissions(project_id: str) -> list[dict]:
    """プロジェクトにアクセス権を持つユーザー一覧。"""
    client = _client()
    if client is None:
        return []
    try:
        res = (
            client.table("project_permissions")
            .select("user_id, users(username, display_name, role)")
            .eq("project_id", project_id)
            .execute()
        )
        return res.data
    except Exception:
        return []


def add_permission(user_id: str, project_id: str) -> bool:
    """ユーザーにプロジェクト閲覧権限を付与。"""
    client = _client()
    if client is None:
        return False
    try:
        client.table("project_permissions").upsert(
            {"user_id": user_id, "project_id": project_id},
            on_conflict="user_id,project_id",
        ).execute()
        return True
    except Exception:
        return False


def remove_permission(user_id: str, project_id: str) -> bool:
    """ユーザーからプロジェクト閲覧権限を削除。"""
    client = _client()
    if client is None:
        return False
    try:
        (client.table("project_permissions")
         .delete()
         .eq("user_id", user_id)
         .eq("project_id", project_id)
         .execute())
        return True
    except Exception:
        return False


def set_user_permissions(user_id: str, project_ids: list[str]) -> None:
    """ユーザーの権限を一括設定（既存を置き換え）。"""
    client = _client()
    if client is None:
        return
    # 既存削除
    client.table("project_permissions").delete().eq("user_id", user_id).execute()
    # 新規挿入
    if project_ids:
        rows = [{"user_id": user_id, "project_id": pid} for pid in project_ids]
        client.table("project_permissions").insert(rows).execute()


def is_available() -> bool:
    """usersテーブルが使用可能か確認。"""
    try:
        client = _client()
        if client is None:
            return False
        client.table("users").select("id").limit(1).execute()
        return True
    except BaseException:
        return False
