"""
db.py
Supabase を使ったプロジェクト永続ストレージ。

テーブル定義（Supabaseのコンソールで1度だけ実行）:

    CREATE TABLE projects (
        id          uuid        DEFAULT gen_random_uuid() PRIMARY KEY,
        kojyo_name  text        UNIQUE NOT NULL,
        sheets      jsonb       NOT NULL DEFAULT '{}',
        updated_at  timestamptz DEFAULT now()
    );

st.secrets の設定例（.streamlit/secrets.toml）:

    [supabase]
    url = "https://xxxxxxxxxx.supabase.co"
    key = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9..."
"""

from __future__ import annotations

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# クライアント（毎回新規作成 ― secretsの読み込みタイミング問題を回避）
# ---------------------------------------------------------------------------

def _client():
    """Supabase クライアントを返す。未設定・エラーなら None。"""
    try:
        from supabase import create_client
        url = st.secrets["supabase"]["url"]
        key = st.secrets["supabase"]["key"]
        return create_client(url, key)
    except KeyError:
        return None          # secrets.toml に [supabase] がない
    except Exception:
        return None


def is_available() -> bool:
    """Supabase が設定・接続可能なら True。"""
    try:
        from supabase import create_client
        url = st.secrets["supabase"]["url"]
        key = st.secrets["supabase"]["key"]
        return bool(url and key)
    except Exception:
        return False


# ---------------------------------------------------------------------------
# シリアライズ / デシリアライズ
# ---------------------------------------------------------------------------

def _dfs_to_json(sheets: dict[str, pd.DataFrame]) -> dict:
    return {k: df.to_dict(orient="records") for k, df in sheets.items()}


def _json_to_dfs(raw: dict) -> dict[str, pd.DataFrame]:
    return {k: pd.DataFrame(v) for k, v in raw.items()}


# ---------------------------------------------------------------------------
# CRUD
# ---------------------------------------------------------------------------

def save_project(kojyo_name: str, sheets: dict[str, pd.DataFrame]) -> None:
    """プロジェクトを保存（存在すれば上書き）。"""
    client = _client()
    if client is None:
        return
    payload = {
        "kojyo_name": kojyo_name,
        "sheets": _dfs_to_json(sheets),
    }
    client.table("projects").upsert(payload, on_conflict="kojyo_name").execute()


def load_project(kojyo_name: str) -> dict[str, pd.DataFrame] | None:
    """プロジェクトを読み込む。見つからなければ None。"""
    client = _client()
    if client is None:
        return None
    res = (
        client.table("projects")
        .select("sheets")
        .eq("kojyo_name", kojyo_name)
        .limit(1)
        .execute()
    )
    if res.data:
        return _json_to_dfs(res.data[0]["sheets"])
    return None


def list_projects() -> list[str]:
    """保存済みプロジェクト名の一覧を返す（更新日時降順）。"""
    client = _client()
    if client is None:
        return []
    res = (
        client.table("projects")
        .select("kojyo_name, updated_at")
        .order("updated_at", desc=True)
        .execute()
    )
    return [row["kojyo_name"] for row in res.data]


def delete_project(kojyo_name: str) -> None:
    """プロジェクトを削除する。"""
    client = _client()
    if client is None:
        return
    client.table("projects").delete().eq("kojyo_name", kojyo_name).execute()


def rename_project(old_name: str, new_name: str) -> bool:
    """プロジェクト名を変更する。成功したら True を返す。"""
    if old_name == new_name or not new_name.strip():
        return False
    sheets = load_project(old_name)
    if sheets is None:
        return False
    save_project(new_name, sheets)
    delete_project(old_name)
    return True
