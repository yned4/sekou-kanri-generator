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
# クライアント（アプリ全体で1インスタンス）
# ---------------------------------------------------------------------------

@st.cache_resource
def _client():
    """Supabase クライアントを返す。未設定なら None。"""
    try:
        from supabase import create_client
        url = st.secrets["supabase"]["url"]
        key = st.secrets["supabase"]["key"]
        return create_client(url, key)
    except Exception:
        return None


def is_available() -> bool:
    """Supabase が設定・接続済みなら True。"""
    return _client() is not None


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
        "updated_at": "now()",
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
