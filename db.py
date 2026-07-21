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


def list_projects(user_id: str | None = None, role: str | None = None) -> list[dict]:
    """
    保存済みプロジェクト一覧を返す（更新日時降順）。

    user_id/role が指定された場合:
      - admin: 全プロジェクト
      - editor/viewer: project_permissions で許可されたもののみ

    Returns:
        [{"id": ..., "kojyo_name": ..., "updated_at": ...}, ...]
    """
    client = _client()
    if client is None:
        return []

    try:
        if role == "admin" or user_id is None:
            res = (
                client.table("projects")
                .select("id, kojyo_name, updated_at")
                .order("updated_at", desc=True)
                .execute()
            )
            return res.data

        # editor/viewer: 権限テーブルでフィルタ
        from auth import get_accessible_project_ids
        allowed_ids = get_accessible_project_ids(user_id)
        if not allowed_ids:
            return []
        res = (
            client.table("projects")
            .select("id, kojyo_name, updated_at")
            .in_("id", allowed_ids)
            .order("updated_at", desc=True)
            .execute()
        )
        return res.data
    except Exception:
        return []


def list_project_names(user_id: str | None = None, role: str | None = None) -> list[str]:
    """list_projects() の後方互換ラッパー。名前のリストのみ返す。"""
    return [r["kojyo_name"] for r in list_projects(user_id=user_id, role=role)]


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


def get_project_id(kojyo_name: str) -> str | None:
    """プロジェクト名からIDを取得する。"""
    client = _client()
    if client is None:
        return None
    res = (
        client.table("projects")
        .select("id")
        .eq("kojyo_name", kojyo_name)
        .limit(1)
        .execute()
    )
    if res.data:
        return res.data[0]["id"]
    return None


# ---------------------------------------------------------------------------
# マッチングセッション（対応表の作業状態永続化）
# ---------------------------------------------------------------------------

def serialize_session(
    suryo_info: dict | None,
    df_match: pd.DataFrame | None,
    row_selections: dict,
    confirmed_keys: set,
    custom_rows: list,
    selected_idx: int | None,
    excluded_rows=None,
    suryo_df_full: pd.DataFrame | None = None,
) -> dict:
    """セッション状態をJSON化可能なdictに変換する。"""
    si_out = None
    if suryo_info is not None:
        si_out = {}
        for k, v in suryo_info.items():
            if isinstance(v, pd.DataFrame):
                si_out[k] = {"__df__": True, "data": v.to_dict(orient="records")}
            else:
                si_out[k] = v
    return {
        "suryo_info": si_out,
        "df_match": df_match.to_dict(orient="records") if df_match is not None else None,
        "row_selections": [
            {"key": list(k), "value": v}
            for k, v in row_selections.items()
        ],
        "confirmed_keys": [list(k) for k in confirmed_keys],
        "custom_rows": custom_rows,
        "selected_idx": selected_idx,
        "suryo_df_full": suryo_df_full.to_dict(orient="records") if suryo_df_full is not None else None,
    }


def deserialize_session(data: dict) -> dict:
    """JSON dictをセッション状態に復元する。"""
    # suryo_info
    si_raw = data.get("suryo_info")
    si = None
    if si_raw is not None:
        si = {}
        for k, v in si_raw.items():
            if isinstance(v, dict) and v.get("__df__"):
                si[k] = pd.DataFrame(v["data"])
            else:
                si[k] = v

    # df_match
    dm_raw = data.get("df_match")
    dm = pd.DataFrame(dm_raw) if dm_raw is not None else None

    # row_selections: list of {key, value} → dict with tuple keys
    rs = {}
    for item in data.get("row_selections", []):
        rs[tuple(item["key"])] = item["value"]

    # confirmed_keys: list of lists → set of tuples
    ck = {tuple(k) for k in data.get("confirmed_keys", [])}

    # suryo_df_full
    sdf_raw = data.get("suryo_df_full")
    sdf = pd.DataFrame(sdf_raw) if sdf_raw is not None else None

    return {
        "suryo_info": si,
        "df_match": dm,
        "row_selections": rs,
        "confirmed_keys": ck,
        "custom_rows": data.get("custom_rows", []),
        "selected_idx": data.get("selected_idx"),
        "suryo_df_full": sdf,
    }


def save_session(
    kojyo_name: str,
    user_id: str | None,
    state: dict,
    progress: str = "作業中",
) -> None:
    """マッチングセッションを保存（upsert）。"""
    client = _client()
    if client is None:
        return
    payload = {
        "kojyo_name": kojyo_name,
        "state": state,
        "progress": progress,
    }
    if user_id is not None:
        payload["user_id"] = user_id
    try:
        # 既存レコードを確認して update or insert
        q = client.table("matching_sessions").select("id").eq("kojyo_name", kojyo_name)
        if user_id is not None:
            q = q.eq("user_id", user_id)
        else:
            q = q.is_("user_id", "null")
        existing = q.limit(1).execute()
        if existing.data:
            client.table("matching_sessions").update(payload).eq(
                "id", existing.data[0]["id"]
            ).execute()
        else:
            client.table("matching_sessions").insert(payload).execute()
    except Exception:
        pass


def load_session(kojyo_name: str, user_id: str | None = None) -> dict | None:
    """マッチングセッションを読み込む。"""
    client = _client()
    if client is None:
        return None
    try:
        q = client.table("matching_sessions").select("state").eq("kojyo_name", kojyo_name)
        if user_id is not None:
            q = q.eq("user_id", user_id)
        else:
            q = q.is_("user_id", "null")
        res = q.limit(1).execute()
        if res.data:
            return res.data[0]["state"]
    except Exception:
        pass
    return None


def list_sessions(user_id: str | None = None, role: str | None = None) -> list[dict]:
    """保存済みセッション一覧を返す（更新日時降順）。"""
    client = _client()
    if client is None:
        return []
    try:
        q = client.table("matching_sessions").select(
            "id, kojyo_name, progress, updated_at"
        )
        if user_id is not None and role not in ("admin", None):
            q = q.eq("user_id", user_id)
        res = q.order("updated_at", desc=True).execute()
        return res.data
    except Exception:
        return []


def delete_session(kojyo_name: str, user_id: str | None = None) -> None:
    """マッチングセッションを削除する。"""
    client = _client()
    if client is None:
        return
    try:
        q = client.table("matching_sessions").delete().eq("kojyo_name", kojyo_name)
        if user_id is not None:
            q = q.eq("user_id", user_id)
        else:
            q = q.is_("user_id", "null")
        q.execute()
    except Exception:
        pass
