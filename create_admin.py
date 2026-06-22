"""
初期管理者ユーザーを作成するスクリプト。

使い方:
    python3 create_admin.py [username] [password]

デフォルト:
    username: admin
    password: changeme (初回ログイン後に変更してください)
"""

import sys
from pathlib import Path

# streamlit secrets を読むために必要
import streamlit as st
st.secrets  # noqa: trigger secrets loading

from auth import create_user, list_users


def main():
    username = sys.argv[1] if len(sys.argv) > 1 else "admin"
    password = sys.argv[2] if len(sys.argv) > 2 else "changeme"

    existing = list_users()
    if any(u["username"] == username for u in existing):
        print(f"ユーザー '{username}' は既に存在します。")
        return

    ok = create_user(username, "管理者", password, "admin")
    if ok:
        print(f"管理者ユーザー '{username}' を作成しました。")
        print("初回ログイン後にパスワードを変更してください。")
    else:
        print("作成に失敗しました。Supabase接続設定を確認してください。")


if __name__ == "__main__":
    main()
