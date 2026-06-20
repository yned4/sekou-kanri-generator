"""
match_filter.py
出来形管理・撮影箇所マッチング用の絞り込みテーブルローダー

【テーブル構成】

NARROWING_TABLE
  アルゴリズム絞り込み後も候補が複数残る場合に適用。
  キー  : 数量総括表のいずれかの階層に現れるキーワード（表記揺れは _normalize() で吸収）
  値    : そのキーワードが存在する場合に「残してよい」DB工種名リスト
  ※ 値のリストに含まれない候補は除外される。完全除外を避けるため、
    絞り込み後が空になる場合は絞り込みを適用しない。
"""

from __future__ import annotations

import json
from pathlib import Path

_JSON_PATH = Path(__file__).parent / "match_filter.json"

# ---------------------------------------------------------------------------
# データ読み込み
# ---------------------------------------------------------------------------

NARROWING_TABLE: dict[str, list[str]] = {}


def _load() -> dict[str, list[str]]:
    """match_filter.json を読み込み、NARROWING_TABLE を更新する。"""
    global NARROWING_TABLE
    if not _JSON_PATH.exists():
        NARROWING_TABLE = {}
        return NARROWING_TABLE

    with open(_JSON_PATH, encoding="utf-8") as f:
        raw = json.load(f)

    NARROWING_TABLE = {k: v for k, v in raw.items() if not k.startswith("_")}
    return NARROWING_TABLE


def reload():
    """キャッシュを破棄して再読込する（UI から JSON を更新した後に呼ぶ）。"""
    _load()


# 初回ロード
_load()
