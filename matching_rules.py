"""
matching_rules.py
マッチングルール設定のローダー

除外リスト・暗黙追加ルール・表示名リマップを matching_rules.json から読み込む。
従来 extractor.py / excel_writer.py にハードコードされていたルールを
ユーザーが編集可能な JSON に外出しする。

【データ形式】 matching_rules.json
  {
    "hinshitsu_exclude_always":        { "keywords": [...] },
    "hinshitsu_exclude_unless_in_suryo": { "keywords": [...] },
    "dekigata_exclude_unless_in_suryo":  { "keywords": [...] },
    "implicit_dekigata":   { "トリガー": ["追加工種", ...] },
    "implicit_hinshitsu":  { "トリガー": ["追加工種", ...] },
    "hinshitsu_kojyo_display": { "DB工種名": "表示名" }
  }
"""

from __future__ import annotations

import json
from pathlib import Path

_JSON_PATH = Path(__file__).parent / "matching_rules.json"

# ---------------------------------------------------------------------------
# キャッシュ
# ---------------------------------------------------------------------------

_cache: dict | None = None


def _load() -> dict:
    """matching_rules.json を読み込みキャッシュする。"""
    global _cache
    if _cache is not None:
        return _cache

    if not _JSON_PATH.exists():
        _cache = {}
        return _cache

    with open(_JSON_PATH, encoding="utf-8") as f:
        _cache = json.load(f)

    return _cache


def reload():
    """キャッシュを破棄して再読込する（UI から JSON を更新した後に呼ぶ）。"""
    global _cache
    _cache = None
    _load()


# ---------------------------------------------------------------------------
# ルックアップ
# ---------------------------------------------------------------------------

def _keys_only(section: dict) -> list[str]:
    """セクション辞書からキーのみを抽出（_comment等を除外）。"""
    return [k for k in section if not k.startswith("_")]


def get_hinshitsu_exclude_always() -> list[str]:
    """品質管理から常に除外する工種キーワードリストを返す。"""
    data = _load()
    return _keys_only(data.get("hinshitsu_exclude_always", {}))


def get_hinshitsu_exclude_unless_in_suryo() -> list[str]:
    """数量総括表に言及がない場合のみ品質管理から除外するキーワードリストを返す。"""
    data = _load()
    return _keys_only(data.get("hinshitsu_exclude_unless_in_suryo", {}))


def get_dekigata_exclude_unless_in_suryo() -> list[str]:
    """数量総括表に言及がない場合のみ出来形・撮影箇所から除外するキーワードリストを返す。"""
    data = _load()
    return _keys_only(data.get("dekigata_exclude_unless_in_suryo", {}))


def get_implicit_dekigata_rules() -> dict[str, list[str]]:
    """暗黙出来形追加ルールを返す。{トリガー: [追加工種, ...]}"""
    data = _load()
    section = data.get("implicit_dekigata", {})
    return {k: v for k, v in section.items() if not k.startswith("_")}


def get_implicit_hinshitsu_rules() -> dict[str, list[str]]:
    """暗黙品管追加ルールを返す。{トリガー: [追加工種, ...]}"""
    data = _load()
    section = data.get("implicit_hinshitsu", {})
    return {k: v for k, v in section.items() if not k.startswith("_")}


def get_hinshitsu_kojyo_display() -> dict[str, str]:
    """品質管理工種の表示名リマップ辞書を返す。{DB工種名: 表示名}"""
    data = _load()
    section = data.get("hinshitsu_kojyo_display", {})
    return {k: v for k, v in section.items() if not k.startswith("_")}


def save(data: dict) -> None:
    """matching_rules.json を保存し、キャッシュを更新する。"""
    with open(_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    reload()


def load_raw() -> dict:
    """JSON の生データを返す（UI編集用）。"""
    if not _JSON_PATH.exists():
        return {}
    with open(_JSON_PATH, encoding="utf-8") as f:
        return json.load(f)


# 初回ロード
_load()
