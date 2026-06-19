"""
photo_alias.py
撮影箇所マッチング用の対応表ローダー

出来形/品質管理DB工種 → 撮影箇所DB工種 の対応表を JSON から読み込み、
_normalize() 済みキーで高速にルックアップできるキャッシュを提供する。

【データ形式】 photo_alias.json
  {
    "出来形管理": { "DB工種名": ["撮影箇所DB工種名", ...] },
    "品質管理":   { "DB工種名": ["撮影箇所DB工種名", ...] }
  }

【使い方】
    from photo_alias import resolve_photo_targets

    targets = resolve_photo_targets("河川土工", "品質管理")
    # → ["河川・海岸土工 （施工）"]
"""

from __future__ import annotations

import json
import re
import unicodedata
from pathlib import Path

_JSON_PATH = Path(__file__).parent / "photo_alias.json"

# ---------------------------------------------------------------------------
# 正規化 (extractor._normalize と同一ロジック)
# ---------------------------------------------------------------------------

def _normalize(s: str) -> str:
    if not s:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[（）()【】\[\]「」『』]", "", s)
    s = re.sub(r"[・･]", "", s)
    return s.lower()


# ---------------------------------------------------------------------------
# キャッシュ
# ---------------------------------------------------------------------------

# {section: {normalized_source: [raw_target, ...]}}
_cache: dict[str, dict[str, list[str]]] | None = None


def _load() -> dict[str, dict[str, list[str]]]:
    """photo_alias.json を読み込み、正規化済みキーのキャッシュを返す。"""
    global _cache
    if _cache is not None:
        return _cache

    if not _JSON_PATH.exists():
        _cache = {}
        return _cache

    with open(_JSON_PATH, encoding="utf-8") as f:
        raw = json.load(f)

    _cache = {}
    for section in ("出来形管理", "品質管理"):
        entries = raw.get(section, {})
        norm_map: dict[str, list[str]] = {}
        for src, targets in entries.items():
            if src.startswith("_"):
                continue
            nk = _normalize(src)
            if nk:
                norm_map[nk] = targets
        _cache[section] = norm_map

    # implicit_photo: キーワード部分一致 → 撮影箇所品管セクション追加工種
    implicit_raw = raw.get("implicit_photo", {})
    implicit_map: dict[str, list[str]] = {}
    for kw, targets in implicit_raw.items():
        if kw.startswith("_"):
            continue
        nk = _normalize(kw)
        if nk:
            implicit_map[nk] = targets
    _cache["implicit_photo"] = implicit_map

    return _cache


def reload():
    """キャッシュを破棄して再読込する（UI から JSON を更新した後に呼ぶ）。"""
    global _cache
    _cache = None
    _load()


# ---------------------------------------------------------------------------
# ルックアップ
# ---------------------------------------------------------------------------

def resolve_photo_targets(source_kojyo: str, section: str) -> list[str]:
    """
    出来形/品管DB工種名から、対応する撮影箇所DB工種名のリストを返す。

    Args:
        source_kojyo: _match_chain() が返した DB 工種名
        section:      "出来形管理" or "品質管理"

    Returns:
        マッピングが存在する場合は撮影箇所DB工種名リスト、
        存在しない場合は空リスト。
    """
    data = _load()
    section_map = data.get(section, {})
    ns = _normalize(source_kojyo)
    if not ns:
        return []
    return list(section_map.get(ns, []))


def get_implicit_photo_kojyo(matched_kojyo_d: list, matched_kojyo_h: list) -> list[str]:
    """
    マッチ済みの出来形/品管工種から、撮影箇所の品質管理写真に追加すべき工種を返す。

    品管一覧には影響しない。撮影箇所の品質管理セクションでの検索用。
    implicit_photo ルールのキーワードが matched_kojyo_d/h のいずれかに
    部分一致した場合、対応するターゲット工種を返す。

    Returns:
        追加すべき品質管理DB工種名のリスト（重複なし）
    """
    data = _load()
    implicit_map = data.get("implicit_photo", {})
    if not implicit_map:
        return []

    # マッチ済み工種を全て正規化
    all_matched_norms = set()
    for k in matched_kojyo_d:
        nk = _normalize(k)
        if nk:
            all_matched_norms.add(nk)
    for k in matched_kojyo_h:
        nk = _normalize(k)
        if nk:
            all_matched_norms.add(nk)

    if not all_matched_norms:
        return []

    # キーワード部分一致でトリガー
    result: list[str] = []
    seen: set[str] = set()
    for kw_norm, targets in implicit_map.items():
        if any(kw_norm in mn or mn in kw_norm for mn in all_matched_norms):
            for t in targets:
                nt = _normalize(t)
                if nt and nt not in seen:
                    result.append(t)
                    seen.add(nt)

    return result
