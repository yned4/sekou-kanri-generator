"""
photo_alias.py
撮影箇所マッチング用の対応表ローダー

数量総括表の工種名 → 撮影箇所DB工種名 の対応表を JSON から読み込み、
_match_chain() のフォールバックとして使用する（kojyo_alias.json と同じ役割）。

【データ形式】 photo_alias.json
  {
    "出来形管理": { "数量総括表の工種名": ["撮影箇所DB工種名", ...] },
    "品質管理":   { "数量総括表の工種名": ["撮影箇所DB工種名", ...] }
  }

【使い方】
    from photo_alias import get_norm_alias_for_match_chain

    alias = get_norm_alias_for_match_chain("出来形管理")
    # → {normalized_key: {normalized_target, ...}}
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


def get_norm_alias_for_match_chain(section: str) -> dict[str, set[str]]:
    """
    _match_chain() の norm_alias_override として使用できる形式でエイリアスを返す。

    Returns:
        {normalized_source: {normalized_target1, normalized_target2, ...}}
    """
    data = _load()
    section_map = data.get(section, {})
    result: dict[str, set[str]] = {}
    for norm_src, raw_targets in section_map.items():
        result[norm_src] = {_normalize(t) for t in raw_targets if _normalize(t)}
    return result


def get_implicit_photo_from_suryo(suryo_df, photo_df) -> list[str]:
    """
    数量総括表のキーワードから、撮影箇所の品質管理セクションに追加すべき工種を返す。

    _get_implicit_hinshitsu と同じ方式: 数量総括表のキーワードを直接スキャンして
    implicit_photo ルールのトリガーに部分一致したら対応する工種を追加。

    Returns:
        撮影箇所DB品質管理セクションに存在する工種名のリスト
    """
    data = _load()
    implicit_map = data.get("implicit_photo", {})
    if not implicit_map:
        return []

    # 数量総括表の全ユニーク値を正規化
    all_suryo_norm: set[str] = set()
    for col in suryo_df.columns:
        for v in suryo_df[col].unique():
            n = _normalize(str(v or ""))
            if len(n) >= 2:
                all_suryo_norm.add(n)

    if not all_suryo_norm:
        return []

    # 撮影箇所DB品質セクションの工種セット（存在確認用）
    photo_h_kojyo_norms: set[str] = set()
    for v in photo_df[photo_df["セクション"] == "品質管理"]["工種"].unique():
        nv = _normalize(str(v or ""))
        if nv:
            photo_h_kojyo_norms.add(nv)

    # キーワード部分一致でトリガー
    result: list[str] = []
    seen: set[str] = set()
    for kw_norm, targets in implicit_map.items():
        hit = any(kw_norm in sv or sv in kw_norm for sv in all_suryo_norm)
        if hit:
            for t in targets:
                nt = _normalize(t)
                if nt and nt not in seen:
                    # 撮影箇所DBに存在するか確認（部分一致）
                    if any(nt in pn or pn in nt for pn in photo_h_kojyo_norms):
                        result.append(t)
                        seen.add(nt)

    return result


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
