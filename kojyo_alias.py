"""
kojyo_alias.py
工種名の対応表ローダー

出来形/品管の _match_chain() で部分一致が0件の場合のフォールバック。
数量総括表の工種名 → DB工種名 のマッピングを JSON から読み込む。

【カバーする4パターン】
  1. 別名称（同一工法）       : 「中層混合処理」 ↔ 「固結工（中層混合処理）」
  2. 合算出力（B分割・A合算） : 「補償関係」「環境対策・…」 → 「補償関係環境対策イメージアップ等」
  3. 末尾「工」省略＋括弧詳細 : 「アスファルト舗装」 → 「アスファルト舗装工（表層工）」等
  4. 括弧付き細分（Bが親）    : 「側溝工」 → 「側溝工（場所打水路工）」等

【使い方】
    from kojyo_alias import ALIAS_B_TO_A, b_covered_by_a, reload

    covered = b_covered_by_a("アスファルト舗装", a_kojyo_set)
"""

from __future__ import annotations

import json
import re
import unicodedata
from pathlib import Path

_JSON_PATH = Path(__file__).parent / "kojyo_alias.json"


def _n(s: str) -> str:
    """正規化: 半角カナ→全角、全角スペース除去、括弧統一（全角→半角）"""
    s = unicodedata.normalize("NFKC", str(s))
    s = re.sub(r"[\s\u3000\r\n]+", "", s)
    return s.strip()


# ---------------------------------------------------------------------------
# データ読み込み
# ---------------------------------------------------------------------------

ALIAS_B_TO_A: dict[str, list[str]] = {}
_norm_alias: dict[str, list[str]] | None = None


def _load() -> dict[str, list[str]]:
    """kojyo_alias.json を読み込み、ALIAS_B_TO_A を更新する。"""
    global ALIAS_B_TO_A, _norm_alias
    if not _JSON_PATH.exists():
        ALIAS_B_TO_A = {}
        _norm_alias = None
        return ALIAS_B_TO_A

    with open(_JSON_PATH, encoding="utf-8") as f:
        raw = json.load(f)

    ALIAS_B_TO_A = {k: v for k, v in raw.items() if not k.startswith("_")}
    _norm_alias = None
    return ALIAS_B_TO_A


def reload():
    """キャッシュを破棄して再読込する（UI から JSON を更新した後に呼ぶ）。"""
    _load()


# 初回ロード
_load()


# ---------------------------------------------------------------------------
# ルックアップ関数
# ---------------------------------------------------------------------------

def _build_norm_alias() -> dict[str, list[str]]:
    """ALIAS_B_TO_A を正規化済みキーで再構築したキャッシュを返す"""
    result: dict[str, list[str]] = {}
    for b_key, a_vals in ALIAS_B_TO_A.items():
        nk = _n(b_key)
        result[nk] = [_n(v) for v in a_vals]
    return result


def b_covered_by_a(b_kojyo: str, a_kojyo_set: set[str]) -> bool:
    """
    Bの工種名がAの工種セットに包含されているかを判定する。
    直接一致に加え、ALIAS_B_TO_A の対応表も考慮する。

    Args:
        b_kojyo      : 目標outputの工種名
        a_kojyo_set  : アプリ出力の工種名セット

    Returns:
        True  → B工種はAに包含されている（直接 or エイリアス経由）
        False → B工種はAに存在しない
    """
    global _norm_alias
    if _norm_alias is None:
        _norm_alias = _build_norm_alias()

    norm_a = {_n(a) for a in a_kojyo_set}
    nb = _n(b_kojyo)

    # 直接一致
    if nb in norm_a:
        return True

    # エイリアス経由
    for alias_n in _norm_alias.get(nb, []):
        if alias_n and alias_n in norm_a:
            return True

    return False


def covered_stats(
    b_kojyo_list: list[str],
    a_kojyo_set: set[str],
) -> dict:
    """
    B工種リスト全体の包含状況を集計して返す。

    Returns:
        {
          "matched"  : 一致した工種リスト,
          "missing"  : 未一致の工種リスト,
          "rate"     : 包含率 (0.0 〜 1.0),
        }
    """
    matched, missing = [], []
    for b in b_kojyo_list:
        (matched if b_covered_by_a(b, a_kojyo_set) else missing).append(b)
    total = len(b_kojyo_list)
    return {
        "matched": matched,
        "missing": missing,
        "rate":    len(matched) / total if total else 0.0,
    }
