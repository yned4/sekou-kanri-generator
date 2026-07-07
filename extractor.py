"""
extractor.py
PDFから施工管理計画に必要なデータを抽出する。

使い方（CLI確認）:
    python extractor.py <施工管理基準.pdf> <写真管理基準.pdf>
"""

import re
import sys
import pdfplumber
import pandas as pd

# ===== ページ範囲定数（1-indexed） =====
# 土木工事施工管理基準及び規格値（案）.pdf
DEKIGATA_START_PAGE  = 21   # 出来形管理基準 本文開始
DEKIGATA_END_PAGE    = 211  # 出来形管理基準 本文終了
HINSHITSU_START_PAGE = 214  # 品質管理基準 本文開始
HINSHITSU_END_PAGE   = 268  # 品質管理基準 本文終了

# 写真管理基準.pdf
PHOTO_ZENTAI_START    = 6   # 撮影箇所一覧表（全体）開始
PHOTO_ZENTAI_END      = 7   # 撮影箇所一覧表（全体）終了
PHOTO_HINSHITSU_START = 8   # 撮影箇所一覧表（品質管理）開始
PHOTO_HINSHITSU_END   = 13  # 撮影箇所一覧表（品質管理）終了
PHOTO_DEKIGATA_START  = 14  # 撮影箇所一覧表（出来形管理）開始
PHOTO_DEKIGATA_END    = 74  # 撮影箇所一覧表（出来形管理）終了

# ===== 有効列数フィルタ =====
# ページによっては図や注釈がテーブルとして誤認識される。列数で本文テーブルのみ抽出する。
DEKIGATA_VALID_COL_COUNTS  = {11, 12, 13, 14}  # 11=規格値1列, 12=標準, 13=面管理, 14=規格値4分割
HINSHITSU_VALID_COL_COUNTS = {9}

# ===== 除外・暗黙追加ルール =====
# matching_rules.json から読み込む（対応表編集UIまたはJSON直接編集で変更可能）
from matching_rules import (
    get_hinshitsu_exclude_always,
    get_hinshitsu_exclude_unless_in_suryo,
    get_dekigata_exclude_unless_in_suryo,
    get_implicit_dekigata_rules,
    get_implicit_hinshitsu_rules,
)

PHOTO_5COL_VALID            = {5}       # 全体・品質管理セクション
PHOTO_9COL_VALID            = {9}       # 出来形管理セクション

# ===== 列名定義 =====
# 出来形管理: 13列に正規化（12列テーブルは「規格値_個々」列を空で補完）
DEKIGATA_COLS = [
    "編", "章", "節", "条", "枝番", "工種",
    "測定項目", "規格値_条件", "規格値", "規格値_個々",
    "測定基準", "測定箇所", "摘要",
]
HINSHITSU_COLS = [
    "工種", "種別", "試験区分", "試験項目", "試験方法",
    "規格値", "試験時期・頻度", "摘要", "試験成績表等による確認",
]
PHOTO_ZENTAI_COLS    = ["区分", "sub区分", "撮影項目", "撮影頻度", "摘要"]
PHOTO_HINSHITSU_COLS = ["番号", "工種", "撮影項目", "撮影頻度", "摘要"]
PHOTO_DEKIGATA_COLS  = ["編", "章", "節", "条", "枝番", "工種", "撮影項目", "撮影頻度", "摘要"]

# ===== ヘッダー行判定キーワード =====
# これらの文字列がセルに含まれる行はヘッダーとみなしてスキップする
DEKIGATA_HEADER_KEYWORDS  = {"編", "工 種", "工種", "測 定 項 目", "測定項目"}
HINSHITSU_HEADER_KEYWORDS = {"工 種", "工種", "種別", "試験\n区分"}
PHOTO_HEADER_KEYWORDS     = {"写真管理項目", "撮影項目", "撮影頻度", "撮影頻度〔時期〕", "撮影頻度[時期]"}


# ---------------------------------------------------------------------------
# 共通ユーティリティ
# ---------------------------------------------------------------------------

def _clean(value) -> str:
    """Noneを空文字に変換し、セル内改行をスペースに正規化する。"""
    if value is None:
        return ""
    return str(value).replace("\n", " ").strip()


def _is_empty_row(row: list) -> bool:
    return all(not _clean(c) for c in row)


def _row_contains(row: list, keywords: set) -> bool:
    """行内のいずれかのセルがキーワードと一致する場合 True を返す。"""
    return any(_clean(c) in keywords for c in row)


def _forward_fill(df: pd.DataFrame, cols: list) -> pd.DataFrame:
    """指定列の空文字を前の値で補完する（セル結合の復元）。"""
    df = df.copy()
    for col in cols:
        if col in df.columns:
            df[col] = df[col].replace("", pd.NA).ffill().fillna("")
    return df


def _collect_rows(pdf, start_page: int, end_page: int, valid_col_counts: set) -> list[list]:
    """
    指定ページ範囲のテーブルから、有効な列数の行をすべて収集する。
    列数が合わないテーブルは図や余白の誤認識とみなし除外する。
    """
    rows = []
    for page_num in range(start_page, end_page + 1):
        page = pdf.pages[page_num - 1]
        for table in page.extract_tables():
            if not table:
                continue
            col_count = len(table[0])
            if col_count not in valid_col_counts:
                continue
            rows.extend(table)
    return rows


# ---------------------------------------------------------------------------
# 出来形管理基準及び規格値
# ---------------------------------------------------------------------------

def extract_dekigata(施工管理基準_path: str,
                     start_page=None,
                     end_page=None) -> pd.DataFrame:
    """
    出来形管理基準及び規格値（案）を抽出する。

    ページ内に12列テーブル（標準）と13列テーブル（面管理）が混在するため、
    両方を13列に正規化して結合する。

    start_page / end_page を指定すると定数を上書きして使用する（改訂版PDF対応）。
    """
    with pdfplumber.open(施工管理基準_path) as pdf:
        total = len(pdf.pages)
        s = start_page if start_page is not None else DEKIGATA_START_PAGE
        e = end_page   if end_page   is not None else min(DEKIGATA_END_PAGE, total)
        raw_rows = _collect_rows(pdf, s, e, DEKIGATA_VALID_COL_COUNTS)

    if not raw_rows:
        raise ValueError(f"出来形管理基準: p{s}〜p{e} からテーブルを取得できませんでした。")

    cleaned = []
    for row in raw_rows:
        if _is_empty_row(row):
            continue
        if _row_contains(row[:1], DEKIGATA_HEADER_KEYWORDS):
            continue

        # 各列数 → 13列に正規化
        if len(row) == 11:
            # 規格値が1列のみ: 規格値_条件="" / 規格値=col[7] / 規格値_個々="" を補完
            row = list(row[:7]) + ["", row[7], ""] + list(row[8:])
        elif len(row) == 12:
            # 規格値が2列: 規格値_個々="" を index 9 に挿入
            row = list(row[:9]) + [""] + list(row[9:])
        elif len(row) == 14:
            # 規格値が4分割（個々・中規模/小規模, 平均・中規模/小規模）: col[10]を除去
            row = list(row[:10]) + list(row[11:])

        cleaned.append([_clean(c) for c in row[:13]])

    if not cleaned:
        raise ValueError("出来形管理基準: ヘッダー除去後にデータ行がありませんでした。")

    df = pd.DataFrame(cleaned, columns=DEKIGATA_COLS)

    # セル結合の復元: 編〜工種, 測定基準, 測定箇所, 摘要 を前方補完
    df = _forward_fill(df, ["編", "章", "節", "条", "枝番", "工種", "測定基準", "測定箇所", "摘要"])

    # 工種名のクリーニング（OCRスペース・脚注を除去）
    df["工種"] = df["工種"].apply(_clean_hinshitsu_kojyo)

    # 測定項目が空の行（ページ区切り等の残骸）を除去
    df = df[df["測定項目"].str.strip() != ""].reset_index(drop=True)

    return df


# ---------------------------------------------------------------------------
# 品質管理基準及び規格値
# ---------------------------------------------------------------------------

def _clean_hinshitsu_kojyo(val: str) -> str:
    """
    品質管理の工種名をクリーニングする。

    国交省基準PDFから抽出した工種名には以下の問題がある:
      1. 先頭に節番号が付く  例: '14 アスファ ルト舗装'
      2. 列幅制限によるOCRスペースが CJK 文字の間に挿入される
      3. 脚注（※ 以降）が連結されることがある

    処理順:
      ① 先頭の番号（半角数字 + スペース）を除去
      ② CJK 文字間の空白を除去（ASCII 文字間のスペースは保持）
      ③ ※ 以降の脚注を除去
      ④ 前後の空白を正規化
    """
    s = val.strip()
    # ① 先頭の番号を除去（例: '14 ', '2 '）
    s = re.sub(r'^\d+\s+', '', s)
    # ② CJK文字間のスペースのみ除去（JIS記号等 ASCII 間のスペースは残す）
    s = re.sub(r'(?<=[\u3000-\u9fff\uff00-\uffef])\s+(?=[\u3000-\u9fff\uff00-\uffef])', '', s)
    # ③ 脚注（※以降、改行以降の※）を除去
    s = re.sub(r'\s*※.*', '', s, flags=re.DOTALL)
    # ④ 前後のスペースを整理
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def extract_hinshitsu(施工管理基準_path: str) -> pd.DataFrame:
    """品質管理基準及び規格値（案）を抽出する。"""
    with pdfplumber.open(施工管理基準_path) as pdf:
        raw_rows = _collect_rows(pdf, HINSHITSU_START_PAGE, HINSHITSU_END_PAGE, HINSHITSU_VALID_COL_COUNTS)

    if not raw_rows:
        raise ValueError(f"品質管理基準: p{HINSHITSU_START_PAGE}〜p{HINSHITSU_END_PAGE} からテーブルを取得できませんでした。")

    cleaned = []
    for row in raw_rows:
        if _is_empty_row(row):
            continue
        if _row_contains(row[:1], HINSHITSU_HEADER_KEYWORDS):
            continue
        cleaned.append([_clean(c) for c in row[:9]])

    if not cleaned:
        raise ValueError("品質管理基準: ヘッダー除去後にデータ行がありませんでした。")

    df = pd.DataFrame(cleaned, columns=HINSHITSU_COLS)
    df = _forward_fill(df, ["工種", "種別"])

    # 工種・種別・試験区分のクリーニング（番号・OCRスペース・脚注を除去）
    df["工種"]   = df["工種"].apply(_clean_hinshitsu_kojyo)
    df["種別"]   = df["種別"].apply(_clean_hinshitsu_kojyo)
    df["試験区分"] = df["試験区分"].apply(_clean_hinshitsu_kojyo)

    # 試験項目が空の行を除去
    df = df[df["試験項目"].str.strip() != ""].reset_index(drop=True)

    return df


# ---------------------------------------------------------------------------
# 撮影箇所一覧表
# ---------------------------------------------------------------------------

def _extract_photo_section(
    pdf,
    start_page: int,
    end_page: int,
    valid_col_counts: set,
    columns: list,
) -> pd.DataFrame:
    """撮影箇所一覧表の各セクションを抽出する。"""
    raw_rows = _collect_rows(pdf, start_page, end_page, valid_col_counts)

    cleaned = []
    for row in raw_rows:
        if _is_empty_row(row):
            continue
        # 2行構成のヘッダーを両方スキップ
        if _row_contains(row, PHOTO_HEADER_KEYWORDS):
            continue
        cleaned.append([_clean(c) for c in row[: len(columns)]])

    if not cleaned:
        return pd.DataFrame(columns=columns)

    df = pd.DataFrame(cleaned, columns=columns)
    return df


def extract_photo(写真管理基準_path: str) -> pd.DataFrame:
    """
    撮影箇所一覧表（全体・品質管理・出来形管理）を抽出し、
    「セクション」列を付けて1つのDataFrameに結合する。
    """
    with pdfplumber.open(写真管理基準_path) as pdf:
        df_zentai = _extract_photo_section(
            pdf, PHOTO_ZENTAI_START, PHOTO_ZENTAI_END,
            PHOTO_5COL_VALID, PHOTO_ZENTAI_COLS,
        )
        # 区分はセル結合されているため前方補完
        df_zentai = _forward_fill(df_zentai, ["区分"])
        df_zentai.insert(0, "セクション", "全体")

        df_hinshitsu = _extract_photo_section(
            pdf, PHOTO_HINSHITSU_START, PHOTO_HINSHITSU_END,
            PHOTO_5COL_VALID, PHOTO_HINSHITSU_COLS,
        )
        df_hinshitsu = _forward_fill(df_hinshitsu, ["番号", "工種"])
        df_hinshitsu.insert(0, "セクション", "品質管理")

        df_dekigata = _extract_photo_section(
            pdf, PHOTO_DEKIGATA_START, PHOTO_DEKIGATA_END,
            PHOTO_9COL_VALID, PHOTO_DEKIGATA_COLS,
        )
        df_dekigata = _forward_fill(df_dekigata, ["編", "章", "節", "条", "枝番", "工種"])
        df_dekigata.insert(0, "セクション", "出来形管理")

    # 列が異なるセクションを結合（不足列はNaNで補完）
    df = pd.concat([df_zentai, df_hinshitsu, df_dekigata], ignore_index=True)
    return df


# ---------------------------------------------------------------------------
# 準用一覧表（p5-20）の抽出と出来形管理DB展開
# ---------------------------------------------------------------------------

# 準用一覧表の有効列数（目次テーブル）
# 7列: ヘッダー行なし [章, 条, 枝番, 工種, 種別, 準用, 頁]
# 8列: ヘッダー行あり [章, 条, 枝番, 工種, 種別, 準用, 頁, ...]
# 9列: ヘッダー行あり
_JUNYO_VALID_COL_COUNTS = {7, 8, 9}
# 準用一覧表のページ範囲
_JUNYO_START_PAGE = 5
_JUNYO_END_PAGE   = 20


def _parse_junyo_kojyo(junyo_str: str) -> str:
    """
    準用する出来形管理基準の文字列から工種名を抽出する。
    例: "1-2-4-3路体盛土工" → "路体盛土工"
         "3-2-3-17根固めブロック工" → "根固めブロック工"
    """
    s = _clean_hinshitsu_kojyo(_clean(junyo_str))
    # 先頭の節番号（数字とハイフン）を除去
    s = re.sub(r'^\d[\d\-]+', '', s).strip()
    return s


def extract_junyo_index(施工管理基準_path: str) -> list:
    """
    準用一覧表（p5-20）から (alias_工種, base_工種) ペアを抽出する。

    「準用する出来形管理基準」列が空でない行のみ対象。
    alias == base の場合（工種名が同じ）は除外する。

    Returns:
        list of (alias_工種名, base_工種名)
    """
    entries = []

    with pdfplumber.open(施工管理基準_path) as pdf:
        for p_num in range(_JUNYO_START_PAGE, _JUNYO_END_PAGE + 1):
            page = pdf.pages[p_num - 1]
            for table in page.extract_tables():
                if not table or len(table[0]) not in _JUNYO_VALID_COL_COUNTS:
                    continue

                # 列数に基づいてインデックスを決定する
                # 全テーブル共通: 工種=col[-4], 準用=col[-2]
                ncols = len(table[0])
                kojyo_idx = ncols - 4
                junyo_idx = ncols - 2

                # ヘッダー行があるテーブル(8/9列)はキーワードで上書き検証
                if ncols >= 8:
                    hdr_kojyo = hdr_junyo = None
                    for row in table[:3]:
                        cells = [_clean(c) for c in row]
                        for i, v in enumerate(cells):
                            if v in {"工種", "工 種"}:
                                hdr_kojyo = i
                            if "準用" in v:
                                hdr_junyo = i
                        if hdr_kojyo is not None and hdr_junyo is not None:
                            break
                    if hdr_kojyo is not None:
                        kojyo_idx = hdr_kojyo
                    if hdr_junyo is not None:
                        junyo_idx = hdr_junyo

                for row in table:
                    if len(row) <= max(kojyo_idx, junyo_idx):
                        continue
                    alias = _clean_hinshitsu_kojyo(_clean(row[kojyo_idx]))
                    junyo = _clean(row[junyo_idx])
                    if not alias or not junyo:
                        continue
                    if alias in {"工種", "工 種"}:
                        continue
                    base = _parse_junyo_kojyo(junyo)
                    if base and alias != base:
                        entries.append((alias, base))

    return entries


def _expand_with_junyo(df: pd.DataFrame, junyo_pairs: list) -> pd.DataFrame:
    """
    準用ペア (alias, base) に基づき、base工種の行をalias工種名で複製して追加する。

    - alias が既に df に存在する場合はスキップ
    - 同じ alias に複数の base がある場合（例: 排水構造物工→側溝工 かつ →集水桝工）は
      全ての base の行をまとめて追加する
    - base のマッチングは正規化後の部分一致で行う
    """
    existing = set(df["工種"].unique())
    norm_existing = {_normalize(k): k for k in existing}

    # alias → [base_rows, ...] で集約してから追加
    alias_rows: dict = {}

    for alias, base in junyo_pairs:
        if alias in existing:
            continue  # 既に基準値として存在する工種はスキップ

        norm_base = _normalize(base)
        matched_base = None
        if norm_base in norm_existing:
            matched_base = norm_existing[norm_base]
        else:
            for nk, k in norm_existing.items():
                if norm_base in nk or nk in norm_base:
                    matched_base = k
                    break

        if matched_base is None:
            continue

        base_rows = df[df["工種"] == matched_base].copy()
        base_rows["工種"] = alias
        alias_rows.setdefault(alias, []).append(base_rows)

    extra_rows = []
    for alias, dfs in alias_rows.items():
        combined = pd.concat(dfs, ignore_index=True).drop_duplicates(
            subset=["測定項目"], keep="first"
        )
        extra_rows.append(combined)

    if extra_rows:
        df = pd.concat([df] + extra_rows, ignore_index=True)

    return df


# ---------------------------------------------------------------------------
# メインインターフェース
# ---------------------------------------------------------------------------

def extract_from_pdf(施工管理基準_path: str,
                     写真管理基準_path: str,
                     出来形管理_path=None,
                     dekigata_start=None,
                     dekigata_end=None) -> dict:
    """
    2つのPDFからデータを抽出してdictで返す。

    出来形管理_path を指定すると、出来形管理の抽出を別PDFから行う。
    その場合、準用一覧表の展開はスキップする（別PDFには準用一覧表がない想定）。
    dekigata_start / dekigata_end でページ範囲を上書き可能。

    Returns:
        {
            "出来形管理": pd.DataFrame,
            "品質管理":   pd.DataFrame,
            "撮影箇所":   pd.DataFrame,
        }
    """
    result = {}

    print("[ 1/3 ] 出来形管理基準を抽出中...")
    if 出来形管理_path:
        result["出来形管理"] = extract_dekigata(出来形管理_path, dekigata_start, dekigata_end)
        print(f"        → {len(result['出来形管理'])} 行取得（別PDF）")
    else:
        result["出来形管理"] = extract_dekigata(施工管理基準_path)
        print(f"        → {len(result['出来形管理'])} 行取得（基準値）")
        print("        準用一覧表（各編→共通編）を展開中...")
        junyo_pairs = extract_junyo_index(施工管理基準_path)
        result["出来形管理"] = _expand_with_junyo(result["出来形管理"], junyo_pairs)
        print(f"        → {len(result['出来形管理'])} 行（準用展開後）")

    print("[ 2/3 ] 品質管理基準を抽出中...")
    result["品質管理"] = extract_hinshitsu(施工管理基準_path)
    print(f"        → {len(result['品質管理'])} 行取得")

    print("[ 3/3 ] 撮影箇所一覧表を抽出中...")
    result["撮影箇所"] = extract_photo(写真管理基準_path)
    print(f"        → {len(result['撮影箇所'])} 行取得")

    return result


# ---------------------------------------------------------------------------
# 数量総括表 抽出
# ---------------------------------------------------------------------------

# 数量総括表の列数（ページによって異なる）
# 22/23列: 国交省標準フォーマット（工種はcol[1]）
# 13列: 一部地方整備局フォーマット（工種はcol[0]）
SURYO_VALID_COL_COUNTS = {13, 22, 23}
SURYO_KOJYO_COL_IDX    = 1   # 22/23列フォーマットの工種列（デフォルト）
_SURYO_KOJYO_COL_MAP   = {13: 0, 22: 1, 23: 1}  # 列数→工種列インデックス
_SURYO_COL_CONFIG      = {
    23: {"規格": 5, "単位": 9,  "数量": 14},
    22: {"規格": 5, "単位": 8,  "数量": 13},
    13: {"規格": 4, "単位": 5,  "数量": 8},
}

# コスト集計項目（工種ではない）- 国交省基準との照合対象外
SURYO_COST_WORDS = {
    "直接工事費", "共通仮設費", "共通仮設費（率計上）", "純工事費",
    "現場管理費", "工事原価", "一般管理費等", "工事価格",
    "消費税相当額", "工事費計", "運搬費", "重建設機械分解組立輸送費",
    "技術管理費", "現場環境改善費（率計上）",
    "ｼｽﾃﾑ初期費(ICT)", "システム初期費(ICT)",
}

# 工種エリアのx座標範囲・階層列名
_SURYO_X_MIN   = 35
_SURYO_X_MAX   = 230
SURYO_LEVEL_COLS = ["工種", "種別", "細別", "名称"]  # 浅い→深い順


def _build_suryo_x0_map(page) -> dict:
    """
    ページ内の工種エリア（x座標範囲）のテキスト→x0マッピングを返す。
    同一テキストが複数出現する場合は最初の出現のx0を使用する。
    """
    words = page.extract_words(x_tolerance=3, y_tolerance=3)
    area  = [w for w in words
             if _SURYO_X_MIN <= w["x0"] < _SURYO_X_MAX and w["text"].strip()]

    # y座標でグループ化（5pt以内を同一行とみなす）
    rows: dict = {}
    for w in area:
        placed = False
        for ky in rows:
            if abs(ky - w["top"]) < 5:
                rows[ky].append(w)
                placed = True
                break
        if not placed:
            rows[w["top"]] = [w]

    result: dict = {}
    for ws in rows.values():
        ws = sorted(ws, key=lambda w: w["x0"])
        text = _clean(" ".join(w["text"] for w in ws))
        x0   = ws[0]["x0"]
        if text and text not in result:
            result[text] = x0
    return result


def extract_suryo(pdf_path: str) -> dict:
    """
    数量総括表から工事名と工種の4階層（工種/種別/細別/名称）を抽出する。

    x座標からインデントレベルを判定し、各行を浅い→深い順に分類する。

    Returns:
        {
            "工事名":   str,
            "工種リスト": list[str],   # 後方互換用フラットリスト（全レベル）
            "工種階層":  pd.DataFrame, # 列: 工種/種別/細別/名称
        }
    """
    工事名 = ""
    all_records: list = []
    excluded_records: list = []
    _all_extracted_text: list[str] = []  # 日本語文字化け検出用

    with pdfplumber.open(pdf_path) as pdf:
        # ── 工事名抽出 ──────────────────────────────────────
        # 全ページのテキストから探す（表紙なしPDFに対応）
        for page in pdf.pages[:4]:
            page_text = page.extract_text() or ""
            _all_extracted_text.append(page_text)
            for line in page_text.split("\n"):
                line    = line.strip()
                compact = re.sub(r"\s", "", line)
                # スペースを詰めた形で「令和」「工事」を含み「工事名」ラベルではない行
                if "令和" in compact and "工事" in compact and "工事名" not in compact:
                    工事名 = re.sub(r"\s+", " ", line).strip()
                    break
            if 工事名:
                break

        # テキストから取れなければテーブルヘッダー行から取得
        if not 工事名:
            for page in pdf.pages[:4]:
                for table in page.extract_tables():
                    if not table or len(table[0]) not in SURYO_VALID_COL_COUNTS:
                        continue
                    for row in table[:2]:
                        for cell in row:
                            v = _clean(cell)
                            if v and "令和" in v and "工事" in v:
                                工事名 = v
                                break
                        if 工事名:
                            break
                if 工事名:
                    break

        # ── 全ページから工種を階層付きで収集 ─────────────────
        # page[0] から処理（表紙なしPDFに対応）
        for page in pdf.pages:
            tables = page.extract_tables()
            valid  = [t for t in tables if t and len(t[0]) in SURYO_VALID_COL_COUNTS]
            if not valid:
                continue

            n_cols      = len(valid[0][0])
            kojyo_col   = _SURYO_KOJYO_COL_MAP.get(n_cols, SURYO_KOJYO_COL_IDX)
            text_to_x0  = _build_suryo_x0_map(page)

            for row in valid[0][2:]:   # 先頭2行はヘッダー
                if len(row) <= kojyo_col:
                    continue
                val = _clean(row[kojyo_col])
                if not val:
                    continue
                if val.startswith("(") or val.startswith("（"):
                    excluded_records.append({"項目名": val, "除外理由": "小計・合計行（括弧始まり）"})
                    continue
                if val in SURYO_COST_WORDS:
                    excluded_records.append({"項目名": val, "除外理由": "費用集計項目"})
                    continue
                if "工事区分" in val or "工事名" in val:
                    excluded_records.append({"項目名": val, "除外理由": "ヘッダー行"})
                    continue

                all_records.append({
                    "_val": val,
                    "_x0":  text_to_x0.get(val),
                })

    # ── x0 → 階層レベル変換 ──────────────────────────────
    x0_vals = [r["_x0"] for r in all_records if r["_x0"] is not None]
    if x0_vals:
        rounded_vals  = sorted(set(round(x / 5) * 5 for x in x0_vals))
        x0_to_level   = {r: i for i, r in enumerate(rounded_vals)}
    else:
        x0_to_level   = {}

    n_lv     = len(SURYO_LEVEL_COLS)
    current  = [""] * n_lv
    prev_lv  = 0
    rows     = []

    for r in all_records:
        x0 = r["_x0"]
        if x0 is not None:
            level = x0_to_level.get(round(x0 / 5) * 5, 0)
        else:
            level = prev_lv
        level = max(0, min(level, n_lv - 1))

        current[level] = r["_val"]
        for i in range(level + 1, n_lv):
            current[i] = ""
        prev_lv = level
        rows.append(dict(zip(SURYO_LEVEL_COLS, current)))

    df_hierarchy = (
        pd.DataFrame(rows, columns=SURYO_LEVEL_COLS)
        if rows else pd.DataFrame(columns=SURYO_LEVEL_COLS)
    )

    # 後方互換用フラットリスト（全レベルの値を結合）
    kojyo_set: set = set()
    for col in SURYO_LEVEL_COLS:
        kojyo_set.update(v for v in df_hierarchy[col].unique() if v)

    df_excluded = pd.DataFrame(excluded_records, columns=["項目名", "除外理由"]) \
        if excluded_records else pd.DataFrame(columns=["項目名", "除外理由"])

    # ── 日本語文字化け検出 ────────────────────────────────────
    # テキストが抽出されているのに CJK 文字が1文字もない場合はフォントエンコーディング問題
    def _has_cjk(text: str) -> bool:
        import unicodedata as _ud
        return any(
            'CJK' in _ud.name(c, '') or 'HIRAGANA' in _ud.name(c, '') or 'KATAKANA' in _ud.name(c, '')
            for c in text
        )

    combined_text = " ".join(_all_extracted_text)
    pdf_text_readable = (not combined_text.strip()) or _has_cjk(combined_text)

    return {
        "工事名":          工事名,
        "工種リスト":       sorted(kojyo_set),
        "工種階層":         df_hierarchy,
        "除外行":           df_excluded,
        "pdf_text_readable": pdf_text_readable,  # Falseならフォントエンコーディング問題
    }


# ---------------------------------------------------------------------------
# 工種マッチング
# ---------------------------------------------------------------------------

def get_unique_kojyo(data: dict) -> dict:
    """
    国交省基準データから各シートのユニークな工種リストを返す。
    Streamlit UIの multiselect 用。
    """
    return {
        "出来形管理": sorted(data["出来形管理"]["工種"].unique().tolist()),
        "品質管理":   sorted(data["品質管理"]["工種"].unique().tolist()),
    }


def _normalize(s: str) -> str:
    """マッチング用の正規化: NFKC変換（半角カナ→全角等）+ スペース・括弧・記号を除去して小文字化。"""
    if not s:
        return ""
    import unicodedata as _ud
    s = _ud.normalize("NFKC", str(s))                  # 半角カナ→全角、全角英数→半角など
    s = re.sub(r"\s+", "", s)                          # 全空白除去
    s = re.sub(r"[（）()【】\[\]「」『』]", "", s)   # 括弧除去
    s = re.sub(r"[・･]", "", s)                        # 中点除去
    return s.lower()


def _is_hinshitsu_excluded(kojyo_name: str, extra_excludes: set = None) -> bool:
    """品質管理マッチングから除外すべき工種かどうか判定する。

    - 施工後試験を含む工種は無条件除外
    - 「を除く」の文脈（例: "セメント・コンクリート(転圧コンクリート…を除く)(施工)"）は保持
    - _HINSHITSU_EXCLUDE_ALWAYS のキーワードを含む工種は除外
    - extra_excludes（数量総括表に明示がないキーワード）も追加で除外

    ※ 正規化して比較（全角/半角の違いを吸収）
    """
    nn = _normalize(kojyo_name)
    if _normalize("施工後試験") in nn:
        return True
    if _normalize("を除く") in nn:
        return False
    excludes = list(get_hinshitsu_exclude_always())
    if extra_excludes:
        excludes += list(extra_excludes)
    return any(_normalize(kw) in nn for kw in excludes)


def build_suryo_match_map(suryo_keywords: list, kojyo_list: list) -> dict:
    """
    国交省基準工種 → 数量総括表工種 の対応辞書を返す。
    suggest_matches と同じマッチングロジック。出来形一覧の工種大分類に使用。

    Returns:
        {国交省工種名: 数量総括表工種名}
    """
    norm_kws = {kw: _normalize(kw) for kw in suryo_keywords
                if len(kw.strip()) >= 2 and len(_normalize(kw)) >= 2}

    result = {}
    for kojyo in kojyo_list:
        norm_kojyo = _normalize(kojyo)
        for suryo_kw, norm_kw in norm_kws.items():
            # suryo ⊂ DB を優先、フォールバックで DB ⊂ suryo
            if norm_kw in norm_kojyo or norm_kojyo in norm_kw:
                result[kojyo] = suryo_kw
                break
    return result


# 出来形管理DBの階層列（最深→最浅の順）
DEKIGATA_HIERARCHY_COLS = ["工種", "条", "節", "章", "編"]


def suggest_matches_hierarchical(suryo_keywords: list, df: pd.DataFrame) -> list:
    """
    出来形管理DBを階層的にマッチングする。

    アルゴリズム:
      1. 数量総括表キーワードを最深の「工種」列に対してマッチング試行。
      2. マッチしなかったキーワードは次の「条」「節」「章」「編」へと遡る。
      3. 各階層で一致した場合、対応する「工種」値をすべて候補として収集する。
         （複数候補が存在する場合はすべてを返し、UIでユーザーが確定する）

    Args:
        suryo_keywords: 数量総括表から抽出した工種キーワードリスト
        df:             国交省基準の出来形管理DataFrame（DEKIGATA_HIERARCHY_COLS列を含む）

    Returns:
        マッチした国交省基準工種のリスト（ソート済み）
    """
    norm_kws = [_normalize(kw) for kw in suryo_keywords if len(kw.strip()) >= 2]
    norm_kws = [n for n in norm_kws if len(n) >= 2]

    if not norm_kws or df.empty:
        return []

    suggested: set = set()
    unmatched: set = set(norm_kws)

    for col in DEKIGATA_HIERARCHY_COLS:
        if not unmatched or col not in df.columns:
            break

        # 正規化済みの列値 → 工種リスト のマッピングを構築
        val_to_kojyo: dict = {}
        for norm_v, kojyo in zip(
            df[col].apply(lambda x: _normalize(str(x or ""))),
            df["工種"],
        ):
            if norm_v:
                val_to_kojyo.setdefault(norm_v, [])
                if kojyo not in val_to_kojyo[norm_v]:
                    val_to_kojyo[norm_v].append(kojyo)

        still_unmatched: set = set()
        for norm_kw in unmatched:
            hit = False
            for col_val, kojyo_list in val_to_kojyo.items():
                if norm_kw in col_val or col_val in norm_kw:
                    suggested.update(kojyo_list)
                    hit = True
            if not hit:
                still_unmatched.add(norm_kw)

        unmatched = still_unmatched

    return sorted(suggested)


def suggest_from_suryo_hierarchy(
    suryo_df: pd.DataFrame,
    db_kojyo_list: list,
) -> list:
    """
    数量総括表の4階層（工種/種別/細別/名称）を使い、
    名称→細別→種別→工種の順（深い→浅い）でDBとマッチングする。

    マッチング戦略:
      各ユニーク行に対して、最も深いレベルから順に照合。
      いずれかのレベルでマッチしたらそこで打ち止め（上位レベルには遡らない）。
      複数マッチは全て採用してユーザーが選択する。

    Args:
        suryo_df:      extract_suryo()["工種階層"]
        db_kojyo_list: 国交省DBのユニーク工種リスト

    Returns:
        マッチした国交省DB工種のリスト（ソート済み）
    """
    if suryo_df.empty or not db_kojyo_list:
        return []

    # 正規化済みDB工種リスト（2文字以上のみ）
    norm_db = [(k, _normalize(k)) for k in db_kojyo_list if len(_normalize(k)) >= 2]

    suggested: set = set()

    # ユニークな階層チェーンを処理（_match_chain に委譲）
    chains = suryo_df[SURYO_LEVEL_COLS].drop_duplicates()
    for _, chain in chains.iterrows():
        matched, _ = _match_chain(chain.to_dict(), norm_db)
        suggested.update(matched)

    return sorted(suggested)


# 正規化済みキャッシュ（初回呼び出し時に構築）
_norm_alias_cache = None
_norm_narrowing_cache = None


def _get_norm_alias():
    """kojyo_alias.ALIAS_B_TO_A を正規化済みキーで再構築したキャッシュを返す。"""
    global _norm_alias_cache
    from kojyo_alias import ALIAS_B_TO_A
    _norm_alias_cache = {
        _normalize(k): {_normalize(v) for v in vals}
        for k, vals in ALIAS_B_TO_A.items()
    }
    return _norm_alias_cache


def _get_norm_narrowing():
    """match_filter.NARROWING_TABLE を正規化済みキーで再構築したキャッシュを返す。"""
    global _norm_narrowing_cache
    from match_filter import NARROWING_TABLE
    _norm_narrowing_cache = {
        _normalize(k): {_normalize(v) for v in vals}
        for k, vals in NARROWING_TABLE.items()
    }
    return _norm_narrowing_cache


def _narrow_with_shallower(matches, shallower_norms):
    """
    浅い階層キーワード（種別・工種）でマッチ候補を絞り込む。

    各キーワードをプレフィックス短縮しながら候補DB名に含まれるか確認し、
    件数が減る場合のみ絞り込みを適用する。
    """
    current = [(k, _normalize(k)) for k in matches]
    for sn in shallower_norms:
        for end in range(len(sn), 1, -1):
            prefix = sn[:end]
            filtered = [(k, nk) for k, nk in current if prefix in nk]
            if filtered and len(filtered) < len(current):
                current = filtered
                break
    result = [k for k, _ in current]
    return result if len(result) < len(matches) else None


# 絞り込み後もこの件数を超えたら「汎用的すぎるキーワード」と判断して次のレベルへ
_MAX_AFTER_NARROW = 8


def _match_chain(chain: dict, norm_db: list, *, norm_alias_override=None) -> tuple:
    """
    1チェーン（工種/種別/細別/名称）に対して名称→細別→種別→工種の順でマッチングする。

    マッチング戦略:
      1. 最も深いレベルからマッチングを試み、最初に「有効な結果」が得られたレベルで確定
         (a) suryo⊂DB を優先、なければ DB⊂suryo をフォールバック
         (b) それでも 0件なら kojyo_alias で補完（語順違い・別名等）
      2. 複数候補が残った場合、浅い階層のキーワードでアルゴリズム絞り込み
      3. さらに残った場合、NARROWING_TABLE で絞り込み（部分一致でテーブルキーを照合）
      4. 絞り込み後も _MAX_AFTER_NARROW を超える場合は汎用キーワードと判断して次のレベルへ
         ※ 盲目的な閾値スキップではなく「全絞り込みを試みた後の最終判断」として遡行する

    Args:
        norm_alias_override: 指定時、kojyo_alias の代わりに使用する別名辞書
                             ({normalized_source: {normalized_target, ...}})

    Returns:
        (matched_kojyo_list, matched_level_name)
    """
    ctx: dict[str, str] = {}
    for lvl in SURYO_LEVEL_COLS:
        v = chain.get(lvl, "")
        if v and len(_normalize(v)) >= 2:
            ctx[lvl] = _normalize(v)

    norm_alias     = norm_alias_override if norm_alias_override is not None else _get_norm_alias()
    norm_narrowing = _get_norm_narrowing()
    ctx_vals       = list(ctx.values())

    for level in reversed(SURYO_LEVEL_COLS):
        if level not in ctx:
            continue
        norm_val = ctx[level]

        # Step 1a: アルゴリズムマッチング（suryo⊂DB 優先、DB⊂suryo フォールバック）
        matches = [k for k, nk in norm_db if norm_val in nk]
        if not matches:
            matches = [k for k, nk in norm_db if nk in norm_val]

        # Step 1b: kojyo_alias 補完（語順違い・別名で 0件の場合）
        if not matches:
            alias_norms = norm_alias.get(norm_val, set())
            matches = [k for k, nk in norm_db if nk in alias_norms]

        if not matches:
            continue

        # Step 2a: 浅い階層キーワードでアルゴリズム絞り込み
        lvl_idx = SURYO_LEVEL_COLS.index(level)
        shallower_norms = [v for l, v in ctx.items()
                           if SURYO_LEVEL_COLS.index(l) < lvl_idx]
        if shallower_norms and len(matches) > 1:
            narrowed = _narrow_with_shallower(matches, shallower_norms)
            if narrowed:
                matches = narrowed

        # Step 2b: NARROWING_TABLE で追加絞り込み
        # テーブルキーが ctx のいずれかの値の部分文字列であれば適用
        if len(matches) > 1:
            for kw_norm, allowed_norms in norm_narrowing.items():
                if any(kw_norm in cv for cv in ctx_vals):
                    filtered = [k for k in matches if _normalize(k) in allowed_norms]
                    if filtered:
                        matches = filtered
                        break

        # Step 3: 全絞り込み後も多すぎる場合は次のレベルへ遡る
        if len(matches) > _MAX_AFTER_NARROW:
            continue

        return matches, level

    return [], ""


def _expand_to_rows(matched_kojyo: list, full_df: pd.DataFrame, detail_cols: list) -> str:
    """
    マッチした工種名リストを DB 行レベルのラベルに展開する。

    - 工種に1行のみ存在 → "工種名"（変換済み扱い）
    - 工種に複数行存在 → "工種名 / 列1 / 列2" を行ごとに "\\n" で結合（要確認扱い）
    - 複数工種がマッチ  → それぞれを展開して "\\n" で結合（要確認扱い）
    """
    if not matched_kojyo:
        return ""

    labels = []
    for kojyo in matched_kojyo:
        rows = full_df[full_df["工種"] == kojyo]
        if len(rows) <= 1:
            labels.append(kojyo)
        else:
            for _, r in rows.iterrows():
                parts = [kojyo]
                for col in detail_cols:
                    v = str(r.get(col, "") or "").strip()
                    if v:
                        parts.append(v)
                label = " / ".join(parts)
                if label not in labels:
                    labels.append(label)

    return "\n".join(labels)


# 種別のソート順（材料 → 施工 → その他）
_SHIKEN_KUBUN_ORDER = {"材料": 0, "施工": 1, "施工前試験": 2, "施工後試験": 3}


def _expand_hinshitsu_rows(matched_kojyo: list, full_df: pd.DataFrame) -> str:
    """
    品質管理マッチ専用の行展開。

    - マッチした DB の工種名を保持してラベルを生成する
    - 工種ごとに種別（材料→施工→施工前試験→施工後試験→その他）の順にソートして並べる
    - 工種に1行のみ存在 → "工種名"
    - 工種に複数行存在 → "工種名 / 種別 / 試験項目" を行ごとに改行区切り
    """
    if not matched_kojyo:
        return ""

    labels = []
    for kojyo in matched_kojyo:
        rows = full_df[full_df["工種"] == kojyo].copy()
        if rows.empty:
            labels.append(kojyo)
            continue

        # 施工後試験の行を除外（品管で不要なケースが多い）
        if "種別" in rows.columns:
            rows = rows[~rows["種別"].str.contains("施工後試験", na=False)]
            if rows.empty:
                continue

        # 種別でソート
        if "種別" in rows.columns:
            rows["_sort_key"] = rows["種別"].apply(
                lambda x: _SHIKEN_KUBUN_ORDER.get(str(x).strip(), 99)
            )
            rows = rows.sort_values("_sort_key")

        if len(rows) == 1:
            labels.append(kojyo)
        else:
            for _, r in rows.iterrows():
                parts = [kojyo]
                for col in ["種別", "試験項目"]:
                    v = str(r.get(col, "") or "").strip()
                    if v:
                        parts.append(v)
                label = " / ".join(parts)
                if label not in labels:
                    labels.append(label)

    return "\n".join(labels)


def _expand_photo_rows(
    matched_photo_d: list,
    matched_photo_h: list,
    photo_df: pd.DataFrame,
    implicit_photo_kojyo: list = None,
    exclude_fn=None,
    exclude_fn_h=None,
) -> str:
    """
    撮影箇所DBの直接マッチ結果からラベルを生成する。

    _match_chain() で撮影箇所DBに直接マッチした工種名リストを受け取り、
    "工種名 / 撮影項目" 形式のラベルに展開する。

    Args:
        matched_photo_d: 撮影箇所DB出来形セクションにマッチした工種名リスト
        matched_photo_h: 撮影箇所DB品質セクションにマッチした工種名リスト
        photo_df: 撮影箇所DB全体の DataFrame
        implicit_photo_kojyo: 暗黙ルールで追加する品質セクション工種名リスト
        exclude_fn: 条件付き除外関数・出来形セクション用（DB工種名を受け取りTrueなら除外）
        exclude_fn_h: 条件付き除外関数・品質セクション用（DB工種名を受け取りTrueなら除外）
    """
    labels: list = []
    seen: set = set()

    def _add_row(r) -> None:
        kojyo_val = str(r.get("工種", "") or "").strip()
        item = str(r.get("撮影項目", "") or "").strip()
        label = f"{kojyo_val} / {item}" if item else kojyo_val
        if label not in seen:
            labels.append(label)
            seen.add(label)

    def _add_from_section(df_section: pd.DataFrame, matched_kojyo: list,
                          section_exclude_fn=None) -> None:
        if not matched_kojyo:
            return
        norm_keys = [_normalize(k) for k in matched_kojyo if k and len(_normalize(k)) >= 2]
        if not norm_keys:
            return
        for idx, r in df_section.iterrows():
            nv = _normalize(str(r.get("工種", "") or ""))
            if nv and any(nv in nk or nk in nv for nk in norm_keys):
                kojyo_str = str(r.get("工種", "") or "")
                # 条件付き除外: 部分一致で拾ったDB工種が除外対象ならスキップ
                if section_exclude_fn and section_exclude_fn(kojyo_str):
                    continue
                _add_row(r)

    _add_from_section(
        photo_df[photo_df["セクション"] == "出来形管理"],
        matched_photo_d,
        section_exclude_fn=exclude_fn,
    )

    # 品質セクション: 出来形除外 + 品質除外の両方を適用
    def _combined_exclude_h(name: str) -> bool:
        if exclude_fn and exclude_fn(name):
            return True
        if exclude_fn_h and exclude_fn_h(name):
            return True
        return False

    _add_from_section(
        photo_df[photo_df["セクション"] == "品質管理"],
        matched_photo_h,
        section_exclude_fn=_combined_exclude_h,
    )

    # implicit_photo: 数量総括表キーワードから追加される品質セクション工種
    if implicit_photo_kojyo:
        _add_from_section(
            photo_df[photo_df["セクション"] == "品質管理"],
            implicit_photo_kojyo,
            section_exclude_fn=_combined_exclude_h,
        )

    return "\n".join(labels)


# ---------------------------------------------------------------------------
# 間接トリガーによる出来形管理・品質管理暗黙マッピング
# ルールは matching_rules.json から読み込む
# ---------------------------------------------------------------------------


def _get_implicit_dekigata(suryo_df: pd.DataFrame, db_dekigata_df: pd.DataFrame) -> list[str]:
    """
    数量総括表の全キーワードを走査し、間接トリガーで追加すべき出来形管理工種を返す。
    """
    rules = get_implicit_dekigata_rules()
    if not rules:
        return []
    db_dekigata_set = set(db_dekigata_df["工種"].unique())

    all_suryo_norm: set[str] = set()
    for col in SURYO_LEVEL_COLS:
        for v in suryo_df[col].unique():
            n = _normalize(str(v or ""))
            if len(n) >= 2:
                all_suryo_norm.add(n)

    result: list[str] = []
    for trigger, targets in rules.items():
        nt = _normalize(trigger)
        if len(nt) < 2:
            continue
        hit = any(nt in sv or sv in nt for sv in all_suryo_norm)
        if hit:
            for tgt in targets:
                if tgt in db_dekigata_set and tgt not in result:
                    result.append(tgt)
    return result


def get_implicit_hinshitsu_labels(suryo_df: pd.DataFrame, db_hinshitsu_df: pd.DataFrame) -> list[str]:
    """
    数量総括表から間接トリガーで追加すべき品質管理ラベル（"工種 / 種別 / 試験項目" 形式）を返す。
    行選択状態に関わらず常に品管一覧に含めるために使用する。
    """
    kojyo_list = _get_implicit_hinshitsu(suryo_df, db_hinshitsu_df)
    return _expand_hinshitsu_rows(kojyo_list, db_hinshitsu_df).split("\n") if kojyo_list else []


def _get_implicit_hinshitsu(suryo_df: pd.DataFrame, db_hinshitsu_df: pd.DataFrame) -> list[str]:
    """
    数量総括表の全キーワードを走査し、間接トリガーで追加すべき品質管理工種を返す。

    例: 数量総括表に「カルバート工」があれば「セメント・コンクリート」と
    「プレキャストコンクリート製品」を追加する。

    Returns:
        追加すべき品質管理工種名のリスト（DB内に存在するもののみ）
    """
    rules = get_implicit_hinshitsu_rules()
    if not rules:
        return []
    db_hinshitsu_set = set(db_hinshitsu_df["工種"].unique())

    # 数量総括表の全ユニーク値を正規化してセットに集める
    all_suryo_norm: set[str] = set()
    for col in SURYO_LEVEL_COLS:
        for v in suryo_df[col].unique():
            n = _normalize(str(v or ""))
            if len(n) >= 2:
                all_suryo_norm.add(n)

    result: list[str] = []
    for trigger, targets in rules.items():
        nt = _normalize(trigger)
        if len(nt) < 2:
            continue
        hit = any(nt in sv or sv in nt for sv in all_suryo_norm)
        if hit:
            for tgt in targets:
                if tgt in db_hinshitsu_set and tgt not in result:
                    result.append(tgt)
    return result


def build_match_detail(
    suryo_df: pd.DataFrame,
    db_dekigata_df: pd.DataFrame,
    db_hinshitsu_df: pd.DataFrame,
    db_photo_df: pd.DataFrame = None,
) -> pd.DataFrame:
    """
    数量総括表の各ユニーク行に対し、出来形管理・品質管理のマッチング結果を返す。

    マッチングロジック: 名称→細別→種別→工種の順（深い→浅い）で照合し、
    最初にマッチしたレベルで確定。マッチした工種に複数のDB行がある場合は
    行レベルのラベル（工種 / 種別 / 試験項目 など）に展開して改行区切りで返す。
    → "\n" が含まれる場合は app 側で「要確認」として扱う。

    Args:
        suryo_df:        extract_suryo()["工種階層"]
        db_dekigata_df:  国交省基準の出来形管理 DataFrame（全行）
        db_hinshitsu_df: 国交省基準の品質管理 DataFrame（全行）

    Returns:
        列: 工種, 種別, 細別, 名称, 出来形マッチ, 品質管理マッチ
    """
    out_cols = SURYO_LEVEL_COLS + ["出来形マッチ", "品質管理マッチ", "撮影箇所マッチ"]
    if suryo_df.empty:
        return pd.DataFrame(columns=out_cols)

    # ── 数量総括表キーワード収集（条件付き除外の基礎データ） ──
    _suryo_norm_all: set[str] = set()
    for col in SURYO_LEVEL_COLS:
        for v in suryo_df[col].dropna().unique():
            _suryo_norm_all.add(_normalize(str(v)))

    # ── 出来形: 条件付き除外（数量総括表に言及がない舗装・防護柵等を除外） ──
    _dekigata_excludes: set[str] = set()
    for kw in get_dekigata_exclude_unless_in_suryo():
        nkw = _normalize(kw)
        if not any(nkw in sv for sv in _suryo_norm_all):
            _dekigata_excludes.add(nkw)

    def _is_dekigata_excluded(name: str) -> bool:
        nn = _normalize(name)
        return any(ex in nn for ex in _dekigata_excludes)

    norm_d = [
        (k, _normalize(k))
        for k in db_dekigata_df["工種"].unique()
        if len(_normalize(k)) >= 2
        and not _is_dekigata_excluded(k)
    ]

    # ── 品管: 条件付き除外 ──
    hin_hissu_kojyo = set(
        db_hinshitsu_df[db_hinshitsu_df["試験区分"] == "必須"]["工種"].unique()
    )
    _conditional_excludes: set[str] = set()
    for kw in get_hinshitsu_exclude_unless_in_suryo():
        nkw = _normalize(kw)
        if not any(nkw in sv for sv in _suryo_norm_all):
            _conditional_excludes.add(kw)
    norm_h = [
        (k, _normalize(k))
        for k in hin_hissu_kojyo
        if len(_normalize(k)) >= 2
        and not _is_hinshitsu_excluded(k, _conditional_excludes)
    ]

    # ── 撮影箇所: 出来形/品質セクション別に norm リストを構築 ──
    # 撮影箇所では dekigata_exclude_unless_in_suryo を無条件適用する
    # （出来形管理では数量総括表にあれば除外しないが、撮影箇所では常に除外）
    _photo_excludes: set[str] = set(
        _normalize(kw) for kw in get_dekigata_exclude_unless_in_suryo()
    )

    def _is_photo_excluded(name: str) -> bool:
        nn = _normalize(name)
        return any(ex in nn for ex in _photo_excludes)

    norm_photo_d: list = []
    norm_photo_h: list = []
    photo_alias_d: dict = {}
    photo_alias_h: dict = {}
    implicit_photo_kojyo: list = []
    if db_photo_df is not None:
        from photo_alias import get_norm_alias_for_match_chain, get_implicit_photo_from_suryo
        df_photo_d = db_photo_df[db_photo_df["セクション"] == "出来形管理"]
        df_photo_h = db_photo_df[db_photo_df["セクション"] == "品質管理"]
        norm_photo_d = [
            (k, _normalize(k)) for k in df_photo_d["工種"].unique()
            if len(_normalize(k)) >= 2 and not _is_photo_excluded(k)
        ]
        norm_photo_h = [
            (k, _normalize(k)) for k in df_photo_h["工種"].unique()
            if len(_normalize(k)) >= 2
            and not _is_hinshitsu_excluded(k, _conditional_excludes)
            and not _is_photo_excluded(k)
        ]
        photo_alias_d = get_norm_alias_for_match_chain("出来形管理")
        photo_alias_h = get_norm_alias_for_match_chain("品質管理")
        implicit_photo_kojyo = [
            k for k in get_implicit_photo_from_suryo(suryo_df, db_photo_df)
            if not _is_hinshitsu_excluded(k, _conditional_excludes)
            and not _is_photo_excluded(k)
        ]

    # 間接トリガーで追加すべき工種（プロジェクト全体レベル）
    implicit_h = _get_implicit_hinshitsu(suryo_df, db_hinshitsu_df)

    # 出来形の間接トリガー: チェーン行ごとに該当ルールを適用するための準備
    _implicit_d_rules = get_implicit_dekigata_rules()
    _db_dekigata_set = set(db_dekigata_df["工種"].unique())

    def _get_chain_implicit_d(chain_dict: dict) -> list[str]:
        """チェーン行のキーワードにマッチする間接出来形工種を返す。"""
        if not _implicit_d_rules:
            return []
        chain_norms = set()
        for col in SURYO_LEVEL_COLS:
            n = _normalize(str(chain_dict.get(col, "") or ""))
            if len(n) >= 2:
                chain_norms.add(n)
        result = []
        for trigger, targets in _implicit_d_rules.items():
            nt = _normalize(trigger)
            if len(nt) < 2:
                continue
            if any(nt in cn or cn in nt for cn in chain_norms):
                for tgt in targets:
                    if tgt in _db_dekigata_set and tgt not in result:
                        result.append(tgt)
        return result

    records = []
    chains = suryo_df[SURYO_LEVEL_COLS].drop_duplicates()

    # 間接品管は最初の種別ルート行に1回だけ付与
    implicit_added = False
    implicit_photo_added = False
    # 出来形の間接トリガーは各チェーン行ごとに付与済みを追跡
    implicit_d_added_set: set = set()

    # ── 1st pass: 全チェーンのマッチングを実行し、候補を収集 ──
    chain_results = []
    for _, chain in chains.iterrows():
        cd = chain.to_dict()
        md, md_level = _match_chain(cd, norm_d)
        mh, mh_level = _match_chain(cd, norm_h)

        # 撮影箇所: 数量総括表から直接マッチ（品質・出来形と同じ方式）
        mp_d, _ = _match_chain(cd, norm_photo_d, norm_alias_override=photo_alias_d) if norm_photo_d else ([], "")
        mp_h, _ = _match_chain(cd, norm_photo_h, norm_alias_override=photo_alias_h) if norm_photo_h else ([], "")

        chain_results.append((cd, md, md_level, mh, mh_level, mp_d, mp_h))

    # ── 2nd pass: 同一種別グループ内で深い階層の具体的マッチがある場合、
    #    浅い階層の広い候補（候補数>1）を間引く ──
    for i, (cd, md, md_level, mh, mh_level, mp_d, mp_h) in enumerate(chain_results):
        if len(md) <= 1:
            continue
        suryo_shubetsu = cd.get("種別", "")
        if not suryo_shubetsu:
            continue
        # 同じ工種+種別のグループ内で、より具体的なマッチ（候補1件）があるか探す
        specific_d: set[str] = set()
        for j, (cd2, md2, md_level2, _, _, _, _) in enumerate(chain_results):
            if i == j:
                continue
            if cd2.get("種別", "") == suryo_shubetsu and cd2.get("工種", "") == cd.get("工種", ""):
                if len(md2) == 1:
                    specific_d.update(md2)
        # 具体的マッチがある場合、現在の候補をその具体的マッチに限定
        if specific_d:
            narrowed = [m for m in md if m in specific_d]
            if narrowed:
                chain_results[i] = (cd, narrowed, md_level, mh, mh_level, mp_d, mp_h)

    # ── 3rd pass: records を組み立て ──
    for cd, md, md_level, mh, mh_level, mp_d, mp_h in chain_results:
        # 間接出来形: チェーン行のキーワードに対応する暗黙工種を付与
        # 細別が空の行（工種ルート行 or 種別ルート行）に対して適用
        extra_d = []
        if not cd.get("細別", ""):
            chain_implicit_d = _get_chain_implicit_d(cd)
            extra_d = [k for k in chain_implicit_d
                       if k not in md and k not in implicit_d_added_set]
            implicit_d_added_set.update(extra_d)

        # 間接品管をまだ追加していない場合、細別が空のルート行に付与
        extra_h = []
        if implicit_h and not implicit_added and not cd.get("細別", ""):
            extra_h = [k for k in implicit_h if k not in mh]
            if extra_h:
                implicit_added = True

        # 撮影箇所の暗黙ルール（implicit_photo）を最初の細別が空のルート行に付与
        extra_photo = []
        if implicit_photo_kojyo and not implicit_photo_added and not cd.get("細別", ""):
            extra_photo = [k for k in implicit_photo_kojyo if k not in mp_h]
            if extra_photo:
                implicit_photo_added = True

        all_md = md + extra_d
        all_mh = mh + extra_h
        photo_match = _expand_photo_rows(
            mp_d, mp_h, db_photo_df, extra_photo,
            exclude_fn=_is_photo_excluded,
            exclude_fn_h=lambda name: _is_hinshitsu_excluded(name, _conditional_excludes),
        ) if db_photo_df is not None else ""
        records.append({
            "工種":          cd.get("工種", ""),
            "種別":          cd.get("種別", ""),
            "細別":          cd.get("細別", ""),
            "名称":          cd.get("名称", ""),
            "出来形マッチ":   _expand_to_rows(all_md, db_dekigata_df, ["測定項目"]),
            "品質管理マッチ": _expand_hinshitsu_rows(all_mh, db_hinshitsu_df),
            "撮影箇所マッチ": photo_match,
        })

    return pd.DataFrame(records, columns=out_cols)


def suggest_matches(suryo_keywords: list, kojyo_list: list) -> list:
    """
    数量総括表の工種キーワードから国交省基準の工種候補を提案する。
    双方向部分一致でマッチングを行う。

    Args:
        suryo_keywords: 数量総括表から抽出した工種名のリスト
        kojyo_list:     国交省基準のユニーク工種リスト

    Returns:
        自動マッチングされた国交省基準工種のリスト
    """
    # 2文字以上のキーワードだけを使う
    norm_kws = [_normalize(kw) for kw in suryo_keywords if len(kw.strip()) >= 2]
    norm_kws = [n for n in norm_kws if len(n) >= 2]

    suggested = []
    for kojyo in kojyo_list:
        norm_kojyo = _normalize(kojyo)
        for norm_kw in norm_kws:
            # suryo ⊂ DB を優先、フォールバックで DB ⊂ suryo
            if norm_kw in norm_kojyo or norm_kojyo in norm_kw:
                suggested.append(kojyo)
                break

    return suggested


def filter_by_row_labels(
    data: dict,
    dekigata_labels: list,
    hinshitsu_labels: list,
    photo_labels: list = None,
) -> dict:
    """
    行レベルのラベルで各シートをフィルタリングする。

    ラベル形式（_expand_to_rows が生成）:
      出来形管理: "工種 / 測定項目"  または "工種"（1行のみの場合）
      品質管理:  "工種 / 種別 / 試験項目"  または "工種"

    Args:
        data:             load_kojyo_db() の戻り値
        dekigata_labels:  ユーザーが選択した出来形管理ラベルのリスト
        hinshitsu_labels: ユーザーが選択した品質管理ラベルのリスト

    Returns:
        絞り込み済みの data と同じ構造の dict
    """
    filtered: dict = {}

    # ── 出来形管理: 行レベルフィルタ ──────────────────────────────
    df_d = data["出来形管理"]
    if dekigata_labels:
        masks = []
        for label in dekigata_labels:
            parts = [p.strip() for p in label.split(" / ")]
            m = df_d["工種"] == parts[0]
            if len(parts) >= 2 and parts[1]:
                m = m & (df_d["測定項目"] == parts[1])
            masks.append(m)
        combined = masks[0]
        for m in masks[1:]:
            combined = combined | m
        filtered["出来形管理"] = df_d[combined].reset_index(drop=True)
    else:
        filtered["出来形管理"] = df_d.iloc[0:0].copy()

    # ── 品質管理: 行レベルフィルタ ────────────────────────────────
    df_h = data["品質管理"]
    if hinshitsu_labels:
        masks = []
        for label in hinshitsu_labels:
            parts = [p.strip() for p in label.split(" / ")]
            m = df_h["工種"] == parts[0]
            if len(parts) >= 2 and parts[1]:
                m = m & (df_h["種別"] == parts[1])
            if len(parts) >= 3 and parts[2]:
                m = m & (df_h["試験項目"] == parts[2])
            masks.append(m)
        combined = masks[0]
        for m in masks[1:]:
            combined = combined | m
        filtered["品質管理"] = df_h[combined].reset_index(drop=True)
    else:
        filtered["品質管理"] = df_h.iloc[0:0].copy()

    # ── 撮影箇所 ─────────────────────────────────────────────────
    df_p = data["撮影箇所"]
    zentai = df_p[df_p["セクション"] == "全体"].copy()

    if photo_labels is not None:
        # photo_labels が明示的に渡された場合: ラベルで直接フィルタ
        df_non_zentai = df_p[df_p["セクション"] != "全体"].copy()
        if photo_labels:
            masks = []
            for label in photo_labels:
                parts = [p.strip() for p in label.split(" / ")]
                kojyo = parts[0]
                item  = parts[1] if len(parts) > 1 else None
                m = df_non_zentai["工種"] == kojyo
                if item:
                    m = m & (df_non_zentai["撮影項目"] == item)
                masks.append(m)
            combined = masks[0]
            for m in masks[1:]:
                combined = combined | m
            df_non_zentai = df_non_zentai[combined]
        else:
            df_non_zentai = df_non_zentai.iloc[0:0]
        filtered["撮影箇所"] = pd.concat([zentai, df_non_zentai], ignore_index=True)
    else:
        # 旧方式: 出来形・品質管理ラベルから自動導出
        selected_d = list({lbl.split(" / ")[0].strip() for lbl in dekigata_labels})
        selected_h = list({lbl.split(" / ")[0].strip() for lbl in hinshitsu_labels})

        df_dek_photo = df_p[df_p["セクション"] == "出来形管理"].copy()
        if selected_d:
            norm_d = [_normalize(x) for x in selected_d]

            def _match_d(val):
                if pd.isna(val):
                    return False
                nv = _normalize(str(val))
                return any((nv in nd or nd in nv) for nd in norm_d if nd)

            df_dek_photo = df_dek_photo[df_dek_photo["工種"].apply(_match_d)]
        else:
            df_dek_photo = df_dek_photo.iloc[0:0]

        df_hin_photo = df_p[df_p["セクション"] == "品質管理"].copy()
        if selected_h:
            selected_nums = set()
            for h in selected_h:
                parts = h.strip().split()
                if parts and parts[0].isdigit():
                    selected_nums.add(parts[0])
            norm_h = [_normalize(x) for x in selected_h]

            def _match_h(row):
                ban = str(row.get("番号", "") or "").strip()
                if ban and ban in selected_nums:
                    return True
                nv = _normalize(str(row.get("工種", "") or ""))
                return any((nv in nh or nh in nv) for nh in norm_h if nh)

            df_hin_photo = df_hin_photo[df_hin_photo.apply(_match_h, axis=1)]
        else:
            df_hin_photo = df_hin_photo.iloc[0:0]

        filtered["撮影箇所"] = pd.concat(
            [zentai, df_hin_photo, df_dek_photo], ignore_index=True
        )

    return filtered


def filter_by_kojyo(
    data: dict,
    selected_dekigata: list,
    selected_hinshitsu: list,
) -> dict:
    """
    ユーザーが選択した工種で各シートのデータを絞り込む。

    Args:
        data:               extract_from_pdf() の戻り値
        selected_dekigata:  出来形管理基準で使用する工種リスト（完全一致）
        selected_hinshitsu: 品質管理基準で使用する工種リスト（完全一致）

    Returns:
        絞り込み済みの data と同じ構造の dict
    """
    filtered: dict = {}

    # ── 出来形管理: 工種の完全一致 ────────────────────────────
    df_d = data["出来形管理"]
    if selected_dekigata:
        filtered["出来形管理"] = df_d[df_d["工種"].isin(selected_dekigata)].reset_index(drop=True)
    else:
        filtered["出来形管理"] = df_d.iloc[0:0].copy()

    # ── 品質管理: 工種の完全一致 ──────────────────────────────
    df_h = data["品質管理"]
    if selected_hinshitsu:
        filtered["品質管理"] = df_h[df_h["工種"].isin(selected_hinshitsu)].reset_index(drop=True)
    else:
        filtered["品質管理"] = df_h.iloc[0:0].copy()

    # ── 撮影箇所: 全体は全件 / 品質管理・出来形管理は工種でフィルタ ──
    df_p = data["撮影箇所"]

    # 全体セクションは常に全件
    zentai = df_p[df_p["セクション"] == "全体"].copy()

    # 出来形管理セクション: 工種列とキーワードの双方向部分一致
    df_dek_photo = df_p[df_p["セクション"] == "出来形管理"].copy()
    if selected_dekigata:
        norm_d = [_normalize(x) for x in selected_dekigata]

        def _match_d(val):
            if pd.isna(val):
                return False
            nv = _normalize(str(val))
            return any((nv in nd or nd in nv) for nd in norm_d if nd)

        df_dek_photo = df_dek_photo[df_dek_photo["工種"].apply(_match_d)]
    else:
        df_dek_photo = df_dek_photo.iloc[0:0]

    # 品質管理セクション: 番号一致 or 工種キーワード一致
    df_hin_photo = df_p[df_p["セクション"] == "品質管理"].copy()
    if selected_hinshitsu:
        # "19 固結工" → 番号="19" を抽出して 番号列と照合
        selected_nums = set()
        for h in selected_hinshitsu:
            parts = h.strip().split()
            if parts and parts[0].isdigit():
                selected_nums.add(parts[0])
        norm_h = [_normalize(x) for x in selected_hinshitsu]

        def _match_h(row):
            ban = str(row.get("番号", "") or "").strip()
            if ban and ban in selected_nums:
                return True
            nv = _normalize(str(row.get("工種", "") or ""))
            return any((nv in nh or nh in nv) for nh in norm_h if nh)

        df_hin_photo = df_hin_photo[df_hin_photo.apply(_match_h, axis=1)]
    else:
        df_hin_photo = df_hin_photo.iloc[0:0]

    filtered["撮影箇所"] = pd.concat(
        [zentai, df_hin_photo, df_dek_photo], ignore_index=True
    )

    return filtered


# ---------------------------------------------------------------------------
# CLI 動作確認
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    施工管理基準 = (
        sys.argv[1]
        if len(sys.argv) > 1
        else "/Users/tomoki/Downloads/OneDrive_1_2026-5-18/土木工事施工管理基準及び規格値（案）.pdf"
    )
    写真管理基準 = (
        sys.argv[2]
        if len(sys.argv) > 2
        else "/Users/tomoki/Downloads/OneDrive_1_2026-5-18/写真管理基準.pdf"
    )

    data = extract_from_pdf(施工管理基準, 写真管理基準)

    for sheet_name, df in data.items():
        print(f"\n{'=' * 70}")
        print(f"【{sheet_name}】  {len(df)} 行 × {len(df.columns)} 列")
        print(f"列: {list(df.columns)}")
        print(df.head(10).to_string(max_colwidth=40))
