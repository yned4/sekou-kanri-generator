"""
excel_writer.py
抽出済みDataFrameを施工管理計画Excelファイルに書き出す。

使い方（単体テスト）:
    python excel_writer.py  # extractor.py と組み合わせて動作確認
"""

import io
import re
from typing import Optional
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ===== シート名定義 =====
SHEET_HINSHITSU = "品質管理基準及び規格値"
SHEET_DEKIGATA  = "出来形管理基準及び規格値一覧"
SHEET_PHOTO     = "撮影箇所一覧表"

# ===== スタイル定数 =====
HEADER_BG_COLOR  = "1F4E79"   # 濃い青
HEADER_FONT_COLOR = "FFFFFF"  # 白文字
SECTION_BG_COLOR = "D6E4F0"   # 薄い青（セクション区切り行）
MAX_COL_WIDTH    = 50          # 列幅の上限（文字数相当）
MIN_COL_WIDTH    = 8           # 列幅の下限


def _make_header_style() -> tuple:
    """ヘッダー行用スタイルを返す。"""
    font  = Font(bold=True, color=HEADER_FONT_COLOR, size=10)
    fill  = PatternFill("solid", fgColor=HEADER_BG_COLOR)
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    return font, fill, align


def _make_data_align(wrap: bool = True) -> Alignment:
    return Alignment(vertical="top", wrap_text=wrap)


def _make_thin_border() -> Border:
    thin = Side(style="thin", color="AAAAAA")
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def _write_sheet(ws, df: pd.DataFrame, section_col: Optional[str] = None) -> None:
    """
    DataFrameを1シートに書き込む。
    section_col: この列の値が変わるタイミングでセクション区切り行の背景色を変える。
    """
    header_font, header_fill, header_align = _make_header_style()
    data_align  = _make_data_align()
    thin_border = _make_thin_border()

    # ── ヘッダー行 ──────────────────────────────────────────
    for col_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = header_align
        cell.border    = thin_border

    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"

    # ── データ行 ──────────────────────────────────────────
    section_fill = PatternFill("solid", fgColor=SECTION_BG_COLOR)
    prev_section = None

    for row_idx, (_, row) in enumerate(df.iterrows(), start=2):
        # セクション区切りで薄い背景色を交互に付ける
        current_section = row.get(section_col) if section_col else None
        use_section_fill = (
            section_col is not None
            and current_section != prev_section
            and current_section is not None
        )

        for col_idx, value in enumerate(row, start=1):
            # NaN / None は空文字で書き込む
            if pd.isna(value) if not isinstance(value, str) else False:
                value = ""
            cell = ws.cell(row=row_idx, column=col_idx, value=str(value) if value != "" else "")
            cell.alignment = data_align
            cell.border    = thin_border
            if use_section_fill:
                cell.fill = section_fill

        if use_section_fill:
            prev_section = current_section

    # ── 列幅の自動調整 ──────────────────────────────────────
    for col_idx, col_name in enumerate(df.columns, start=1):
        col_letter = get_column_letter(col_idx)

        # ヘッダー文字数と各セルの最大文字数（改行前）から推定
        max_len = len(str(col_name))
        for value in df.iloc[:, col_idx - 1]:
            if pd.isna(value) if not isinstance(value, str) else False:
                continue
            # 複数行の場合は最も長い行を基準にする
            cell_max = max((len(line) for line in str(value).split("\n")), default=0)
            max_len = max(max_len, cell_max)

        width = min(max(max_len + 2, MIN_COL_WIDTH), MAX_COL_WIDTH)
        ws.column_dimensions[col_letter].width = width


def _write_title_row(ws, title: str, n_cols: int) -> None:
    """シート先頭に工事名タイトル行を書き込む（結合セル）。"""
    ws.insert_rows(1)
    cell = ws.cell(row=1, column=1, value=title)
    cell.font      = Font(bold=True, size=11)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    if n_cols > 1:
        ws.merge_cells(
            start_row=1, start_column=1,
            end_row=1,   end_column=min(n_cols, 10),
        )
    ws.row_dimensions[1].height = 20
    # フリーズ行を1行ずらす
    ws.freeze_panes = "A3"


def _compute_shauchi(規格値_val: str) -> str:
    """
    社内規格値を算出する。

    ・数値が含まれる場合: 各数値に0.8を乗じる（小数桁数は元に合わせる）。
    ・数値が含まれない場合: 規格値をそのまま転記。
    """
    s = str(規格値_val or "").strip()
    if not s or not re.search(r'\d', s):
        return s

    def _scale(m: re.Match) -> str:
        orig = m.group()
        val = float(orig)
        scaled = val * 0.8
        if '.' in orig:
            decimals = len(orig.split('.')[1])
            return f"{round(scaled, decimals):.{decimals}f}"
        else:
            # 元が整数でも結果が小数になる場合は小数1桁で表示
            if scaled == int(scaled):
                return str(int(scaled))
            return f"{scaled:.1f}"

    return re.sub(r'\d+(?:\.\d+)?', _scale, s)


def _reshape_hinshitsu(df: pd.DataFrame) -> pd.DataFrame:
    """
    品質管理DataFrameを出力列形式に変換する。
    削除: 試験区分、試験成績表等による確認
    追加: 社内規格値（空欄）
    リネーム: 試験時期・頻度→試験基準、摘要→備考
    """
    out = pd.DataFrame()
    out["工種"]       = df.get("工種", "")
    out["種別"]       = df.get("種別", "")
    out["試験項目"]   = df.get("試験項目", "")
    out["試験方法"]   = df.get("試験方法", "")
    out["規格値"]     = df.get("規格値", "")
    out["社内規格値"] = df["規格値"].apply(_compute_shauchi)
    out["試験基準"]   = df.get("試験時期・頻度", "")
    out["備考"]       = df.get("摘要", "")
    return out


def _reshape_dekigata(df: pd.DataFrame, kojyo_to_suryo: dict = None) -> pd.DataFrame:
    """
    出来形管理DataFrameを出力列形式に変換する。
    削除: 編、章、節、条、枝番、規格値_個々、測定基準
    追加: 工種（数量総括表由来の大分類）、社内規格値（空欄）
    リネーム: 工種→種別、摘要→備考

    kojyo_to_suryo: {国交省工種名: 数量総括表工種名} の対応辞書
    """
    out = pd.DataFrame()
    if kojyo_to_suryo:
        out["工種"] = df["工種"].map(lambda x: kojyo_to_suryo.get(str(x), "") if x else "")
    else:
        out["工種"] = ""
    out["種別"]        = df.get("工種", "")
    out["測定項目"]    = df.get("測定項目", "")
    out["規格値_条件"] = df.get("規格値_条件", "")
    out["規格値"]      = df.get("規格値", "")
    out["社内規格値"]  = df["規格値"].apply(_compute_shauchi)
    out["測定箇所"]    = df.get("測定基準", "")
    out["備考"]        = df.get("摘要", "")
    return out


_SECTION_TO_KUBUN = {
    "出来形管理": "出来形管理写真",
    "品質管理":   "品質管理写真",
}


def _reshape_photo(df: pd.DataFrame) -> pd.DataFrame:
    """
    撮影箇所DataFrameを出力列形式に変換する。

    区分:
      全体セクション  → 区分列の値をそのまま使用
      品質/出来形セクション → セクション名を「品質管理写真」「出来形管理写真」に変換

    工種:
      全体セクション  → sub区分列の値を使用（「着手前」「工事施工中」等）
      品質/出来形セクション → 工種列の値を使用
    """
    records = []

    for _, row in df.iterrows():
        sec = str(row.get("セクション") or "")

        # 全体セクションの除外条件
        if sec == "全体":
            # ① sub区分・撮影項目・撮影頻度がすべて空 → PDF上のカテゴリ見出し行
            if (not (row.get("sub区分") or "")
                    and not (row.get("撮影項目") or "")
                    and not (row.get("撮影頻度") or "")):
                continue
            # ② 区分が出来形管理・品質管理 → 専用セクションで詳細に出力済みのため重複除外
            raw_k = str(row.get("区分") or "")
            if raw_k in ("出来形管理", "品質管理"):
                continue

        # 区分
        raw_kubun = row.get("区分") or ""
        kubun = raw_kubun if raw_kubun else _SECTION_TO_KUBUN.get(sec, sec)

        # 工種
        raw_kojyo = row.get("工種") or ""
        kojyo = raw_kojyo if raw_kojyo else (row.get("sub区分") or "")

        records.append({
            "区分":     kubun,
            "工種":     kojyo,
            "撮影項目": row.get("撮影項目") or "",
            "撮影基準": row.get("撮影頻度") or "",
            "摘要":     row.get("摘要") or "",
        })

    return pd.DataFrame(records, columns=["区分", "工種", "撮影項目", "撮影基準", "摘要"])


def write_excel(
    data: dict,
    output_path: Optional[str] = None,
    工事名: str = "",
    dekigata_kojyo_map: dict = None,
) -> bytes:
    """
    抽出データをExcelに書き出す。

    Args:
        data:               filter_by_kojyo() または extract_from_pdf() の戻り値
        output_path:        保存先パス。None の場合はバイト列を返す（Streamlit用）。
        工事名:             数量総括表から取得した工事名。シート先頭行に表示する。
        dekigata_kojyo_map: {国交省工種名: 数量総括表工種名} の対応辞書。
                            出来形管理シートの「工種（大分類）」列に使用。

    Returns:
        output_path が None の場合は Excel バイト列、それ以外は None。
    """
    wb = Workbook()
    wb.remove(wb.active)  # デフォルトシートを削除

    df_h = _reshape_hinshitsu(data["品質管理"])
    df_d = _reshape_dekigata(data["出来形管理"], kojyo_to_suryo=dekigata_kojyo_map)
    df_p = _reshape_photo(data["撮影箇所"])

    # ── 品質管理 ──────────────────────────────────────────
    ws_h = wb.create_sheet(SHEET_HINSHITSU)
    _write_sheet(ws_h, df_h)
    if 工事名:
        _write_title_row(ws_h, 工事名, len(df_h.columns))

    # ── 出来形管理 ────────────────────────────────────────
    ws_d = wb.create_sheet(SHEET_DEKIGATA)
    _write_sheet(ws_d, df_d)
    if 工事名:
        _write_title_row(ws_d, 工事名, len(df_d.columns))

    # ── 撮影箇所 ──────────────────────────────────────────
    ws_p = wb.create_sheet(SHEET_PHOTO)
    _write_sheet(ws_p, df_p)
    if 工事名:
        _write_title_row(ws_p, 工事名, len(df_p.columns))

    # 保存
    if output_path:
        wb.save(output_path)
        return None
    else:
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()


# ---------------------------------------------------------------------------
# 単体動作確認
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    from extractor import extract_from_pdf

    施工管理基準 = "/Users/tomoki/Downloads/OneDrive_1_2026-5-18/土木工事施工管理基準及び規格値（案）.pdf"
    写真管理基準 = "/Users/tomoki/Downloads/OneDrive_1_2026-5-18/写真管理基準.pdf"
    output      = "output_test.xlsx"

    print("抽出中...")
    data = extract_from_pdf(施工管理基準, 写真管理基準)

    print("Excel生成中...")
    write_excel(data, output_path=output)
    print(f"保存完了: {output}")
