"""
app.py
施工管理計画 自動生成アプリ（Streamlit UI）

起動: python3 -m streamlit run app.py
"""

import re
import tempfile
import traceback
from collections import defaultdict

import pandas as pd
import streamlit as st

from extractor import (
    extract_suryo,
    get_unique_kojyo,
    build_match_detail,
    filter_by_row_labels,
    build_suryo_match_map,
    SURYO_LEVEL_COLS,
)
from excel_writer import write_excel, SHEET_HINSHITSU, SHEET_DEKIGATA, SHEET_PHOTO
from build_db import (
    DB_PATH,
    SHEET_DEKIGATA as DB_DEKIGATA,
    SHEET_HINSHITSU as DB_HINSHITSU,
    SHEET_PHOTO as DB_PHOTO,
    SHEET_VERSION,
)

# ===========================================================================
# ページ設定
# ===========================================================================
st.set_page_config(
    page_title="施工管理計画 自動生成",
    page_icon="☐",
    layout="wide",
)

# ===========================================================================
# カスタム CSS
# ===========================================================================
st.markdown("""
<style>
/* ─── フォント ──────────────────────────────────────────── */
html, body, [class*="css"], .stMarkdown, .stText,
button, input, select, textarea, th, td {
    font-family: 'Yu Gothic', '游ゴシック', YuGothic,
                 'Hiragino Kaku Gothic ProN', 'Hiragino Sans',
                 Meiryo, sans-serif !important;
}
.material-icons, .material-symbols-outlined,
.material-symbols-rounded, [class*="material-icons"] {
    font-family: 'Material Icons', 'Material Symbols Outlined',
                 'Material Symbols Rounded' !important;
}

/* ─── ページ背景 ────────────────────────────────────────── */
.stApp { background-color: #FFFFFF; }
[data-testid="stAppViewContainer"],
[data-testid="stAppViewContainer"]::before,
[data-testid="stAppViewContainer"]::after,
[data-testid="stHeader"],
[data-testid="stHeader"]::before,
[data-testid="stHeader"]::after,
.stApp, .stApp::before, .stApp::after,
.main, .main::before, .main::after {
    background-image: none !important;
    background-color: #FFFFFF !important;
    filter: none !important;
    backdrop-filter: none !important;
}
* { text-shadow: none !important; }

/* ─── メトリクスカード ─────────────────────────────────── */
[data-testid="metric-container"] {
    background: #FFFFFF;
    border: 1px solid #E5E3DC;
    border-radius: 2px;
    padding: 12px 16px;
}
[data-testid="stMetricLabel"] > div {
    font-size: 0.75rem !important;
    font-weight: 600 !important;
    color: #6B6A66 !important;
}

/* ─── プライマリボタン ─────────────────────────────────── */
.stButton > button[kind="primary"] {
    background: #C01820;
    color: #FFFFFF;
    border: none;
    border-radius: 2px;
    font-weight: 600;
    letter-spacing: 0.03em;
}
.stButton > button[kind="primary"]:hover {
    background: #8E1119;
}
.stButton > button[kind="primary"]:disabled {
    background: #9A9893 !important;
}

/* ─── ダウンロードボタン ───────────────────────────────── */
[data-testid="stDownloadButton"] > button {
    background: #FFFFFF !important;
    color: #C01820 !important;
    border: 1px solid #C01820 !important;
    border-radius: 2px !important;
    font-weight: 600 !important;
}
[data-testid="stDownloadButton"] > button:hover {
    background: #FBEBEC !important;
}

/* ─── info ─────────────────────────────────────────────── */
[data-testid="stInfo"] {
    background: #F4F2EE;
    border-left: 3px solid #C01820;
    border-radius: 0;
    color: #2C2C2A;
}

/* ─── サイドバー ───────────────────────────────────────── */
[data-testid="stSidebar"] {
    background: #2B2A28;
    border-right: 1px solid #3A3936;
}
[data-testid="stSidebar"] * { color: #EDEBE6 !important; }
[data-testid="stSidebar"] strong, [data-testid="stSidebar"] b {
    color: #F5F4F0 !important;
}
[data-testid="stSidebar"] h3 {
    color: #C01820 !important;
    font-size: 0.70rem !important;
    letter-spacing: 0.16em !important;
    text-transform: uppercase !important;
    font-weight: 700 !important;
}
[data-testid="stSidebar"] hr { border-color: #3A3936 !important; }
[data-testid="stSidebar"] code {
    background: #1E1D1B !important;
    color: #EDEBE6 !important;
    border: 1px solid #3A3936 !important;
    padding: 1px 5px !important;
    border-radius: 2px !important;
    font-family: Consolas, 'Courier New', monospace !important;
    font-size: 0.80rem !important;
}
[data-testid="stSidebar"] [data-testid="stExpander"] {
    background: #1E1D1B !important;
    border: 1px solid #3A3936 !important;
    border-radius: 2px !important;
}
[data-testid="stSidebar"] [data-testid="stExpander"] summary {
    color: #EDEBE6 !important;
    font-weight: 600 !important;
    font-size: 0.84rem !important;
}
[data-testid="stSidebar"] [data-testid="stExpander"] .stMarkdown p,
[data-testid="stSidebar"] [data-testid="stExpander"] .stMarkdown li {
    font-size: 0.82rem !important;
    line-height: 1.80 !important;
    color: #C8C6C0 !important;
}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
    background: #1E1D1B !important;
    border: 1px dashed #5A5855 !important;
    border-radius: 2px !important;
}
[data-testid="stSidebar"] .stButton > button[kind="primary"] {
    background: #C01820 !important;
    color: #FFFFFF !important;
    border-radius: 2px !important;
}

/* ─── ラジオ・チェックボックス ─────────────────────────── */
[data-testid="stRadio"] label,
[data-testid="stCheckbox"] label { font-size: 0.85rem; }

/* ─── divider ──────────────────────────────────────────── */
hr { border-color: #E5E3DC !important; }
</style>
""", unsafe_allow_html=True)

# ===========================================================================
# 国交省基準DB読み込み
# ===========================================================================
@st.cache_data
def load_kojyo_db():
    if not DB_PATH.exists():
        return None, None
    data = {
        "出来形管理": pd.read_excel(str(DB_PATH), sheet_name=DB_DEKIGATA,  dtype=str).fillna(""),
        "品質管理":   pd.read_excel(str(DB_PATH), sheet_name=DB_HINSHITSU, dtype=str).fillna(""),
        "撮影箇所":   pd.read_excel(str(DB_PATH), sheet_name=DB_PHOTO,     dtype=str).fillna(""),
    }
    df_ver = pd.read_excel(str(DB_PATH), sheet_name=SHEET_VERSION, dtype=str)
    version_info = df_ver.iloc[0].to_dict() if not df_ver.empty else {}
    return data, version_info


kojyo_data, version_info = load_kojyo_db()

if kojyo_data is None:
    st.error(
        "国交省基準データベースが見つかりません。\n"
        "先に python build_db.py を実行してください。"
    )
    st.stop()

unique_kojyo = get_unique_kojyo(kojyo_data)

# ===========================================================================
# セッション初期化
# ===========================================================================
for _k in ["suryo_info", "df_match", "selected_idx"]:
    if _k not in st.session_state:
        st.session_state[_k] = None
if "row_selections" not in st.session_state:
    st.session_state["row_selections"] = {}

# ===========================================================================
# ヘルパー関数
# ===========================================================================
def _chain_key(row) -> tuple:
    return tuple(str(row.get(c, "")) for c in SURYO_LEVEL_COLS)


def _group_items(items: list) -> dict:
    grouped: dict = {}
    for item in items:
        parts = [p.strip() for p in item.split(" / ")]
        kojyo = parts[0]
        sub   = " / ".join(parts[1:]) if len(parts) > 1 else parts[0]
        grouped.setdefault(kojyo, []).append((item, sub))
    return grouped


def _deepest_name(row) -> str:
    for col in reversed(SURYO_LEVEL_COLS):
        v = row.get(col, "")
        if v: return v
    return ""


def _depth(row) -> int:
    d = 0
    for col in SURYO_LEVEL_COLS:
        if row.get(col, ""):
            d = SURYO_LEVEL_COLS.index(col)
    return d


def _sec_header(num: str, label: str) -> None:
    """番号付きセクションヘッダーを描画する。"""
    st.markdown(
        f'<div style="display:flex; align-items:center; gap:10px; '
        f'margin:14px 0 6px 0; border-bottom:2px solid #C01820; padding-bottom:5px;">'
        f'<span style="background:#C01820; color:#FFFFFF; '
        f'font-size:0.70rem; font-weight:700; letter-spacing:0.06em; '
        f'padding:2px 7px; border-radius:1px;">{num}</span>'
        f'<span style="font-weight:700; color:#C01820; font-size:0.92rem; '
        f'letter-spacing:0.01em;">{label}</span>'
        f'</div>',
        unsafe_allow_html=True,
    )


def _subsection_label(label: str) -> None:
    """候補選択内の小見出しを描画する（えんじ統一）。"""
    st.markdown(
        f'<div style="border-bottom:1px solid #C01820; padding:0 0 3px 0; '
        f'font-weight:700; font-size:0.83rem; color:#C01820; '
        f'margin-bottom:6px; letter-spacing:0.02em;">{label}</div>',
        unsafe_allow_html=True,
    )


def _metric_html(label: str, value, color: str = "#6B6A66") -> str:
    return (
        f'<div style="background:#FFFFFF; border:1px solid #E5E3DC; '
        f'border-radius:2px; padding:12px 16px;">'
        f'<div style="font-size:0.75rem; font-weight:600; color:#6B6A66; '
        f'margin-bottom:4px;">{label}</div>'
        f'<div style="font-size:1.5rem; font-weight:700; color:{color};">{value}</div>'
        f'</div>'
    )


# ===========================================================================
# サイドバー
# ===========================================================================
with st.sidebar:
    # ── ヘッダー ─────────────────────────────────────────────
    st.markdown(
        '<div style="padding:14px 0 10px 0; '
        'border-bottom:1px solid #3A3936; margin-bottom:10px;">'
        '<div style="font-size:0.78rem; font-weight:700; color:#C01820; '
        'letter-spacing:0.10em; text-transform:uppercase;">朝日土木株式会社</div>'
        '<div style="font-size:0.92rem; font-weight:700; color:#F5F4F0; '
        'letter-spacing:0.03em; margin-top:3px;">施工管理計画 自動生成</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    # ── DB情報 ───────────────────────────────────────────────
    st.markdown("### 国交省基準 DB")
    st.caption(f"Ver. {version_info.get('バージョン', '不明')}　／　{version_info.get('作成日時', '不明')}")
    st.caption(
        f"出来形管理 {len(kojyo_data['出来形管理'])} 行　"
        f"品質管理 {len(kojyo_data['品質管理'])} 行"
    )
    st.divider()

    # ── アップロード ─────────────────────────────────────────
    st.markdown("### 数量総括表 PDF")
    uploaded_file = st.file_uploader(
        "PDFをアップロード", type="pdf", label_visibility="collapsed"
    )

    if st.button("解析する", type="primary", disabled=not uploaded_file, use_container_width=True):
        try:
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
                f.write(uploaded_file.read())
                path_tmp = f.name

            with st.spinner("解析・マッチング中..."):
                si = extract_suryo(path_tmp)
                dm = build_match_detail(
                    si["工種階層"],
                    kojyo_data["出来形管理"],
                    kojyo_data["品質管理"],
                    kojyo_data["撮影箇所"],
                )

            for k in list(st.session_state.keys()):
                if k.startswith(("chk_d_", "chk_h_", "chk_p_")):
                    del st.session_state[k]

            st.session_state.suryo_info     = si
            st.session_state.df_match       = dm
            st.session_state.selected_idx   = None
            st.session_state.row_selections = {}
            st.rerun()

        except Exception:
            st.error("解析中にエラーが発生しました。")
            with st.expander("エラー詳細"):
                st.code(traceback.format_exc())

    st.divider()

    # ── 使い方（折りたたみ） ──────────────────────────────────
    with st.expander("使い方"):
        st.markdown("""
**①** 数量総括表PDFをアップロードし「解析する」

**②** 01欄でマッチング結果を確認・行をクリック

**③** 03欄でDB候補のチェックを調整

**④** 04欄で「施工管理計画を出力」→ダウンロード

---
基準DB更新時は `build_db.py` を再実行してください。
""")

# ===========================================================================
# ヘッダー
# ===========================================================================
st.markdown(
    '<div style="background-color:#C01820; background-image:none; '
    'color:#FFFFFF; padding:14px 24px; margin-bottom:6px; '
    'filter:none; box-shadow:none;">'
    '<div style="font-size:0.68rem; font-weight:700; letter-spacing:0.14em; '
    'color:rgba(255,255,255,0.7); text-transform:uppercase; margin-bottom:3px;">'
    '朝日土木株式会社</div>'
    '<div style="font-size:1.1rem; font-weight:700; letter-spacing:0.04em; '
    'color:#FFFFFF;">'
    '施工管理計画 自動生成システム</div>'
    '<div style="font-size:0.76rem; color:rgba(255,255,255,0.75); margin-top:3px;">'
    '数量総括表 PDF  →  国交省基準 DB マッピング  →  施工管理計画 Excel 出力'
    '</div></div>',
    unsafe_allow_html=True,
)

if st.session_state.suryo_info:
    name = st.session_state.suryo_info.get("工事名", "")
    st.markdown(
        f'<div style="background:#FBEBEC; border:1px solid #C01820; '
        f'border-radius:2px; padding:8px 16px; margin:4px 0 0 0; '
        f'font-size:0.88rem; color:#8E1119;">'
        f'<strong>{name or "（工事名不明）"}</strong>　読込済み'
        f'</div>',
        unsafe_allow_html=True,
    )

# ===========================================================================
# 未解析時
# ===========================================================================
if st.session_state.df_match is None:
    st.divider()
    st.info("← サイドバーから数量総括表PDFをアップロードして「解析する」を押してください。")
    st.stop()

# ===========================================================================
# データ準備
# ===========================================================================
suryo_info = st.session_state.suryo_info
df_raw     = st.session_state.df_match.copy()


def _status(row) -> str:
    has_d = bool(str(row.get("出来形マッチ", "")).strip())
    has_h = bool(str(row.get("品質管理マッチ", "")).strip())
    has_p = bool(str(row.get("撮影箇所マッチ", "")).strip())
    if not (has_d or has_h or has_p):
        return "対象外"
    if _chain_key(row) in st.session_state.row_selections:
        return "選択済み"
    return "候補あり"


df_raw["状態"]   = df_raw.apply(_status, axis=1)
df_raw["_name"]  = df_raw.apply(_deepest_name, axis=1)
df_raw["_depth"] = df_raw.apply(_depth, axis=1)
df_raw.insert(0, "No", range(1, len(df_raw) + 1))

n_total    = len(df_raw)
n_selected = (df_raw["状態"] == "選択済み").sum()
n_pending  = (df_raw["状態"] == "候補あり").sum()
n_out      = (df_raw["状態"] == "対象外").sum()

# ===========================================================================
# 統計バー
# ===========================================================================
st.divider()
mc = st.columns(4)
mc[0].markdown(_metric_html("総行数",   n_total,    "#6B6A66"), unsafe_allow_html=True)
mc[1].markdown(_metric_html("候補あり", n_pending,  "#C01820"), unsafe_allow_html=True)
mc[2].markdown(_metric_html("選択済み", n_selected, "#8E1119"), unsafe_allow_html=True)
mc[3].markdown(_metric_html("対象外",   n_out,      "#9A9893"), unsafe_allow_html=True)
st.divider()

# 状態ごとの行スタイル
STATUS_BG = {
    "候補あり": "#FBEBEC",
    "選択済み": "#FFFFFF",
    "対象外":   "#F1EFE8",
}
STATUS_COLOR = {
    "候補あり": "#8E1119",
    "選択済み": "#2C2C2A",
    "対象外":   "#9A9893",
}


def _row_style(statuses: dict):
    def _style(row):
        bg = STATUS_BG.get(statuses.get(row.name, ""), "")
        return [f"background-color: {bg}" if bg else "" for _ in row]
    return _style


# ===========================================================================
# 01  数量総括表ツリー
# ===========================================================================
_sec_header("01", "数量総括表（元データ）")
st.caption("「直接工事費」以下は除外済み　　行をクリックすると下部に候補が表示されます")

statuses_tree = df_raw["状態"].reset_index(drop=True).to_dict()
df_tree = pd.DataFrame({
    "工事区分・工種・種別・細別": [
        "　" * row["_depth"] + row["_name"] for _, row in df_raw.iterrows()
    ],
    "状態": list(statuses_tree.values()),
})

ev_l = st.dataframe(
    df_tree.style.apply(_row_style(statuses_tree), axis=1),
    use_container_width=True,
    height=600,
    hide_index=True,
    selection_mode="single-row",
    on_select="rerun",
    column_config={
        "工事区分・工種・種別・細別": st.column_config.TextColumn(width="large"),
        "状態": st.column_config.TextColumn(width="small"),
    },
)
if ev_l.selection.rows:
    st.session_state.selected_idx = ev_l.selection.rows[0]

# 凡例
st.markdown(
    '<div style="font-size:0.76rem; color:#6B6A66; margin-top:4px;">'
    '<span style="background:#FBEBEC; border-left:3px solid #C01820; '
    'padding:1px 8px; margin-right:8px;">候補あり（要対応）</span>'
    '<span style="background:#FFFFFF; border-left:3px solid #C01820; '
    'border:1px solid #E5E3DC; padding:1px 8px; margin-right:8px;">選択済み</span>'
    '<span style="background:#F1EFE8; color:#9A9893; padding:1px 8px;">'
    '対象外</span>'
    '</div>',
    unsafe_allow_html=True,
)

# ===========================================================================
# 02  DB マッピング一覧
# ===========================================================================
_sec_header("02", "DB マッピング一覧")

filter_opt = st.radio(
    "フィルタ",
    ["すべて", "選択済みのみ", "候補ありのみ", "対象外のみ"],
    horizontal=True,
    label_visibility="collapsed",
)
FILTER_MAP = {
    "選択済みのみ": "選択済み",
    "候補ありのみ": "候補あり",
    "対象外のみ":   "対象外",
}
df_view = (
    df_raw[df_raw["状態"] == FILTER_MAP[filter_opt]].copy()
    if filter_opt in FILTER_MAP else df_raw.copy()
)


def _fmt_match(v: str) -> str:
    items = [x.strip() for x in str(v).split("\n") if x.strip()]
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    first_kojyo = items[0].split(" / ")[0]
    return f"{first_kojyo}（他{len(items) - 1}件）"


statuses_map = df_view["状態"].reset_index(drop=True).to_dict()
df_map = pd.DataFrame({
    "工種・行名":    ["　" * d + n for d, n in zip(df_view["_depth"].values, df_view["_name"].values)],
    "出来形マッチ":  [_fmt_match(v) for v in df_view["出来形マッチ"]],
    "品質管理マッチ": [_fmt_match(v) for v in df_view["品質管理マッチ"]],
    "撮影箇所マッチ": [_fmt_match(v) for v in df_view.get("撮影箇所マッチ", [""] * len(df_view))],
    "状態":         list(statuses_map.values()),
})

ev_m = st.dataframe(
    df_map.style.apply(_row_style(statuses_map), axis=1),
    use_container_width=True,
    height=600,
    hide_index=True,
    selection_mode="single-row",
    on_select="rerun",
    column_config={
        "工種・行名":    st.column_config.TextColumn(width="medium"),
        "出来形マッチ":  st.column_config.TextColumn(width="medium"),
        "品質管理マッチ": st.column_config.TextColumn(width="medium"),
        "撮影箇所マッチ": st.column_config.TextColumn(width="medium"),
        "状態":         st.column_config.TextColumn(width="small"),
    },
)
if ev_m.selection.rows:
    selected_no = int(df_view.iloc[ev_m.selection.rows[0]]["No"])
    st.session_state.selected_idx = selected_no - 1

st.divider()

# ===========================================================================
# 03 + 04  下部 2カラム
# ===========================================================================
col_detail, col_out = st.columns([3, 2])

# ── 03  候補選択 ────────────────────────────────────────────────────────
with col_detail:
    _sec_header("03", "候補選択")
    st.caption("行をクリックしてDB候補を確認・チェックを調整してください。")

    sel_idx = st.session_state.selected_idx

    if sel_idx is not None and 0 <= sel_idx < len(df_raw):
        sel  = df_raw.iloc[sel_idx]
        ckey = _chain_key(sel)

        chain = " › ".join(sel[c] for c in SURYO_LEVEL_COLS if sel.get(c, ""))
        st.markdown(
            f'<div style="background:#F4F2EE; border-left:3px solid #C01820; '
            f'padding:7px 12px; margin-bottom:10px; '
            f'font-size:0.87rem; color:#2C2C2A;">'
            f'<strong>{sel["_name"]}</strong><br>'
            f'<span style="font-size:0.76rem; color:#6B6A66;">{chain}</span>'
            f'</div>',
            unsafe_allow_html=True,
        )

        match_d = str(sel.get("出来形マッチ", "")).strip()
        match_h = str(sel.get("品質管理マッチ", "")).strip()
        match_p = str(sel.get("撮影箇所マッチ", "")).strip()
        items_d = [x.strip() for x in match_d.split("\n") if x.strip()]
        items_h = [x.strip() for x in match_h.split("\n") if x.strip()]
        items_p = [x.strip() for x in match_p.split("\n") if x.strip()]

        if not items_d and not items_h and not items_p:
            st.caption("この行はDBマッチなし（対象外）")
            sel_d, sel_h, sel_p = [], [], []
        else:
            st.caption("出力に含める候補にチェックを入れてください。")

            c_d, c_h, c_p = st.columns(3)

            with c_d:
                sel_d = []
                if items_d:
                    _subsection_label("出来形管理")
                    grouped_d = _group_items(items_d)
                    for kojyo, sub_items in grouped_d.items():
                        if len(grouped_d) > 1:
                            st.caption(kojyo)
                        for full_label, display_label in sub_items:
                            chk_key = f"chk_d_{sel_idx}_{items_d.index(full_label)}"
                            if st.checkbox(display_label, value=True, key=chk_key):
                                sel_d.append(full_label)
                else:
                    st.caption("出来形管理：該当なし")

            with c_h:
                sel_h = []
                if items_h:
                    _subsection_label("品質管理")
                    grouped_h = _group_items(items_h)
                    for kojyo, sub_items in grouped_h.items():
                        if len(grouped_h) > 1:
                            st.caption(kojyo)
                        for full_label, display_label in sub_items:
                            chk_key = f"chk_h_{sel_idx}_{items_h.index(full_label)}"
                            if st.checkbox(display_label, value=True, key=chk_key):
                                sel_h.append(full_label)
                else:
                    st.caption("品質管理：該当なし")

            with c_p:
                sel_p = []
                if items_p:
                    _subsection_label("撮影箇所")
                    grouped_p = _group_items(items_p)
                    for kojyo, sub_items in grouped_p.items():
                        if len(grouped_p) > 1:
                            st.caption(kojyo)
                        for full_label, display_label in sub_items:
                            chk_key = f"chk_p_{sel_idx}_{items_p.index(full_label)}"
                            if st.checkbox(display_label, value=True, key=chk_key):
                                sel_p.append(full_label)
                else:
                    st.caption("撮影箇所：該当なし")

            st.session_state.row_selections[ckey] = {
                "出来形":   sel_d,
                "品質管理":  sel_h,
                "撮影箇所": sel_p,
            }

        if items_d:
            first_d = items_d[0].split(" / ")[0]
            db_rows = kojyo_data["出来形管理"][kojyo_data["出来形管理"]["工種"] == first_d]
            if not db_rows.empty:
                r  = db_rows.iloc[0]
                bc = " › ".join(
                    x for x in [r.get("編", ""), r.get("章", ""), r.get("節", ""), first_d] if x
                )
                st.caption(f"DB 目次位置：{bc}")

    else:
        st.info("上の表から行をクリックすると候補が表示されます。")

# ── 04  出力 ─────────────────────────────────────────────────────────────
with col_out:
    _sec_header("04", "施工管理計画を出力")

    def _collect_labels():
        out_d: list = []
        out_h: list = []
        out_p: list = []
        seen_d: set = set()
        seen_h: set = set()
        seen_p: set = set()

        for _, row in df_raw[df_raw["状態"].isin(["選択済み", "候補あり"])].iterrows():
            ckey  = _chain_key(row)
            saved = st.session_state.row_selections.get(ckey)

            all_d = [x.strip() for x in str(row.get("出来形マッチ", "")).split("\n") if x.strip()]
            all_h = [x.strip() for x in str(row.get("品質管理マッチ", "")).split("\n") if x.strip()]
            all_p = [x.strip() for x in str(row.get("撮影箇所マッチ", "")).split("\n") if x.strip()]

            labels_d = saved["出来形"]              if saved is not None else all_d
            labels_h = saved["品質管理"]            if saved is not None else all_h
            labels_p = saved.get("撮影箇所", all_p) if saved is not None else all_p

            for lbl in labels_d:
                if lbl and lbl not in seen_d:
                    out_d.append(lbl); seen_d.add(lbl)
            for lbl in labels_h:
                if lbl and lbl not in seen_h:
                    out_h.append(lbl); seen_h.add(lbl)
            for lbl in labels_p:
                if lbl and lbl not in seen_p:
                    out_p.append(lbl); seen_p.add(lbl)

        return out_d, out_h, out_p

    out_d_labels, out_h_labels, out_p_labels = _collect_labels()

    st.markdown(
        f'<div style="background:#F4F2EE; border:1px solid #E5E3DC; '
        f'padding:10px 14px; margin-bottom:12px; border-radius:2px;">'
        f'<table style="width:100%; font-size:0.84rem; border:none; '
        f'border-collapse:collapse; color:#2C2C2A;">'
        f'<tr><td style="padding:3px 0;">出来形管理</td>'
        f'<td style="text-align:right; font-weight:700; color:#C01820;">{len(out_d_labels)} 項目</td></tr>'
        f'<tr><td style="padding:3px 0;">品質管理</td>'
        f'<td style="text-align:right; font-weight:700; color:#C01820;">{len(out_h_labels)} 項目</td></tr>'
        f'<tr><td style="padding:3px 0;">撮影箇所</td>'
        f'<td style="text-align:right; font-weight:700; color:#C01820;">{len(out_p_labels)} 項目</td></tr>'
        f'</table>'
        f'<div style="font-size:0.72rem; color:#6B6A66; margin-top:6px; '
        f'border-top:1px solid #E5E3DC; padding-top:5px;">'
        f'未確認行は全候補を採用</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    if st.button(
        "施工管理計画を出力",
        type="primary",
        use_container_width=True,
        disabled=not (out_d_labels or out_h_labels or out_p_labels),
    ):
        try:
            with st.spinner("Excel 生成中..."):
                filtered = filter_by_row_labels(
                    kojyo_data, out_d_labels, out_h_labels, out_p_labels
                )
                out_d_kojyo = list({lbl.split(" / ")[0].strip() for lbl in out_d_labels})
                dmap = build_suryo_match_map(suryo_info["工種リスト"], out_d_kojyo)
                excel_bytes = write_excel(
                    filtered,
                    工事名=suryo_info["工事名"],
                    dekigata_kojyo_map=dmap,
                )

            safe  = re.sub(r'[\\/:*?"<>|　 ]', "_", suryo_info["工事名"])
            fname = f"施工管理計画_{safe}.xlsx" if safe else "施工管理計画.xlsx"

            st.download_button(
                label=f"ダウンロード  {fname}",
                data=excel_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        except Exception:
            st.error("Excel 生成中にエラーが発生しました。")
            with st.expander("エラー詳細"):
                st.code(traceback.format_exc())
