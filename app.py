"""
app.py
施工管理計画 自動生成アプリ（Streamlit UI）

起動: python3 -m streamlit run app.py
"""

import re
import tempfile
import traceback

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
/* ─── フォント ──────────────────────────────────────────────── */
html, body, [class*="css"], .stMarkdown, .stText,
button, input, select, textarea, th, td {
    font-family: 'Yu Gothic', '游ゴシック', YuGothic,
                 'Hiragino Kaku Gothic ProN', 'Hiragino Sans',
                 Meiryo, sans-serif !important;
}

/* ─── ページ背景 ─────────────────────────────────────────────── */
.stApp,
[data-testid="stAppViewContainer"],
.main { background-color: #F5F6F8 !important; }
[data-testid="stAppViewContainer"]::before,
[data-testid="stAppViewContainer"]::after,
[data-testid="stHeader"],
[data-testid="stHeader"]::before,
[data-testid="stHeader"]::after {
    background-image: none !important;
    filter: none !important;
    backdrop-filter: none !important;
}
* { text-shadow: none !important; }
.block-container { padding-top: 0 !important; padding-bottom: 2rem !important; }

/* ─── サイドバー（ライト） ───────────────────────────────────── */
[data-testid="stSidebar"] {
    background: #FFFFFF !important;
    border-right: 1px solid #E8EAED !important;
    min-width: 200px !important;
    max-width: 200px !important;
}
[data-testid="stSidebar"] * { color: #333333 !important; }
[data-testid="stSidebar"] h3 {
    color: #AAAAAA !important;
    font-size: 0.62rem !important;
    letter-spacing: 0.18em !important;
    text-transform: uppercase !important;
    font-weight: 700 !important;
    margin: 10px 0 6px 0 !important;
}
[data-testid="stSidebar"] hr { border-color: #EEEEEE !important; }
[data-testid="stSidebar"] code {
    background: #F4F6F8 !important;
    color: #1565C0 !important;
    border: 1px solid #D8E4F0 !important;
    padding: 1px 5px !important;
    border-radius: 3px !important;
    font-size: 0.78rem !important;
}
[data-testid="stSidebar"] [data-testid="stExpander"] {
    background: #F8F9FA !important;
    border: 1px solid #EEEEEE !important;
    border-radius: 6px !important;
}
[data-testid="stSidebar"] [data-testid="stExpander"] summary {
    color: #555555 !important;
    font-size: 0.84rem !important;
}
[data-testid="stSidebar"] [data-testid="stExpander"] .stMarkdown p,
[data-testid="stSidebar"] [data-testid="stExpander"] .stMarkdown li {
    font-size: 0.82rem !important;
    line-height: 1.8 !important;
    color: #555555 !important;
}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
    background: #F8F9FA !important;
    border: 1px dashed #CCCCCC !important;
    border-radius: 5px !important;
}

/* ─── サイドバーナビボタン ──────────────────────────────────── */
[data-testid="stSidebar"] .stButton > button {
    width: 100% !important;
    border-radius: 6px !important;
    font-size: 0.85rem !important;
    font-weight: 500 !important;
    padding: 8px 12px !important;
    text-align: left !important;
    border: none !important;
    transition: background 0.15s;
}
[data-testid="stSidebar"] .stButton > button[kind="primary"] {
    background: #EEF4FF !important;
    color: #1565C0 !important;
    font-weight: 700 !important;
}
[data-testid="stSidebar"] .stButton > button[kind="secondary"] {
    background: transparent !important;
    color: #444444 !important;
}
[data-testid="stSidebar"] .stButton > button[kind="secondary"]:hover {
    background: #F0F2F5 !important;
}

/* ─── プライマリボタン（メインエリア） ──────────────────────── */
.stButton > button[kind="primary"] {
    background: #1565C0;
    color: #FFFFFF;
    border: none;
    border-radius: 5px;
    font-weight: 600;
    letter-spacing: 0.03em;
}
.stButton > button[kind="primary"]:hover { background: #0D47A1; }
.stButton > button[kind="primary"]:disabled { background: #CCCCCC !important; color: #888 !important; }

/* ─── ダウンロードボタン ─────────────────────────────────────── */
[data-testid="stDownloadButton"] > button {
    background: #FFFFFF !important; color: #1565C0 !important;
    border: 1.5px solid #1565C0 !important; border-radius: 5px !important;
    font-weight: 600 !important;
}
[data-testid="stDownloadButton"] > button:hover { background: #E3F2FD !important; }

/* ─── info ──────────────────────────────────────────────────── */
[data-testid="stInfo"] {
    background: #E8F0FE; border-left: 3px solid #1565C0;
    border-radius: 5px; color: #1A2B3C;
}

/* ─── ラジオ・チェック ─────────────────────────────────────── */
[data-testid="stRadio"] label,
[data-testid="stCheckbox"] label { font-size: 0.85rem; color: #333333; }

/* ─── divider ───────────────────────────────────────────────── */
hr { border-color: #E8EAED !important; }

/* ─── ステップバー ───────────────────────────────────────────── */
.stepbar-wrap {
    background: #FFFFFF;
    border-bottom: 1px solid #E8EAED;
    padding: 12px 24px;
    margin: 0 -1rem 20px -1rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
}
.stepbar-title {
    font-size: 1.0rem; font-weight: 800; color: #1A2332;
    display: flex; align-items: center; gap: 8px; white-space: nowrap;
}
.stepbar-title .icon {
    width: 28px; height: 28px; background: #1565C0; border-radius: 5px;
    display: inline-flex; align-items: center; justify-content: center;
    color: #FFFFFF; font-size: 0.70rem; font-weight: 700;
}
.stepbar-steps {
    display: flex; align-items: center; gap: 2px;
    font-size: 0.72rem;
}
.step-seg {
    display: flex; align-items: center; gap: 3px;
    padding: 3px 8px; border-radius: 12px;
}
.step-seg.done  { color: #888888; }
.step-seg.active { background: #EEF4FF; color: #1565C0; font-weight: 700; border: 1px solid #C7D9F8; }
.step-seg.pending { color: #CCCCCC; }
.step-sep { color: #DDDDDD; margin: 0 1px; }

/* ─── メトリクスカード ───────────────────────────────────────── */
.metric-card {
    background: #FFFFFF; border-radius: 8px;
    border: 1px solid #E8EAED; padding: 18px 22px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
}
.m-label { font-size: 0.72rem; font-weight: 600; color: #888888;
           text-transform: uppercase; letter-spacing: 0.08em; margin-bottom: 6px; }
.m-value { font-size: 2.6rem; font-weight: 800; line-height: 1; }
.metric-card.kakutei   .m-value { color: #1B6E2A; }
.metric-card.yosentaku .m-value { color: #B45309; }
.metric-card.mimatch   .m-value { color: #9E9E9E; }

/* ─── テーブルラッパー ──────────────────────────────────────── */
.table-wrap {
    background: #FFFFFF; border: 1px solid #E8EAED; border-radius: 8px;
    overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    margin-bottom: 12px;
}

/* ─── 状態バッジ ─────────────────────────────────────────────── */
.badge {
    display: inline-block; padding: 2px 10px; border-radius: 20px;
    font-size: 0.72rem; font-weight: 600; white-space: nowrap;
}
.badge-kakutei   { background: #D1FAE5; color: #065F46; }
.badge-yosentaku { background: #FEF3C7; color: #92400E; }
.badge-mimatch   { background: #F3F4F6; color: #6B7280; }

/* ─── 候補パネル ─────────────────────────────────────────────── */
.cand-panel {
    background: #FFFBEB; border: 1.5px solid #F59E0B; border-radius: 8px;
    padding: 16px 20px; margin-bottom: 16px;
}
.cand-panel-header {
    font-size: 0.88rem; font-weight: 700; color: #92400E;
    margin-bottom: 14px; display: flex; align-items: center; gap: 6px;
}
.cand-card {
    background: #FFFFFF; border: 1px solid #E8EAED; border-radius: 6px;
    padding: 14px 16px;
}
.cand-card.selected { border-color: #1565C0; background: #F0F6FF; }
.cand-card-title { font-size: 0.88rem; font-weight: 700; color: #1A2332; margin-bottom: 8px; }
.cand-card-detail { font-size: 0.79rem; color: #555555; line-height: 1.9; }
.diff-tag {
    display: inline-block; background: #DBEAFE; color: #1E40AF;
    border-radius: 3px; padding: 0px 6px; font-size: 0.75rem; font-weight: 600;
}
.cand-panel-foot {
    font-size: 0.72rem; color: #888888; margin-top: 10px;
}

/* ─── 凡例 ──────────────────────────────────────────────────── */
.legend {
    display: flex; gap: 14px; font-size: 0.74rem; color: #555555;
    align-items: center; margin-top: 6px;
}
.legend-dot { width: 10px; height: 10px; border-radius: 2px; display: inline-block; margin-right: 4px; }

/* ─── セクションヘッダー ─────────────────────────────────────── */
.sec-header {
    display: flex; align-items: center; gap: 8px;
    margin: 16px 0 8px 0; padding-bottom: 6px;
    border-bottom: 2px solid #1565C0;
}
.sec-title { font-weight: 700; color: #1565C0; font-size: 0.90rem; }

/* ─── 選択情報パネル ─────────────────────────────────────────── */
.sel-info {
    background: #F0F4F8; border-left: 3px solid #1565C0;
    padding: 10px 14px; margin-bottom: 10px; border-radius: 0 5px 5px 0;
}
.sel-info .sel-name { font-size: 0.90rem; font-weight: 700; color: #1A2332; }
.sel-info .sel-chain { font-size: 0.74rem; color: #666666; margin-top: 3px; }

/* ─── subsec label ──────────────────────────────────────────── */
.subsec-label {
    border-bottom: 1px solid #1565C0; padding-bottom: 3px;
    font-weight: 700; font-size: 0.82rem; color: #333333;
    margin-bottom: 6px;
}

/* ─── 出力サマリーカード ─────────────────────────────────────── */
.out-card {
    background: #FFFFFF; border: 1px solid #E8EAED; border-radius: 8px;
    padding: 16px 20px; margin-bottom: 14px;
}
.out-card table { width: 100%; border-collapse: collapse; font-size: 0.85rem; }
.out-card td { padding: 5px 0; color: #555555; }
.out-card td.num { text-align: right; font-weight: 700; color: #1565C0; }
.out-card .note { font-size: 0.72rem; color: #999999; margin-top: 8px;
                  border-top: 1px solid #F0F0F0; padding-top: 6px; }
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
    st.error("国交省基準データベースが見つかりません。先に python build_db.py を実行してください。")
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
if "page" not in st.session_state:
    st.session_state["page"] = "matching"

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
        if v:
            return v
    return ""


def _depth(row) -> int:
    d = 0
    for col in SURYO_LEVEL_COLS:
        if row.get(col, ""):
            d = SURYO_LEVEL_COLS.index(col)
    return d


def _sec_header(label: str) -> None:
    st.markdown(
        f'<div class="sec-header"><span class="sec-title">{label}</span></div>',
        unsafe_allow_html=True,
    )


def _subsection_label(label: str) -> None:
    st.markdown(f'<div class="subsec-label">{label}</div>', unsafe_allow_html=True)


# ===========================================================================
# ステータス定義
# ===========================================================================
def _calc_status(row) -> str:
    has_d = bool(str(row.get("出来形マッチ", "")).strip())
    has_h = bool(str(row.get("品質管理マッチ", "")).strip())
    has_p = bool(str(row.get("撮影箇所マッチ", "")).strip())
    if not (has_d or has_h or has_p):
        return "未マッチ"
    if _chain_key(row) in st.session_state.row_selections:
        return "確定"
    return "要選択"


STATUS_BG = {
    "確定":   "#F0FDF4",
    "要選択": "#FFFBEB",
    "未マッチ": "#F9FAFB",
}

# ===========================================================================
# DB詳細ルックアップ（候補カード用）
# ===========================================================================
_DB_DISP_COLS_D = ["測定項目", "規格値", "管理基準値", "測定頻度", "摘要"]
_DB_DISP_COLS_H = ["試験項目", "試験方法", "試験基準", "摘要"]
_DB_DISP_COLS_P = ["撮影箇所", "提出頻度", "摘要"]


def _lookup_db_row(label: str, db_key: str) -> dict:
    parts = [p.strip() for p in label.split(" / ")]
    df_db = kojyo_data[db_key]
    level_cols = ["工種", "種別", "細別"]
    mask = pd.Series([True] * len(df_db), index=df_db.index)
    for i, part in enumerate(parts):
        col = level_cols[i] if i < len(level_cols) else None
        if col and col in df_db.columns:
            mask = mask & (df_db[col] == part)
    rows = df_db[mask]
    return rows.iloc[0].to_dict() if not rows.empty else {}


def _card_details_html(row_dict: dict, disp_cols: list, diff_keys: set) -> str:
    html = ""
    for col in disp_cols:
        val = str(row_dict.get(col, "")).strip()
        if not val:
            continue
        if col in diff_keys:
            html += (
                f'<div><span class="diff-tag">{val}</span></div>'
            )
        else:
            html += f'<div>{col}：{val}</div>'
    return html if html else '<div style="color:#AAAAAA;">詳細情報なし</div>'


def _find_diff_cols(rows: list[dict], disp_cols: list) -> set:
    diff = set()
    if len(rows) < 2:
        return diff
    for col in disp_cols:
        vals = [str(r.get(col, "")).strip() for r in rows]
        if len(set(vals)) > 1:
            diff.add(col)
    return diff

# ===========================================================================
# ステップバー
# ===========================================================================
def _stepbar(current: int) -> None:
    steps = [
        ("①", "基準DB"),
        ("②", "取込"),
        ("③", "構造化"),
        ("④", "マッチング"),
        ("⑤", "候補選択"),
        ("…", "出力"),
    ]
    segs_html = ""
    for i, (num, label) in enumerate(steps):
        n = i + 1
        if n < current:
            cls = "done"
        elif n == current:
            cls = "active"
        else:
            cls = "pending"
        segs_html += f'<span class="step-seg {cls}">{num} {label}</span>'
        if i < len(steps) - 1:
            segs_html += '<span class="step-sep">›</span>'

    st.markdown(
        f'<div class="stepbar-wrap">'
        f'<div class="stepbar-title">'
        f'<span class="icon">施</span>'
        f'施工管理計画 自動生成'
        f'</div>'
        f'<div class="stepbar-steps">{segs_html}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

# ===========================================================================
# サイドバー
# ===========================================================================
with st.sidebar:
    st.markdown(
        '<div style="padding:16px 4px 12px 4px; border-bottom:1px solid #EEEEEE; '
        'margin-bottom:12px;">'
        '<div style="font-size:0.80rem; font-weight:800; color:#1A2332; '
        'letter-spacing:0.02em;">施工管理計画</div>'
        '<div style="font-size:0.68rem; color:#AAAAAA; margin-top:2px;">'
        '自動生成システム</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    page = st.session_state.get("page", "matching")

    if st.button(
        "📊  マッチング",
        use_container_width=True,
        type="primary" if page == "matching" else "secondary",
        key="nav_matching",
    ):
        st.session_state.page = "matching"
        st.rerun()

    if st.button(
        "🗄  基準DB確認",
        use_container_width=True,
        type="primary" if page == "db_view" else "secondary",
        key="nav_db",
    ):
        st.session_state.page = "db_view"
        st.rerun()

    st.markdown('<div style="height:4px"></div>', unsafe_allow_html=True)

    if st.button("↺  リセット", use_container_width=True, key="nav_reset"):
        for k in list(st.session_state.keys()):
            if k not in ("page",):
                del st.session_state[k]
        st.session_state.suryo_info     = None
        st.session_state.df_match       = None
        st.session_state.selected_idx   = None
        st.session_state.row_selections = {}
        st.rerun()

    with st.expander("？  使い方"):
        st.markdown("""
**①** 数量総括表PDFをアップロード

**②** 「解析する」でマッチング実行

**③** 表の行をクリックして候補を確認

**④** 「出力」でExcel生成

---
基準DB更新時は `build_db.py` を再実行
""")

    st.divider()

    # ── ファイルアップロード ───────────────────────────────────
    st.markdown("### PDF アップロード")
    uploaded_file = st.file_uploader(
        "数量総括表PDF", type="pdf", label_visibility="collapsed"
    )
    if st.button(
        "解析する",
        type="primary",
        disabled=not uploaded_file,
        use_container_width=True,
    ):
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
            st.session_state.page           = "matching"
            st.rerun()
        except Exception:
            st.error("解析エラー")
            with st.expander("詳細"):
                st.code(traceback.format_exc())

    st.divider()
    st.markdown("### 国交省基準 DB")
    st.caption(f"Ver. {version_info.get('バージョン', '不明')}")
    st.caption(
        f"出来形 {len(kojyo_data['出来形管理'])} 行　"
        f"品質 {len(kojyo_data['品質管理'])} 行"
    )


# ===========================================================================
# ステップバー（現在ステップ算出）
# ===========================================================================
def _compute_step() -> int:
    if st.session_state.df_match is None:
        return 1 if st.session_state.suryo_info is None else 2
    df_tmp = st.session_state.df_match
    n_yo = sum(1 for _, r in df_tmp.iterrows() if _calc_status(r) == "要選択")
    return 5 if n_yo > 0 else 6


_stepbar(_compute_step())


# ===========================================================================
# 基準DB確認ページ
# ===========================================================================
def _render_db_view():
    st.markdown(
        '<div style="font-size:1.0rem; font-weight:700; color:#1A2332; '
        'margin-bottom:12px;">🗄 国交省基準 DB</div>',
        unsafe_allow_html=True,
    )
    tab_d, tab_h, tab_p = st.tabs(["出来形管理", "品質管理", "撮影箇所"])
    with tab_d:
        st.dataframe(kojyo_data["出来形管理"], use_container_width=True, height=580, hide_index=True)
    with tab_h:
        st.dataframe(kojyo_data["品質管理"], use_container_width=True, height=580, hide_index=True)
    with tab_p:
        st.dataframe(kojyo_data["撮影箇所"], use_container_width=True, height=580, hide_index=True)


# ===========================================================================
# マッチングページ
# ===========================================================================
def _render_matching():
    if st.session_state.df_match is None:
        st.info("← サイドバーから数量総括表PDFをアップロードして「解析する」を押してください。")
        return

    suryo_info = st.session_state.suryo_info
    df_raw     = st.session_state.df_match.copy()

    df_raw["状態"]   = df_raw.apply(_calc_status, axis=1)
    df_raw["_name"]  = df_raw.apply(_deepest_name, axis=1)
    df_raw["_depth"] = df_raw.apply(_depth, axis=1)
    df_raw.insert(0, "No", range(1, len(df_raw) + 1))

    n_kakutei   = int((df_raw["状態"] == "確定").sum())
    n_yosentaku = int((df_raw["状態"] == "要選択").sum())
    n_mimatch   = int((df_raw["状態"] == "未マッチ").sum())

    # ── 工事名バナー ──────────────────────────────────────────
    if suryo_info:
        name = suryo_info.get("工事名", "")
        if name:
            st.markdown(
                f'<div style="background:#E8F0FE; border:1px solid #C7D9F8; '
                f'border-radius:6px; padding:8px 16px; margin-bottom:14px; '
                f'font-size:0.86rem; color:#1565C0;">'
                f'<strong>{name}</strong>　読込済み'
                f'</div>',
                unsafe_allow_html=True,
            )

    # ── メトリクス ───────────────────────────────────────────
    mc = st.columns(3)
    mc[0].markdown(
        f'<div class="metric-card kakutei">'
        f'<div class="m-label">確定</div>'
        f'<div class="m-value">{n_kakutei}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )
    mc[1].markdown(
        f'<div class="metric-card yosentaku">'
        f'<div class="m-label">要選択</div>'
        f'<div class="m-value">{n_yosentaku}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )
    mc[2].markdown(
        f'<div class="metric-card mimatch">'
        f'<div class="m-label">未マッチ</div>'
        f'<div class="m-value">{n_mimatch}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )
    st.markdown('<div style="height:14px"></div>', unsafe_allow_html=True)

    # ── フィルタ ─────────────────────────────────────────────
    filter_opt = st.radio(
        "フィルタ",
        ["すべて", "要選択のみ", "確定のみ", "未マッチのみ"],
        horizontal=True,
        label_visibility="collapsed",
    )
    FILTER_MAP = {"確定のみ": "確定", "要選択のみ": "要選択", "未マッチのみ": "未マッチ"}
    df_view = (
        df_raw[df_raw["状態"] == FILTER_MAP[filter_opt]].copy()
        if filter_opt in FILTER_MAP else df_raw.copy()
    )

    # ── テーブル ─────────────────────────────────────────────
    def _fmt_match_cell(row) -> str:
        d = str(row.get("出来形マッチ", "")).strip()
        h = str(row.get("品質管理マッチ", "")).strip()
        p = str(row.get("撮影箇所マッチ", "")).strip()
        if not (d or h or p):
            return "—"
        cats = []
        if d: cats.append("出来形")
        if h: cats.append("品質")
        if p: cats.append("撮影")
        # Count total candidates in 出来形
        d_lines = [x for x in d.split("\n") if x.strip()]
        if len(d_lines) > 1:
            return f"候補{len(d_lines)}件（工法で分岐）"
        return "・".join(cats)

    statuses_idx = {i: row["状態"] for i, (_, row) in enumerate(df_view.iterrows())}

    df_table = pd.DataFrame({
        "工種・項目": [
            "　" * row["_depth"] + row["_name"]
            for _, row in df_view.iterrows()
        ],
        "マッチした基準": [_fmt_match_cell(row) for _, row in df_view.iterrows()],
        "状態": [row["状態"] for _, row in df_view.iterrows()],
    })

    def _row_style(statuses: dict):
        def _style(row):
            bg = STATUS_BG.get(statuses.get(row.name, ""), "")
            return [f"background-color: {bg}" if bg else "" for _ in row]
        return _style

    ev = st.dataframe(
        df_table.style.apply(_row_style(statuses_idx), axis=1),
        use_container_width=True,
        height=min(400, max(180, len(df_view) * 35 + 50)),
        hide_index=True,
        selection_mode="single-row",
        on_select="rerun",
        column_config={
            "工種・項目":    st.column_config.TextColumn(width="large"),
            "マッチした基準": st.column_config.TextColumn(width="large"),
            "状態":         st.column_config.TextColumn(width="small"),
        },
    )
    if ev.selection.rows:
        selected_no = int(df_view.iloc[ev.selection.rows[0]]["No"])
        st.session_state.selected_idx = selected_no - 1

    # 凡例
    st.markdown(
        '<div class="legend">'
        '<span><span class="legend-dot" style="background:#1B6E2A;"></span>確定</span>'
        '<span><span class="legend-dot" style="background:#B45309;"></span>要選択</span>'
        '<span><span class="legend-dot" style="background:#9E9E9E;"></span>未マッチ</span>'
        '</div>',
        unsafe_allow_html=True,
    )
    st.markdown('<div style="height:16px"></div>', unsafe_allow_html=True)

    # ── 候補パネル ───────────────────────────────────────────
    sel_idx = st.session_state.selected_idx
    if sel_idx is not None and 0 <= sel_idx < len(df_raw):
        sel  = df_raw.iloc[sel_idx]
        ckey = _chain_key(sel)

        match_d = str(sel.get("出来形マッチ", "")).strip()
        match_h = str(sel.get("品質管理マッチ", "")).strip()
        match_p = str(sel.get("撮影箇所マッチ", "")).strip()
        items_d = [x.strip() for x in match_d.split("\n") if x.strip()]
        items_h = [x.strip() for x in match_h.split("\n") if x.strip()]
        items_p = [x.strip() for x in match_p.split("\n") if x.strip()]

        n_remaining = n_yosentaku

        if not items_d and not items_h and not items_p:
            st.info(f"「{sel['_name']}」はDBマッチなし（未マッチ）")
        else:
            # 候補パネルヘッダー
            st.markdown(
                f'<div class="cand-panel">'
                f'<div class="cand-panel-header">'
                f'⚠ {sel["_name"]} — 候補を確認・選択してください（残り{n_remaining}件）'
                f'</div>',
                unsafe_allow_html=True,
            )

            # 出来形候補カード（複数あれば比較表示）
            saved = st.session_state.row_selections.get(ckey)

            if len(items_d) >= 2:
                db_rows = [_lookup_db_row(lbl, "出来形管理") for lbl in items_d[:4]]
                diff_keys = _find_diff_cols(db_rows, _DB_DISP_COLS_D)
                diff_label = "・".join(sorted(diff_keys)) if diff_keys else ""

                n_cards = min(len(items_d), 4)
                cols = st.columns(n_cards)
                sel_d = saved["出来形"] if saved else items_d  # default: all

                new_sel_d = []
                for i, (col, lbl) in enumerate(zip(cols, items_d[:n_cards])):
                    with col:
                        parts = [p.strip() for p in lbl.split(" / ")]
                        card_title = " / ".join(parts[1:]) if len(parts) > 1 else parts[0]
                        db_r = db_rows[i]
                        details_html = _card_details_html(db_r, _DB_DISP_COLS_D, diff_keys)
                        is_sel = lbl in sel_d
                        border_css = "border:1.5px solid #1565C0; background:#F0F6FF;" if is_sel else "border:1px solid #E8EAED;"
                        st.markdown(
                            f'<div class="cand-card" style="{border_css}">'
                            f'<div class="cand-card-title">候補{chr(65+i)}：{card_title}</div>'
                            f'<div class="cand-card-detail">{details_html}</div>'
                            f'</div>',
                            unsafe_allow_html=True,
                        )
                        chk_key = f"chk_d_{sel_idx}_{i}"
                        checked = st.checkbox(
                            "採用",
                            value=is_sel,
                            key=chk_key,
                        )
                        if checked:
                            new_sel_d.append(lbl)

                if diff_label:
                    st.markdown(
                        f'<div class="cand-panel-foot">ⓘ 差分（{diff_label}）をハイライト表示</div>',
                        unsafe_allow_html=True,
                    )
            else:
                # 単一候補 or なし
                new_sel_d = items_d

            # 品質管理・撮影箇所は従来のチェックボックス方式
            st.markdown('</div>', unsafe_allow_html=True)  # close cand-panel (partial)

            # 品質管理・撮影箇所のチェックボックスをexpanderに
            if items_h or items_p:
                with st.expander("品質管理・撮影箇所の候補を調整"):
                    c_h, c_p = st.columns(2)
                    with c_h:
                        new_sel_h = []
                        if items_h:
                            _subsection_label("品質管理")
                            grouped_h = _group_items(items_h)
                            for kojyo, sub_items in grouped_h.items():
                                if len(grouped_h) > 1:
                                    st.caption(kojyo)
                                for full_label, display_label in sub_items:
                                    chk_key = f"chk_h_{sel_idx}_{items_h.index(full_label)}"
                                    if st.checkbox(display_label, value=True, key=chk_key):
                                        new_sel_h.append(full_label)
                        else:
                            st.caption("品質管理：該当なし")
                            new_sel_h = []
                    with c_p:
                        new_sel_p = []
                        if items_p:
                            _subsection_label("撮影箇所")
                            grouped_p = _group_items(items_p)
                            for kojyo, sub_items in grouped_p.items():
                                if len(grouped_p) > 1:
                                    st.caption(kojyo)
                                for full_label, display_label in sub_items:
                                    chk_key = f"chk_p_{sel_idx}_{items_p.index(full_label)}"
                                    if st.checkbox(display_label, value=True, key=chk_key):
                                        new_sel_p.append(full_label)
                        else:
                            st.caption("撮影箇所：該当なし")
                            new_sel_p = []
            else:
                new_sel_h = items_h
                new_sel_p = items_p

            # 選択を保存
            st.session_state.row_selections[ckey] = {
                "出来形":   new_sel_d,
                "品質管理":  new_sel_h,
                "撮影箇所": new_sel_p,
            }

            # DB 目次位置
            if items_d:
                first_d = items_d[0].split(" / ")[0]
                db_rows_ref = kojyo_data["出来形管理"][kojyo_data["出来形管理"]["工種"] == first_d]
                if not db_rows_ref.empty:
                    r = db_rows_ref.iloc[0]
                    bc = " › ".join(
                        x for x in [r.get("編", ""), r.get("章", ""), r.get("節", ""), first_d] if x
                    )
                    st.caption(f"DB 目次：{bc}")
    else:
        st.info("上の表から行をクリックすると候補が表示されます。")

    st.divider()

    # ── 出力セクション ───────────────────────────────────────
    _sec_header("施工管理計画を出力")

    def _collect_labels():
        out_d, out_h, out_p = [], [], []
        seen_d, seen_h, seen_p = set(), set(), set()
        for _, row in df_raw[df_raw["状態"].isin(["確定", "要選択"])].iterrows():
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

    col_info, col_btn = st.columns([2, 1])
    with col_info:
        st.markdown(
            f'<div class="out-card">'
            f'<table>'
            f'<tr><td>出来形管理</td><td class="num">{len(out_d_labels)} 項目</td></tr>'
            f'<tr><td>品質管理</td><td class="num">{len(out_h_labels)} 項目</td></tr>'
            f'<tr><td>撮影箇所</td><td class="num">{len(out_p_labels)} 項目</td></tr>'
            f'</table>'
            f'<div class="note">未確認の要選択行は全候補を自動採用</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

    with col_btn:
        st.markdown('<div style="height:12px"></div>', unsafe_allow_html=True)
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
                    label=f"↓ ダウンロード",
                    data=excel_bytes,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception:
                st.error("Excel 生成中にエラーが発生しました。")
                with st.expander("エラー詳細"):
                    st.code(traceback.format_exc())


# ===========================================================================
# ページルーティング
# ===========================================================================
if st.session_state.page == "db_view":
    _render_db_view()
else:
    _render_matching()
