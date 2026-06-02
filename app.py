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
    page_icon="📋",
    layout="wide",
)

# ===========================================================================
# カスタム CSS
# ===========================================================================
st.markdown("""
<style>
/* ─── フォント: 游ゴシック ─────────────────────────────── */
html, body, [class*="css"], .stMarkdown, .stText,
button, input, select, textarea, th, td {
    font-family: 'Yu Gothic', '游ゴシック', YuGothic,
                 'Hiragino Kaku Gothic ProN', 'Hiragino Sans',
                 Meiryo, sans-serif !important;
}

/* ─── ページ背景 ────────────────────────────────────────── */
.stApp { background-color: #F7FAFD; }

/* ─── メトリクスカード ─────────────────────────────────── */
[data-testid="metric-container"] {
    background: #FFFFFF;
    border: 1px solid #C8DCF0;
    border-top: 3px solid #1565C0;
    border-radius: 6px;
    padding: 14px 18px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}
[data-testid="stMetricLabel"] > div {
    font-size: 0.72rem !important;
    font-weight: 700 !important;
    color: #4A6FA5 !important;
    letter-spacing: 0.06em;
    text-transform: uppercase;
}
[data-testid="stMetricValue"] > div {
    font-size: 1.7rem !important;
    font-weight: 700 !important;
    color: #1F4E79 !important;
}

/* ─── プライマリボタン ─────────────────────────────────── */
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #1565C0 0%, #1976D2 100%);
    color: #FFFFFF;
    border: none;
    border-radius: 6px;
    font-weight: 700;
    letter-spacing: 0.04em;
    box-shadow: 0 2px 8px rgba(21,101,192,0.30);
    transition: box-shadow 0.2s, transform 0.15s;
}
.stButton > button[kind="primary"]:hover {
    box-shadow: 0 4px 14px rgba(21,101,192,0.45);
    transform: translateY(-1px);
}

/* ─── セカンダリボタン ─────────────────────────────────── */
.stButton > button[kind="secondary"] {
    border: 1.5px solid #1565C0;
    color: #1565C0;
    border-radius: 6px;
    font-weight: 600;
}

/* ─── ダウンロードボタン ───────────────────────────────── */
[data-testid="stDownloadButton"] > button {
    background: #E3F2FD !important;
    color: #1565C0 !important;
    border: 2px solid #1565C0 !important;
    border-radius: 6px !important;
    font-weight: 700 !important;
}

/* ─── info/success ボックス ────────────────────────────── */
[data-testid="stInfo"] {
    background: #EBF3FD;
    border-left: 4px solid #1565C0;
    border-radius: 0 6px 6px 0;
    color: #1A2332;
}
[data-testid="stSuccess"] {
    border-radius: 6px;
}

/* ─── サイドバー ───────────────────────────────────────── */
[data-testid="stSidebar"] {
    background: #0C2340;
    border-right: 1px solid #1A3A5C;
}
[data-testid="stSidebar"] * { color: #D0E4F7 !important; }
[data-testid="stSidebar"] h3 {
    color: #7EB8E8 !important;
    font-size: 0.75rem !important;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    font-weight: 700 !important;
}
[data-testid="stSidebar"] hr { border-color: #1E3F60 !important; }
[data-testid="stSidebar"] .stButton > button {
    background: #1565C0 !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 6px !important;
    font-weight: 700 !important;
}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
    background: #0F2D4A !important;
    border: 1.5px dashed #2E6DA4 !important;
    border-radius: 8px !important;
}

/* ─── ラジオボタン ─────────────────────────────────────── */
[data-testid="stRadio"] label { font-size: 0.85rem; }

/* ─── チェックボックス ─────────────────────────────────── */
[data-testid="stCheckbox"] label { font-size: 0.85rem; }

/* ─── divider ──────────────────────────────────────────── */
hr { border-color: #D0DFF0 !important; }
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
        "先に `python build_db.py` を実行してください。"
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


def _sec_header(label: str) -> None:
    """ブルー左ボーダー付きセクションヘッダーを描画する。"""
    st.markdown(
        f'<div style="border-left:4px solid #1565C0; padding:5px 12px; '
        f'background:#EBF3FD; border-radius:0 6px 6px 0; '
        f'font-weight:700; color:#1F4E79; margin:10px 0 6px 0; '
        f'font-size:0.95rem; letter-spacing:0.02em;">{label}</div>',
        unsafe_allow_html=True,
    )


# ===========================================================================
# サイドバー: DB情報 + アップロード + 利用説明
# ===========================================================================
with st.sidebar:
    # ── ロゴ風ヘッダー ──────────────────────────────────────
    st.markdown(
        '<div style="padding:12px 0 8px 0; text-align:center;">'
        '<span style="font-size:1.6rem;">📋</span><br>'
        '<span style="font-size:0.95rem; font-weight:700; '
        'color:#A8D4F5; letter-spacing:0.06em;">施工管理計画</span><br>'
        '<span style="font-size:0.7rem; color:#5A88B0; letter-spacing:0.1em;">'
        'AUTO GENERATOR</span></div>',
        unsafe_allow_html=True,
    )
    st.divider()

    # ── DB情報 ───────────────────────────────────────────────
    st.markdown("### 📚 国交省基準DB")
    st.caption(f"バージョン: **{version_info.get('バージョン', '不明')}**")
    st.caption(f"作成日時: {version_info.get('作成日時', '不明')}")
    st.caption(
        f"出来形管理 {len(kojyo_data['出来形管理'])}行 ／ "
        f"品質管理 {len(kojyo_data['品質管理'])}行"
    )
    st.divider()

    # ── アップロード ─────────────────────────────────────────
    st.markdown("### 📄 数量総括表 PDF")
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
            with st.expander("詳細（開発者向け）"):
                st.code(traceback.format_exc())

    st.divider()

    # ── 利用説明書 ───────────────────────────────────────────
    with st.expander("📖 利用説明書"):
        st.markdown("""
**STEP 1 — PDFをアップロード**
数量総括表のPDFファイルを上の枠にドラッグ＆ドロップし、「解析する」を押します。

---

**STEP 2 — ① 数量総括表を確認**
抽出された工種ツリーが表示されます。行の背景色で状態を確認できます。
- 🟡 **候補あり** — 国交省基準DBにマッチした行
- 🟢 **選択済み** — 候補を確認・選択済みの行
- ⬜ **対象外** — DBマッチなし

---

**STEP 3 — ② DBマッピングを確認**
各工種に対する出来形・品質管理・撮影箇所のマッチ結果を確認します。
フィルタで「候補ありのみ」などに絞り込めます。

---

**STEP 4 — ③ 候補を選択**
② の表から行をクリックすると、その行の詳細候補がチェックボックスで表示されます。
不要な候補のチェックを外して絞り込みます。

---

**STEP 5 — 出力**
右側の「📥 施工管理計画を出力」ボタンを押すとExcelを生成します。
3シート構成（出来形管理 ／ 品質管理 ／ 撮影箇所）でダウンロードできます。

---

**基準DB更新時**
国交省基準改定時は `python build_db.py` を再実行してDBを更新してください。
""")

    st.caption("基準改定時は `python build_db.py` を実行")

# ===========================================================================
# ヘッダー
# ===========================================================================
st.markdown(
    '<div style="background:linear-gradient(135deg,#0C2340 0%,#1565C0 100%); '
    'color:#FFFFFF; padding:16px 24px; border-radius:8px; margin-bottom:4px;">'
    '<span style="font-size:1.25rem; font-weight:700; letter-spacing:0.06em;">'
    '📋 施工管理計画 自動生成アプリ</span><br>'
    '<span style="font-size:0.78rem; color:#A8CCEF;">'
    '数量総括表 PDF → 国交省基準 DBマッピング → 施工管理計画 Excel 出力</span>'
    '</div>',
    unsafe_allow_html=True,
)

if st.session_state.suryo_info:
    name = st.session_state.suryo_info.get("工事名", "")
    st.success(f"📄 **{name or '（工事名不明）'}**　 読込済み")

# ===========================================================================
# 未解析時
# ===========================================================================
if st.session_state.df_match is None:
    st.divider()
    st.info("⬅️ サイドバーから数量総括表PDFをアップロードして「解析する」を押してください。")
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
mc[0].metric("総行数",             n_total)
mc[1].metric("✅ 選択済み",        n_selected)
mc[2].metric("⚠️ 候補あり（未選択）", n_pending)
mc[3].metric("➖ 対象外",          n_out)
st.divider()

STATUS_BG = {"選択済み": "#E8F5E9", "候補あり": "#FFF9C4", "対象外": "#F5F5F5"}


def _row_style(statuses: dict):
    def _style(row):
        bg = STATUS_BG.get(statuses.get(row.name, ""), "")
        return [f"background-color: {bg}" if bg else "" for _ in row]
    return _style


# ===========================================================================
# ① 数量総括表ツリー（フルワイド）
# ===========================================================================
_sec_header("① 数量総括表（元データ）")
st.caption("※「直接工事費」以下は除外済み　　行をクリック → 下部③に詳細・候補が表示されます")

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

st.caption("🟢 選択済み　　🟡 候補あり（未選択）　　⬜ 対象外")

# ===========================================================================
# ② DBマッピング一覧（フルワイド）
# ===========================================================================
_sec_header("② DBマッピング一覧")

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
# ③ + ④ 下部 2カラム: 候補選択（左）+ 出力（右）
# ===========================================================================
col_detail, col_out = st.columns([3, 2])

# ── ③ 左: 選択行の詳細・候補選択 ─────────────────────────────────────
with col_detail:
    _sec_header("③ 候補選択")

    sel_idx = st.session_state.selected_idx

    if sel_idx is not None and 0 <= sel_idx < len(df_raw):
        sel  = df_raw.iloc[sel_idx]
        ckey = _chain_key(sel)

        chain = " › ".join(sel[c] for c in SURYO_LEVEL_COLS if sel.get(c, ""))
        st.info(f"**{sel['_name']}**　　{chain}")

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
                    st.markdown("🔵 **出来形管理**")
                    grouped_d = _group_items(items_d)
                    for kojyo, sub_items in grouped_d.items():
                        if len(grouped_d) > 1:
                            st.caption(f"■ {kojyo}")
                        for full_label, display_label in sub_items:
                            chk_key = f"chk_d_{sel_idx}_{items_d.index(full_label)}"
                            if st.checkbox(display_label, value=True, key=chk_key):
                                sel_d.append(full_label)
                else:
                    st.caption("🔵 出来形：該当なし")

            with c_h:
                sel_h = []
                if items_h:
                    st.markdown("🟢 **品質管理**")
                    grouped_h = _group_items(items_h)
                    for kojyo, sub_items in grouped_h.items():
                        if len(grouped_h) > 1:
                            st.caption(f"■ {kojyo}")
                        for full_label, display_label in sub_items:
                            chk_key = f"chk_h_{sel_idx}_{items_h.index(full_label)}"
                            if st.checkbox(display_label, value=True, key=chk_key):
                                sel_h.append(full_label)
                else:
                    st.caption("🟢 品質管理：該当なし")

            with c_p:
                sel_p = []
                if items_p:
                    st.markdown("🟠 **撮影箇所**")
                    grouped_p = _group_items(items_p)
                    for kojyo, sub_items in grouped_p.items():
                        if len(grouped_p) > 1:
                            st.caption(f"■ {kojyo}")
                        for full_label, display_label in sub_items:
                            chk_key = f"chk_p_{sel_idx}_{items_p.index(full_label)}"
                            if st.checkbox(display_label, value=True, key=chk_key):
                                sel_p.append(full_label)
                else:
                    st.caption("🟠 撮影箇所：該当なし")

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
                st.caption(f"DB目次: {bc}")

    else:
        st.info("↑ 上の表から行をクリックすると候補が表示されます")

# ── ④ 右: 出力 ─────────────────────────────────────────────────────────
with col_out:
    _sec_header("④ 施工管理計画を出力")

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

            labels_d = saved["出来形"]             if saved is not None else all_d
            labels_h = saved["品質管理"]           if saved is not None else all_h
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

    # 出力サマリ
    st.markdown(
        f'<div style="background:#EBF3FD; border-radius:6px; '
        f'padding:10px 14px; font-size:0.85rem; color:#1F4E79; margin-bottom:10px;">'
        f'出来形管理　<strong>{len(out_d_labels)}</strong> 項目<br>'
        f'品質管理　　<strong>{len(out_h_labels)}</strong> 項目<br>'
        f'撮影箇所　　<strong>{len(out_p_labels)}</strong> 項目<br>'
        f'<span style="font-size:0.75rem; color:#4A6FA5;">※ 未確認行は全候補を採用</span>'
        f'</div>',
        unsafe_allow_html=True,
    )

    if st.button(
        "📥 施工管理計画を出力",
        type="primary",
        use_container_width=True,
        disabled=not (out_d_labels or out_h_labels or out_p_labels),
    ):
        try:
            with st.spinner("Excel生成中..."):
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
                label=f"⬇️ {fname} をダウンロード",
                data=excel_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        except Exception:
            st.error("Excel生成中にエラーが発生しました。")
            with st.expander("詳細（開発者向け）"):
                st.code(traceback.format_exc())
