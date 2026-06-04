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
# CSS
# ===========================================================================
st.markdown("""
<style>
/* ─── 全体 ───────────────────────────────────────────────────── */
html, body, [class*="css"], .stMarkdown, .stText,
button, input, select, textarea, th, td {
    font-family: 'Yu Gothic','游ゴシック',YuGothic,
                 'Hiragino Kaku Gothic ProN','Hiragino Sans',
                 Meiryo,sans-serif !important;
}
* { text-shadow: none !important; }
.stApp,
[data-testid="stAppViewContainer"],
.main { background: #F4F6F9 !important; }
[data-testid="stAppViewContainer"]::before,
[data-testid="stAppViewContainer"]::after,
[data-testid="stHeader"],
[data-testid="stHeader"]::before,
[data-testid="stHeader"]::after {
    background-image: none !important;
    filter: none !important; backdrop-filter: none !important;
}
/* padding 詰め — スクロールが出にくくなる */
.block-container {
    padding-top: 0 !important;
    padding-bottom: 0.5rem !important;
    max-width: 100% !important;
}

/* ─── サイドバー ─────────────────────────────────────────────── */
[data-testid="stSidebar"] {
    background: #FFFFFF !important;
    border-right: 1px solid #E2E6EA !important;
    min-width: 230px !important;
    max-width: 230px !important;
}
[data-testid="stSidebar"] * { color: #333 !important; }
[data-testid="stSidebar"] hr { border-color: #EEEEEE !important; }
[data-testid="stSidebar"] h3 {
    color: #AAAAAA !important; font-size: 0.62rem !important;
    letter-spacing: .18em !important; text-transform: uppercase !important;
    font-weight: 700 !important; margin: 8px 0 5px !important;
}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
    background: #F8F9FA !important; border: 1px dashed #CCC !important;
    border-radius: 5px !important;
}
[data-testid="stSidebar"] code {
    background: #F4F6F8 !important; color: #1565C0 !important;
    border: 1px solid #D0DCF0 !important; padding: 1px 5px !important;
    border-radius: 3px !important; font-size: .78rem !important;
}
[data-testid="stSidebar"] [data-testid="stExpander"] {
    background: #F8F9FA !important; border: 1px solid #EEE !important;
    border-radius: 5px !important;
}
[data-testid="stSidebar"] [data-testid="stExpander"] summary {
    font-size: .83rem !important; color: #555 !important;
}
[data-testid="stSidebar"] [data-testid="stExpander"] .stMarkdown p,
[data-testid="stSidebar"] [data-testid="stExpander"] .stMarkdown li {
    font-size: .81rem !important; line-height: 1.8 !important; color: #555 !important;
}

/* サイドバーボタン */
[data-testid="stSidebar"] .stButton > button {
    width: 100% !important; border-radius: 6px !important;
    font-size: .84rem !important; font-weight: 500 !important;
    padding: 7px 10px !important; text-align: left !important;
    border: none !important; transition: background .12s;
}
[data-testid="stSidebar"] .stButton > button[kind="primary"] {
    background: #EEF4FF !important; color: #1565C0 !important;
    font-weight: 700 !important;
}
[data-testid="stSidebar"] .stButton > button[kind="secondary"] {
    background: transparent !important; color: #444 !important;
}
[data-testid="stSidebar"] .stButton > button[kind="secondary"]:hover {
    background: #F0F2F6 !important;
}

/* ─── メインボタン ───────────────────────────────────────────── */
.stButton > button[kind="primary"] {
    background: #1565C0; color: #FFF;
    border: none; border-radius: 5px;
    font-weight: 600; letter-spacing: .03em;
}
.stButton > button[kind="primary"]:hover { background: #0D47A1; }
.stButton > button[kind="primary"]:disabled {
    background: #CCC !important; color: #999 !important;
}

/* ─── ダウンロード ───────────────────────────────────────────── */
[data-testid="stDownloadButton"] > button {
    background: #FFF !important; color: #1565C0 !important;
    border: 1.5px solid #1565C0 !important; border-radius: 5px !important;
    font-weight: 600 !important;
}
[data-testid="stDownloadButton"] > button:hover { background: #E3F2FD !important; }

/* ─── info / warning ─────────────────────────────────────────── */
[data-testid="stInfo"] {
    background: #E8F0FE; border-left: 3px solid #1565C0;
    border-radius: 4px; color: #1A2B3C;
}

/* ─── ラジオ ─────────────────────────────────────────────────── */
[data-testid="stRadio"] label { font-size: .84rem; }
[data-testid="stCheckbox"] label { font-size: .84rem; }

hr { border-color: #E2E6EA !important; }

/* ─── トップバー（アプリヘッダー） ──────────────────────────── */
.topbar {
    background: #FFFFFF;
    border-bottom: 1px solid #E2E6EA;
    padding: 10px 20px;
    margin: 0 -1rem 14px -1rem;
    display: flex; align-items: center;
    justify-content: space-between;
    box-shadow: 0 1px 4px rgba(0,0,0,.04);
}
.topbar-left {
    display: flex; align-items: center; gap: 10px;
}
.topbar-icon {
    width: 30px; height: 30px; background: #1565C0; border-radius: 6px;
    display: flex; align-items: center; justify-content: center;
    color: #FFF; font-size: .70rem; font-weight: 800; flex-shrink: 0;
}
.topbar-title { font-size: .96rem; font-weight: 800; color: #1A2332; }
.topbar-sub { font-size: .68rem; color: #AAAAAA; margin-top: 1px; }
.stepbar { display: flex; align-items: center; gap: 2px; font-size: .70rem; }
.step-chip {
    display: flex; align-items: center; gap: 3px;
    padding: 3px 8px; border-radius: 12px; white-space: nowrap;
}
.step-chip.done   { color: #888; }
.step-chip.active {
    background: #EEF4FF; color: #1565C0;
    font-weight: 700; border: 1px solid #C5D8F8;
}
.step-chip.pending { color: #CCCCCC; }
.step-sep { color: #DDD; padding: 0 1px; }

/* ─── メトリクス（コンパクト） ──────────────────────────────── */
.metrics-row {
    display: flex; gap: 10px; margin-bottom: 10px;
}
.m-card {
    flex: 1; background: #FFFFFF; border: 1px solid #E2E6EA;
    border-radius: 7px; padding: 10px 16px;
    box-shadow: 0 1px 3px rgba(0,0,0,.04);
    display: flex; align-items: center; gap: 12px;
}
.m-card-val { font-size: 2rem; font-weight: 800; line-height: 1; }
.m-card-label { font-size: .70rem; font-weight: 600; color: #888;
                text-transform: uppercase; letter-spacing: .07em; }
.m-card.kakutei   .m-card-val { color: #1B6E2A; }
.m-card.yosentaku .m-card-val { color: #B45309; }
.m-card.mimatch   .m-card-val { color: #9E9E9E; }
.m-card.kakutei   { border-left: 3px solid #34A853; }
.m-card.yosentaku { border-left: 3px solid #F59E0B; }
.m-card.mimatch   { border-left: 3px solid #CCC; }

/* ─── 候補パネル ─────────────────────────────────────────────── */
.cand-panel {
    background: #FFFBEB; border: 1.5px solid #F59E0B;
    border-radius: 8px; padding: 14px 18px; margin-top: 10px;
}
.cand-hdr {
    font-size: .86rem; font-weight: 700; color: #92400E;
    margin-bottom: 12px; display: flex; align-items: center; gap: 6px;
}
.cand-card {
    background: #FFF; border: 1px solid #E2E6EA; border-radius: 6px; padding: 12px 14px;
}
.cand-card.sel { border-color: #1565C0; background: #EEF6FF; }
.cand-card-title { font-size: .86rem; font-weight: 700; color: #1A2332; margin-bottom: 6px; }
.cand-card-body  { font-size: .78rem; color: #555; line-height: 1.85; }
.diff-chip {
    display: inline-block; background: #DBEAFE; color: #1D4ED8;
    border-radius: 3px; padding: 0 5px; font-size: .74rem; font-weight: 600;
}
.cand-foot { font-size: .71rem; color: #AAA; margin-top: 8px; }

/* ─── 凡例 ──────────────────────────────────────────────────── */
.legend {
    display: flex; gap: 14px; font-size: .73rem; color: #666;
    align-items: center; margin-top: 5px;
}
.ldot {
    width: 9px; height: 9px; border-radius: 2px;
    display: inline-block; margin-right: 3px; vertical-align: middle;
}

/* ─── 出力サマリー（サイドバー内） ──────────────────────────── */
.out-summary {
    background: #F0F6FF; border: 1px solid #C5D8F8;
    border-radius: 6px; padding: 10px 12px; margin-bottom: 8px;
}
.out-summary table { width: 100%; border-collapse: collapse; font-size: .82rem; }
.out-summary td { padding: 3px 0; color: #444; }
.out-summary td.n { text-align: right; font-weight: 700; color: #1565C0; }
.out-summary .note {
    font-size: .70rem; color: #999; margin-top: 6px;
    border-top: 1px solid #DCE8F8; padding-top: 5px;
}

/* ─── サブラベル ─────────────────────────────────────────────── */
.sublabel {
    border-bottom: 1px solid #1565C0; padding-bottom: 2px;
    font-weight: 700; font-size: .81rem; color: #333; margin-bottom: 5px;
}
</style>
""", unsafe_allow_html=True)

# ===========================================================================
# DB 読み込み
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
    ver = df_ver.iloc[0].to_dict() if not df_ver.empty else {}
    return data, ver


kojyo_data, version_info = load_kojyo_db()
if kojyo_data is None:
    st.error("国交省基準DBが見つかりません。build_db.py を実行してください。")
    st.stop()

unique_kojyo = get_unique_kojyo(kojyo_data)

# ===========================================================================
# セッション初期化
# ===========================================================================
for _k in ["suryo_info", "df_match", "selected_idx"]:
    if _k not in st.session_state:
        st.session_state[_k] = None
if "row_selections"  not in st.session_state: st.session_state["row_selections"]  = {}
if "page"            not in st.session_state: st.session_state["page"]            = "matching"
if "excel_cache"     not in st.session_state: st.session_state["excel_cache"]     = None
if "excel_fname"     not in st.session_state: st.session_state["excel_fname"]     = None

# ===========================================================================
# ヘルパー
# ===========================================================================
def _chain_key(row) -> tuple:
    return tuple(str(row.get(c, "")) for c in SURYO_LEVEL_COLS)

def _group_items(items):
    g = {}
    for item in items:
        parts = [p.strip() for p in item.split(" / ")]
        kojyo = parts[0]
        sub   = " / ".join(parts[1:]) if len(parts) > 1 else parts[0]
        g.setdefault(kojyo, []).append((item, sub))
    return g

def _deepest_name(row) -> str:
    for col in reversed(SURYO_LEVEL_COLS):
        v = row.get(col, "")
        if v: return v
    return ""

def _depth(row) -> int:
    d = 0
    for col in SURYO_LEVEL_COLS:
        if row.get(col, ""): d = SURYO_LEVEL_COLS.index(col)
    return d

def _sublabel(label: str):
    st.markdown(f'<div class="sublabel">{label}</div>', unsafe_allow_html=True)

# ─── ステータス ────────────────────────────────────────────────
def _calc_status(row) -> str:
    has = (bool(str(row.get("出来形マッチ","")).strip()) or
           bool(str(row.get("品質管理マッチ","")).strip()) or
           bool(str(row.get("撮影箇所マッチ","")).strip()))
    if not has: return "未マッチ"
    if _chain_key(row) in st.session_state.row_selections: return "確定"
    return "要選択"

STATUS_BG = {"確定": "#F0FDF4", "要選択": "#FFFBEB", "未マッチ": "#F9FAFB"}

# ─── DB ルックアップ ────────────────────────────────────────────
_DISP_D = ["測定項目","規格値","管理基準値","測定頻度","摘要"]
_DISP_H = ["試験項目","試験方法","試験基準","摘要"]
_DISP_P = ["撮影箇所","提出頻度","摘要"]

def _lookup_db(label: str, db_key: str) -> dict:
    parts = [p.strip() for p in label.split(" / ")]
    df_db = kojyo_data[db_key]
    lvl   = ["工種","種別","細別"]
    mask  = pd.Series([True]*len(df_db), index=df_db.index)
    for i, p in enumerate(parts):
        c = lvl[i] if i < len(lvl) else None
        if c and c in df_db.columns:
            mask = mask & (df_db[c] == p)
    rows = df_db[mask]
    return rows.iloc[0].to_dict() if not rows.empty else {}

def _card_html(rdict: dict, cols: list, diff_keys: set) -> str:
    html = ""
    for c in cols:
        v = str(rdict.get(c,"")).strip()
        if not v: continue
        if c in diff_keys:
            html += f'<div><span class="diff-chip">{v}</span></div>'
        else:
            html += f'<div>{c}：{v}</div>'
    return html or '<div style="color:#CCC">情報なし</div>'

def _diff_cols(rows: list, cols: list) -> set:
    if len(rows) < 2: return set()
    return {c for c in cols if len({str(r.get(c,"")).strip() for r in rows}) > 1}

# ─── 出力ラベル収集 ─────────────────────────────────────────────
def _collect_labels(df_raw):
    out_d, out_h, out_p = [], [], []
    seen_d, seen_h, seen_p = set(), set(), set()
    for _, row in df_raw[df_raw["状態"].isin(["確定","要選択"])].iterrows():
        ckey  = _chain_key(row)
        saved = st.session_state.row_selections.get(ckey)
        all_d = [x.strip() for x in str(row.get("出来形マッチ","")).split("\n") if x.strip()]
        all_h = [x.strip() for x in str(row.get("品質管理マッチ","")).split("\n") if x.strip()]
        all_p = [x.strip() for x in str(row.get("撮影箇所マッチ","")).split("\n") if x.strip()]
        labels_d = saved["出来形"]              if saved else all_d
        labels_h = saved["品質管理"]            if saved else all_h
        labels_p = saved.get("撮影箇所",all_p)  if saved else all_p
        for lbl in labels_d:
            if lbl and lbl not in seen_d: out_d.append(lbl); seen_d.add(lbl)
        for lbl in labels_h:
            if lbl and lbl not in seen_h: out_h.append(lbl); seen_h.add(lbl)
        for lbl in labels_p:
            if lbl and lbl not in seen_p: out_p.append(lbl); seen_p.add(lbl)
    return out_d, out_h, out_p

# ===========================================================================
# ステップバー
# ===========================================================================
def _stepbar(current: int):
    steps = [("①","基準DB"),("②","取込"),("③","構造化"),
             ("④","マッチング"),("⑤","候補選択"),("⑥","出力")]
    chips = ""
    for i,(num,lbl) in enumerate(steps):
        n = i+1
        cls = "done" if n < current else ("active" if n == current else "pending")
        chips += f'<span class="step-chip {cls}">{num} {lbl}</span>'
        if i < len(steps)-1:
            chips += '<span class="step-sep">›</span>'

    st.markdown(
        f'<div class="topbar">'
        f'<div class="topbar-left">'
        f'<div class="topbar-icon">施</div>'
        f'<div><div class="topbar-title">施工管理計画 自動生成</div>'
        f'<div class="topbar-sub">数量総括表PDF → 国交省基準DBマッチング → Excel出力</div></div>'
        f'</div>'
        f'<div class="stepbar">{chips}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

def _compute_step() -> int:
    if st.session_state.df_match is None:
        return 1 if st.session_state.suryo_info is None else 2
    n_yo = sum(1 for _,r in st.session_state.df_match.iterrows()
               if _calc_status(r) == "要選択")
    return 5 if n_yo > 0 else 6

# ===========================================================================
# サイドバー
# ===========================================================================
with st.sidebar:
    # ロゴ
    st.markdown(
        '<div style="padding:14px 4px 12px;border-bottom:1px solid #EEE;margin-bottom:12px;">'
        '<div style="font-size:.82rem;font-weight:800;color:#1A2332;">施工管理計画</div>'
        '<div style="font-size:.68rem;color:#AAA;margin-top:1px;">自動生成システム</div>'
        '</div>', unsafe_allow_html=True,
    )

    # ── ナビ ───────────────────────────────────────────────────
    page = st.session_state.get("page","matching")
    if st.button("📊  マッチング", use_container_width=True,
                 type="primary" if page=="matching" else "secondary", key="nav_m"):
        st.session_state.page = "matching"; st.rerun()
    if st.button("🗄  基準DB確認", use_container_width=True,
                 type="primary" if page=="db_view" else "secondary", key="nav_d"):
        st.session_state.page = "db_view"; st.rerun()

    st.divider()

    # ── PDF アップロード ───────────────────────────────────────
    st.markdown("### PDF アップロード")
    uploaded = st.file_uploader("数量総括表PDF", type="pdf", label_visibility="collapsed")
    if st.button("解析する", type="primary", disabled=not uploaded, use_container_width=True):
        try:
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
                f.write(uploaded.read()); path_tmp = f.name
            with st.spinner("解析・マッチング中..."):
                si = extract_suryo(path_tmp)
                dm = build_match_detail(
                    si["工種階層"],
                    kojyo_data["出来形管理"],
                    kojyo_data["品質管理"],
                    kojyo_data["撮影箇所"],
                )
            for k in list(st.session_state.keys()):
                if k.startswith(("chk_d_","chk_h_","chk_p_")):
                    del st.session_state[k]
            st.session_state.suryo_info     = si
            st.session_state.df_match       = dm
            st.session_state.selected_idx   = None
            st.session_state.row_selections = {}
            st.session_state.excel_cache    = None
            st.session_state.excel_fname    = None
            st.session_state.page           = "matching"
            st.rerun()
        except Exception:
            st.error("解析エラー")
            with st.expander("詳細"): st.code(traceback.format_exc())

    st.divider()

    # ── 出力セクション（データあり時） ────────────────────────
    if st.session_state.df_match is not None:
        df_tmp = st.session_state.df_match.copy()
        df_tmp["_s"] = df_tmp.apply(_calc_status, axis=1)
        df_tmp["状態"] = df_tmp["_s"]
        df_tmp.insert(0,"No",range(1,len(df_tmp)+1))

        out_d_l, out_h_l, out_p_l = _collect_labels(df_tmp)

        st.markdown("### 出力")
        st.markdown(
            f'<div class="out-summary">'
            f'<table>'
            f'<tr><td>出来形管理</td><td class="n">{len(out_d_l)}</td></tr>'
            f'<tr><td>品質管理</td><td class="n">{len(out_h_l)}</td></tr>'
            f'<tr><td>撮影箇所</td><td class="n">{len(out_p_l)}</td></tr>'
            f'</table>'
            f'<div class="note">未確認の要選択は全候補を自動採用</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

        can_out = bool(out_d_l or out_h_l or out_p_l)
        if st.button("施工管理計画を出力", type="primary",
                     use_container_width=True, disabled=not can_out):
            try:
                si = st.session_state.suryo_info
                with st.spinner("Excel生成中..."):
                    filtered = filter_by_row_labels(kojyo_data, out_d_l, out_h_l, out_p_l)
                    dmap = build_suryo_match_map(
                        si["工種リスト"],
                        list({l.split(" / ")[0].strip() for l in out_d_l})
                    )
                    excel_bytes = write_excel(filtered, 工事名=si["工事名"], dekigata_kojyo_map=dmap)
                safe = re.sub(r'[\\/:*?"<>|　 ]','_', si["工事名"])
                st.session_state.excel_cache = excel_bytes
                st.session_state.excel_fname = f"施工管理計画_{safe}.xlsx" if safe else "施工管理計画.xlsx"
                st.rerun()
            except Exception:
                st.error("生成エラー")
                with st.expander("詳細"): st.code(traceback.format_exc())

        if st.session_state.excel_cache:
            st.download_button(
                "⬇ ダウンロード",
                data=st.session_state.excel_cache,
                file_name=st.session_state.excel_fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        st.divider()

    # ── DB 情報 ────────────────────────────────────────────────
    st.markdown("### 国交省基準 DB")
    st.caption(f"Ver. {version_info.get('バージョン','不明')}　"
               f"{version_info.get('作成日時','')}")
    st.caption(f"出来形 {len(kojyo_data['出来形管理'])} 行　"
               f"品質 {len(kojyo_data['品質管理'])} 行")

    st.divider()

    # ── リセット + 使い方 ──────────────────────────────────────
    if st.button("↺  リセット", use_container_width=True, key="btn_reset"):
        for k in list(st.session_state.keys()):
            if k != "page": del st.session_state[k]
        st.session_state.suryo_info = st.session_state.df_match = st.session_state.selected_idx = None
        st.session_state.row_selections = {}
        st.session_state.excel_cache = st.session_state.excel_fname = None
        st.rerun()

    with st.expander("？  使い方"):
        st.markdown("""
**①** 数量総括表PDFをアップロード
**②** 「解析する」でマッチング実行
**③** 表の行をクリック → 候補確認
**④** 「施工管理計画を出力」→ダウンロード

---
基準DB更新時は `build_db.py` を再実行
""")

# ===========================================================================
# ステップバー（トップバー）
# ===========================================================================
_stepbar(_compute_step())

# ===========================================================================
# 基準DB確認ページ
# ===========================================================================
def _render_db_view():
    tab_d, tab_h, tab_p = st.tabs(["出来形管理","品質管理","撮影箇所"])
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

    si     = st.session_state.suryo_info
    df_raw = st.session_state.df_match.copy()
    df_raw["状態"]   = df_raw.apply(_calc_status, axis=1)
    df_raw["_name"]  = df_raw.apply(_deepest_name, axis=1)
    df_raw["_depth"] = df_raw.apply(_depth, axis=1)
    df_raw.insert(0,"No",range(1,len(df_raw)+1))

    n_kaku = int((df_raw["状態"]=="確定").sum())
    n_yo   = int((df_raw["状態"]=="要選択").sum())
    n_mi   = int((df_raw["状態"]=="未マッチ").sum())

    # 工事名
    if si:
        name = si.get("工事名","")
        if name:
            st.markdown(
                f'<div style="background:#E8F0FE;border:1px solid #C5D8F8;'
                f'border-radius:5px;padding:6px 14px;margin-bottom:10px;'
                f'font-size:.85rem;color:#1565C0;">'
                f'<strong>{name}</strong>　読込済み</div>',
                unsafe_allow_html=True,
            )

    # ── メトリクス ───────────────────────────────────────────
    st.markdown(
        f'<div class="metrics-row">'
        f'<div class="m-card kakutei">'
        f'<div class="m-card-val">{n_kaku}</div>'
        f'<div class="m-card-label">確定</div>'
        f'</div>'
        f'<div class="m-card yosentaku">'
        f'<div class="m-card-val">{n_yo}</div>'
        f'<div class="m-card-label">要選択</div>'
        f'</div>'
        f'<div class="m-card mimatch">'
        f'<div class="m-card-val">{n_mi}</div>'
        f'<div class="m-card-label">未マッチ</div>'
        f'</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # ── フィルタ ─────────────────────────────────────────────
    filter_opt = st.radio(
        "filter",
        ["すべて","要選択のみ","確定のみ","未マッチのみ"],
        horizontal=True, label_visibility="collapsed",
    )
    FM = {"確定のみ":"確定","要選択のみ":"要選択","未マッチのみ":"未マッチ"}
    df_v = (df_raw[df_raw["状態"]==FM[filter_opt]].copy()
            if filter_opt in FM else df_raw.copy())

    # ── テーブル ─────────────────────────────────────────────
    sel_idx = st.session_state.selected_idx
    has_sel = sel_idx is not None and 0 <= sel_idx < len(df_raw)

    # 行選択時はテーブルを縮めて候補パネルのスペースを確保
    tbl_h = 240 if has_sel else 450

    def _fmt(row) -> str:
        d = str(row.get("出来形マッチ","")).strip()
        h = str(row.get("品質管理マッチ","")).strip()
        p = str(row.get("撮影箇所マッチ","")).strip()
        if not (d or h or p): return "—"
        dl = [x for x in d.split("\n") if x.strip()]
        cats = (["出来形"] if d else []) + (["品質"] if h else []) + (["撮影"] if p else [])
        if len(dl) >= 2:
            return f"候補{len(dl)}件（工法で分岐）"
        return "・".join(cats)

    sts_idx = {i: row["状態"] for i,(_, row) in enumerate(df_v.iterrows())}

    df_tbl = pd.DataFrame({
        "工種・項目":   ["　"*row["_depth"]+row["_name"] for _,row in df_v.iterrows()],
        "マッチした基準": [_fmt(row) for _,row in df_v.iterrows()],
        "状態":        [row["状態"] for _,row in df_v.iterrows()],
    })

    def _row_style(sts):
        def _s(row):
            bg = STATUS_BG.get(sts.get(row.name,""),"")
            return [f"background-color:{bg}" if bg else "" for _ in row]
        return _s

    ev = st.dataframe(
        df_tbl.style.apply(_row_style(sts_idx), axis=1),
        use_container_width=True, height=tbl_h, hide_index=True,
        selection_mode="single-row", on_select="rerun",
        column_config={
            "工種・項目":    st.column_config.TextColumn(width="large"),
            "マッチした基準": st.column_config.TextColumn(width="large"),
            "状態":         st.column_config.TextColumn(width="small"),
        },
    )
    if ev.selection.rows:
        no = int(df_v.iloc[ev.selection.rows[0]]["No"])
        st.session_state.selected_idx = no - 1
        st.rerun()

    # 凡例
    st.markdown(
        '<div class="legend">'
        '<span><span class="ldot" style="background:#1B6E2A"></span>確定</span>'
        '<span><span class="ldot" style="background:#B45309"></span>要選択</span>'
        '<span><span class="ldot" style="background:#9E9E9E"></span>未マッチ</span>'
        '</div>',
        unsafe_allow_html=True,
    )

    # ── 候補パネル ───────────────────────────────────────────
    if not has_sel:
        st.markdown(
            '<div style="margin-top:10px;padding:18px 20px;background:#FFFFFF;'
            'border:1px dashed #DDD;border-radius:8px;text-align:center;'
            'color:#AAAAAA;font-size:.84rem;">'
            '行をクリックすると候補がここに展開されます</div>',
            unsafe_allow_html=True,
        )
        return

    sel  = df_raw.iloc[sel_idx]
    ckey = _chain_key(sel)

    items_d = [x.strip() for x in str(sel.get("出来形マッチ","")).split("\n") if x.strip()]
    items_h = [x.strip() for x in str(sel.get("品質管理マッチ","")).split("\n") if x.strip()]
    items_p = [x.strip() for x in str(sel.get("撮影箇所マッチ","")).split("\n") if x.strip()]

    if not items_d and not items_h and not items_p:
        st.markdown(
            f'<div style="margin-top:10px;padding:14px 18px;background:#F9FAFB;'
            f'border:1px solid #DDD;border-radius:8px;font-size:.85rem;color:#888;">'
            f'「{sel["_name"]}」はDBマッチなし（未マッチ）</div>',
            unsafe_allow_html=True,
        )
        return

    saved     = st.session_state.row_selections.get(ckey)
    n_remain  = n_yo
    chain     = " › ".join(sel[c] for c in SURYO_LEVEL_COLS if sel.get(c,""))

    st.markdown(
        f'<div class="cand-panel">'
        f'<div class="cand-hdr">⚠ {sel["_name"]}'
        f'<span style="font-weight:400;font-size:.78rem;color:#B45309;margin-left:6px;">'
        f'{chain}</span>'
        f'<span style="margin-left:auto;font-weight:400;font-size:.76rem;">'
        f'残り {n_remain} 件</span>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # 出来形：比較カード
    new_sel_d = []
    if len(items_d) >= 2:
        db_rows_d  = [_lookup_db(lbl,"出来形管理") for lbl in items_d[:4]]
        diff_d     = _diff_cols(db_rows_d, _DISP_D)
        cur_sel_d  = saved["出来形"] if saved else items_d

        n_cards = min(len(items_d), 4)
        cols = st.columns(n_cards)
        for i,(col,lbl) in enumerate(zip(cols, items_d[:n_cards])):
            with col:
                parts = [p.strip() for p in lbl.split(" / ")]
                ctitle = " / ".join(parts[1:]) if len(parts)>1 else parts[0]
                is_sel = lbl in cur_sel_d
                brd = "border:1.5px solid #1565C0;background:#EEF6FF;" if is_sel else ""
                body = _card_html(db_rows_d[i], _DISP_D, diff_d)
                st.markdown(
                    f'<div class="cand-card" style="{brd}">'
                    f'<div class="cand-card-title">候補{chr(65+i)}：{ctitle}</div>'
                    f'<div class="cand-card-body">{body}</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )
                if st.checkbox("採用", value=is_sel, key=f"chk_d_{sel_idx}_{i}"):
                    new_sel_d.append(lbl)

        if diff_d:
            st.markdown(
                f'<div class="cand-foot">ⓘ 差分（{"・".join(sorted(diff_d))}）をハイライト表示</div>',
                unsafe_allow_html=True,
            )
    else:
        new_sel_d = items_d  # 単一候補はそのまま採用

    st.markdown('</div>', unsafe_allow_html=True)  # /cand-panel

    # 品質・撮影は expander
    new_sel_h, new_sel_p = items_h, items_p
    if items_h or items_p:
        with st.expander("品質管理・撮影箇所の候補を調整"):
            ch, cp = st.columns(2)
            with ch:
                new_sel_h = []
                if items_h:
                    _sublabel("品質管理")
                    gh = _group_items(items_h)
                    for ky, sub in gh.items():
                        if len(gh)>1: st.caption(ky)
                        for fl, dl in sub:
                            if st.checkbox(dl, value=True,
                                           key=f"chk_h_{sel_idx}_{items_h.index(fl)}"):
                                new_sel_h.append(fl)
                else:
                    st.caption("該当なし")
            with cp:
                new_sel_p = []
                if items_p:
                    _sublabel("撮影箇所")
                    gp = _group_items(items_p)
                    for ky, sub in gp.items():
                        if len(gp)>1: st.caption(ky)
                        for fl, dl in sub:
                            if st.checkbox(dl, value=True,
                                           key=f"chk_p_{sel_idx}_{items_p.index(fl)}"):
                                new_sel_p.append(fl)
                else:
                    st.caption("該当なし")

    st.session_state.row_selections[ckey] = {
        "出来形":  new_sel_d,
        "品質管理": new_sel_h,
        "撮影箇所": new_sel_p,
    }

    # DB目次
    if items_d:
        fd = items_d[0].split(" / ")[0]
        dr = kojyo_data["出来形管理"][kojyo_data["出来形管理"]["工種"]==fd]
        if not dr.empty:
            r = dr.iloc[0]
            bc = " › ".join(x for x in [r.get("編",""),r.get("章",""),r.get("節",""),fd] if x)
            st.caption(f"DB 目次：{bc}")


# ===========================================================================
# ルーティング
# ===========================================================================
if st.session_state.page == "db_view":
    _render_db_view()
else:
    _render_matching()
