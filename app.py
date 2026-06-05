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

st.set_page_config(page_title="施工管理計画 自動生成", page_icon="☐", layout="wide")

# ===========================================================================
# CSS
# ===========================================================================
st.markdown("""
<style>
/* ── フォント ─────────────────────────────────────────────── */
html,body,[class*="css"],.stMarkdown,.stText,
button,input,select,textarea,th,td{
    font-family:'Yu Gothic','游ゴシック',YuGothic,
               'Hiragino Kaku Gothic ProN','Hiragino Sans',
               Meiryo,sans-serif!important;
}
*{text-shadow:none!important;}

/* ── ページ背景 ───────────────────────────────────────────── */
.stApp,[data-testid="stAppViewContainer"],.main{background:#F4F2EE!important;}
[data-testid="stAppViewContainer"]::before,
[data-testid="stAppViewContainer"]::after,
[data-testid="stHeader"],[data-testid="stHeader"]::before,
[data-testid="stHeader"]::after{
    background-image:none!important;filter:none!important;
    backdrop-filter:none!important;
}
/* ヘッダー(~2.5rem)の下に出るよう余白を確保 */
.block-container{
    padding-top:3.5rem!important;
    padding-bottom:1.5rem!important;
    max-width:100%!important;
}

/* ── サイドバー（チャコール） ─────────────────────────────── */
[data-testid="stSidebar"]{
    background:#2B2A28!important;
    border-right:1px solid #3D3C39!important;
    min-width:220px!important; max-width:220px!important;
}
[data-testid="stSidebar"] *{color:#EDEBE6!important;}
[data-testid="stSidebar"] strong,[data-testid="stSidebar"] b{color:#FFFFFF!important;}
[data-testid="stSidebar"] h3{
    color:#9A9893!important; font-size:.60rem!important;
    letter-spacing:.22em!important; text-transform:uppercase!important;
    font-weight:700!important; margin:12px 0 4px 16px!important;
}
[data-testid="stSidebar"] hr{border-color:#3D3C39!important;}
[data-testid="stSidebar"] code{
    background:#1C1B19!important; color:#EDEBE6!important;
    border:1px solid #4A4845!important; padding:1px 5px!important;
    border-radius:3px!important; font-size:.78rem!important;
}
[data-testid="stSidebar"] [data-testid="stExpander"]{
    background:#1C1B19!important; border:1px solid #3D3C39!important;
    border-radius:5px!important;
}
[data-testid="stSidebar"] [data-testid="stExpander"] summary{
    color:#EDEBE6!important; font-size:.83rem!important; font-weight:600!important;
}
[data-testid="stSidebar"] [data-testid="stExpander"] .stMarkdown p,
[data-testid="stSidebar"] [data-testid="stExpander"] .stMarkdown li{
    font-size:.81rem!important; line-height:1.8!important; color:#9A9893!important;
}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"]{
    background:#1C1B19!important; border:1px dashed #5A5855!important;
    border-radius:5px!important;
}

/* ── サイドバー タブ型ナビ ────────────────────────────────── */
[data-testid="stSidebar"] .stButton>button{
    width:100%!important;
    text-align:left!important;
    border:none!important;
    border-left:3px solid transparent!important;
    border-radius:0!important;
    padding:10px 16px!important;
    font-size:.85rem!important; font-weight:500!important;
    background:transparent!important; color:#9A9893!important;
    transition:background .12s,color .12s;
}
[data-testid="stSidebar"] .stButton>button:hover{
    background:rgba(192,24,32,.12)!important; color:#EDEBE6!important;
}
[data-testid="stSidebar"] .stButton>button[kind="primary"]{
    background:rgba(192,24,32,.20)!important;
    border-left:3px solid #C01820!important;
    color:#FBEBEC!important; font-weight:700!important;
}
[data-testid="stSidebar"] .stButton>button:disabled{
    opacity:.35!important; cursor:not-allowed!important;
}

/* ── メインエリア ボタン ─────────────────────────────────── */
.stButton>button[kind="primary"]{
    background:#C01820; color:#FFF;
    border:none; border-radius:5px; font-weight:600;
}
.stButton>button[kind="primary"]:hover{background:#8E1119;}
.stButton>button[kind="primary"]:disabled{background:#999!important;color:#CCC!important;}

/* ── ダウンロード ─────────────────────────────────────────── */
[data-testid="stDownloadButton"]>button{
    background:#FFF!important; color:#C01820!important;
    border:1.5px solid #C01820!important; border-radius:5px!important;
    font-weight:600!important;
}
[data-testid="stDownloadButton"]>button:hover{background:#FBEBEC!important;}

/* ── info ───────────────────────────────────────────────── */
[data-testid="stInfo"]{
    background:#FBEBEC; border-left:3px solid #C01820;
    border-radius:4px; color:#2C2C2A;
}
[data-testid="stRadio"] label,[data-testid="stCheckbox"] label{font-size:.84rem;}
hr{border-color:#E5E3DC!important;}

/* ── ページタイトルカード ────────────────────────────────── */
.page-card{
    background:#FFF; border:1px solid #E5E3DC; border-radius:8px;
    padding:14px 20px; margin-bottom:16px;
    box-shadow:0 1px 4px rgba(0,0,0,.05);
}
.page-card-title{font-size:1.0rem;font-weight:800;color:#2C2C2A;margin-bottom:2px;}
.page-card-sub{font-size:.74rem;color:#6B6A66;}

/* ── メトリクスカード ────────────────────────────────────── */
.metrics-row{display:flex;gap:10px;margin-bottom:12px;}
.m-card{
    flex:1; background:#FFF; border:1px solid #E5E3DC;
    border-radius:7px; padding:10px 16px;
    box-shadow:0 1px 3px rgba(0,0,0,.04);
    display:flex; align-items:center; gap:12px;
}
.m-val{font-size:2rem;font-weight:800;line-height:1;}
.m-lbl{font-size:.70rem;font-weight:600;color:#6B6A66;
        text-transform:uppercase;letter-spacing:.07em;}
.m-card.kaku{border-left:3px solid #8E1119;} .m-card.kaku .m-val{color:#8E1119;}
.m-card.yo  {border-left:3px solid #C01820;} .m-card.yo   .m-val{color:#C01820;}
.m-card.mi  {border-left:3px solid #9A9893;} .m-card.mi   .m-val{color:#9A9893;}

/* ── 候補パネル ──────────────────────────────────────────── */
.cand-panel{
    background:#FBEBEC; border:1.5px solid #C01820;
    border-radius:8px; padding:16px 20px; margin-bottom:12px;
}
.cand-hdr{font-size:.90rem;font-weight:700;color:#8E1119;
          margin-bottom:14px;display:flex;align-items:center;gap:8px;}
.cand-card{background:#FFF;border:1px solid #E5E3DC;border-radius:6px;padding:14px;}
.cand-card.sel{border-color:#C01820;background:#FBEBEC;}
.cand-card-title{font-size:.88rem;font-weight:700;color:#2C2C2A;margin-bottom:8px;}
.cand-card-body{font-size:.79rem;color:#6B6A66;line-height:1.9;}
.diff-chip{
    display:inline-block;background:#FBEBEC;color:#8E1119;
    border-radius:3px;padding:0 6px;font-size:.75rem;font-weight:600;
}
.cand-foot{font-size:.72rem;color:#9A9893;margin-top:10px;}

/* ── 凡例 ────────────────────────────────────────────────── */
.legend{display:flex;gap:14px;font-size:.73rem;color:#6B6A66;
        align-items:center;margin-top:6px;}
.ldot{width:9px;height:9px;border-radius:2px;
      display:inline-block;margin-right:3px;vertical-align:middle;}

/* ── 出力サマリー ─────────────────────────────────────────── */
.out-summary{
    background:#1C1B19; border:1px solid #3D3C39;
    border-radius:6px; padding:10px 14px; margin-bottom:8px;
    font-size:.82rem;
}
.out-summary table{width:100%;border-collapse:collapse;}
.out-summary td{padding:3px 0;color:#9A9893;}
.out-summary td.n{text-align:right;font-weight:700;color:#FBEBEC;}
.out-summary .note{font-size:.70rem;color:#5A5855;margin-top:6px;
                   border-top:1px solid #3D3C39;padding-top:5px;}

/* ── sublabel ────────────────────────────────────────────── */
.sublabel{border-bottom:1px solid #C01820;padding-bottom:2px;
          font-weight:700;font-size:.82rem;color:#2C2C2A;margin-bottom:6px;}

/* ── 構造化ツリー行 ─────────────────────────────────────── */
.tree-row{
    padding:5px 12px; border-bottom:1px solid #E5E3DC;
    font-size:.84rem; color:#2C2C2A; display:flex; align-items:center; gap:6px;
}
.tree-row:hover{background:#F4F2EE;}
.status-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0;}

/* ── 候補ナビ（前/次） ───────────────────────────────────── */
.cand-nav{
    display:flex; align-items:center; justify-content:space-between;
    padding:8px 0; margin-bottom:10px; font-size:.82rem; color:#6B6A66;
}

/* ── ステップ進捗バー ────────────────────────────────────── */
.step-bar{
    display:flex; align-items:flex-start;
    background:#FFF; border:1px solid #E5E3DC; border-radius:8px;
    padding:16px 24px; margin-bottom:18px;
}
.step-item{display:flex;flex-direction:column;align-items:center;min-width:80px;}
.step-line{flex:1;height:2px;margin-top:13px;border-radius:1px;}
.step-line.done{background:#8E1119;}
.step-line.future{background:#E5E3DC;}
.step-circle{
    width:26px;height:26px;border-radius:50%;
    display:flex;align-items:center;justify-content:center;
    font-size:.75rem;font-weight:700;
}
.step-circle.done{background:#8E1119;color:#FFF;}
.step-circle.active{background:#C01820;color:#FFF;box-shadow:0 0 0 3px #FBEBEC;}
.step-circle.future{background:#FFF;color:#9A9893;border:2px solid #E5E3DC;}
.step-label{font-size:.68rem;margin-top:5px;font-weight:600;white-space:nowrap;}
.step-label.done{color:#8E1119;}
.step-label.active{color:#C01820;}
.step-label.future{color:#9A9893;}

/* ── ページ下部ナビ ──────────────────────────────────────── */
.page-nav{
    display:flex; justify-content:space-between; align-items:center;
    margin-top:24px; padding-top:16px; border-top:1px solid #E5E3DC;
}
</style>
""", unsafe_allow_html=True)

# ===========================================================================
# ステップ進捗バー
# ===========================================================================
_STEP_ORDER  = ["upload", "structure", "matching", "output"]
_STEP_LABELS = ["① 取込", "② 構造化", "③ マッチング", "④ 出力"]


def _render_step_bar(current_page: str) -> None:
    """メインエリア上部のステップ進捗インジケータを描画する。"""
    cur_i = _STEP_ORDER.index(current_page) if current_page in _STEP_ORDER else 0

    html = '<div class="step-bar">'
    for i, label in enumerate(_STEP_LABELS):
        if i < cur_i:
            state, symbol = "done", "✓"
        elif i == cur_i:
            state, symbol = "active", str(i + 1)
        else:
            state, symbol = "future", str(i + 1)

        html += (
            f'<div class="step-item">'
            f'<div class="step-circle {state}">{symbol}</div>'
            f'<div class="step-label {state}">{label}</div>'
            f'</div>'
        )
        if i < len(_STEP_LABELS) - 1:
            line_cls = "done" if i < cur_i else "future"
            html += f'<div class="step-line {line_cls}"></div>'
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)


# ===========================================================================
# DB 読み込み
# ===========================================================================
def _sort_key(s: str) -> int:
    """全幅・半幅数字を含む先頭数字を整数化してソートキーにする。"""
    t = s.translate(str.maketrans("０１２３４５６７８９", "0123456789"))
    m = re.match(r"(\d+)", t.strip())
    return int(m.group(1)) if m else 9999


@st.cache_data
def load_kojyo_db():
    if not DB_PATH.exists():
        return None, None
    data = {
        "出来形管理": pd.read_excel(str(DB_PATH), sheet_name=DB_DEKIGATA,  dtype=str).fillna(""),
        "品質管理":   pd.read_excel(str(DB_PATH), sheet_name=DB_HINSHITSU, dtype=str).fillna(""),
        "撮影箇所":   pd.read_excel(str(DB_PATH), sheet_name=DB_PHOTO,     dtype=str).fillna(""),
    }
    # 編→章→節→条→枝番 の数値順にソート
    df = data["出来形管理"]
    sort_cols = ["編", "章", "節", "条", "枝番"]
    for col in sort_cols:
        if col in df.columns:
            df[f"_s_{col}"] = df[col].map(_sort_key)
    key_cols = [f"_s_{c}" for c in sort_cols if f"_s_{c}" in df.columns]
    df = df.sort_values(key_cols).drop(columns=key_cols).reset_index(drop=True)
    data["出来形管理"] = df

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
for _k in ["suryo_info","df_match","selected_idx","excluded_rows"]:
    if _k not in st.session_state: st.session_state[_k] = None
if "row_selections"  not in st.session_state: st.session_state["row_selections"]  = {}
if "confirmed_keys"  not in st.session_state: st.session_state["confirmed_keys"]  = set()
if "page"            not in st.session_state: st.session_state["page"]            = "upload"
if "excel_cache"     not in st.session_state: st.session_state["excel_cache"]     = None
if "excel_fname"     not in st.session_state: st.session_state["excel_fname"]     = None

# ===========================================================================
# ヘルパー
# ===========================================================================
def _chain_key(row):
    return tuple(str(row.get(c,"")) for c in SURYO_LEVEL_COLS)

def _group_items(items):
    g = {}
    for item in items:
        parts = [p.strip() for p in item.split(" / ")]
        g.setdefault(parts[0],[]).append(
            (item, " / ".join(parts[1:]) if len(parts)>1 else parts[0])
        )
    return g

def _deepest_name(row):
    for col in reversed(SURYO_LEVEL_COLS):
        v = row.get(col,"")
        if v: return v
    return ""

def _depth(row):
    d = 0
    for col in SURYO_LEVEL_COLS:
        if row.get(col,""): d = SURYO_LEVEL_COLS.index(col)
    return d

def _sublabel(label):
    st.markdown(f'<div class="sublabel">{label}</div>', unsafe_allow_html=True)

def _calc_status(row):
    has = (bool(str(row.get("出来形マッチ","")).strip()) or
           bool(str(row.get("品質管理マッチ","")).strip()) or
           bool(str(row.get("撮影箇所マッチ","")).strip()))
    if not has: return "未マッチ"
    if _chain_key(row) in st.session_state.confirmed_keys: return "確定"
    return "要選択"

STATUS_BG = {"確定":"#FFFFFF","要選択":"#FBEBEC","未マッチ":"#F1EFE8"}

# ─── DB ルックアップ ─────────────────────────────────────────
_DISP_D = ["測定項目","規格値","管理基準値","測定頻度","摘要"]

def _lookup_db(label, db_key):
    parts = [p.strip() for p in label.split(" / ")]
    df_db = kojyo_data[db_key]
    lvl   = ["工種","種別","細別"]
    mask  = pd.Series([True]*len(df_db), index=df_db.index)
    for i,p in enumerate(parts):
        c = lvl[i] if i<len(lvl) else None
        if c and c in df_db.columns: mask = mask & (df_db[c]==p)
    rows = df_db[mask]
    return rows.iloc[0].to_dict() if not rows.empty else {}

def _card_html(rdict, cols, diff_keys):
    html = ""
    for c in cols:
        v = str(rdict.get(c,"")).strip()
        if not v: continue
        if c in diff_keys: html += f'<div><span class="diff-chip">{v}</span></div>'
        else: html += f'<div>{c}：{v}</div>'
    return html or '<div style="color:#CCC">情報なし</div>'

def _diff_cols(rows, cols):
    if len(rows)<2: return set()
    return {c for c in cols if len({str(r.get(c,"")).strip() for r in rows})>1}

def _collect_labels(df_raw):
    out_d,out_h,out_p = [],[],[]
    seen_d,seen_h,seen_p = set(),set(),set()
    for _,row in df_raw[df_raw["状態"].isin(["確定","要選択"])].iterrows():
        ckey  = _chain_key(row)
        saved = st.session_state.row_selections.get(ckey)
        all_d = [x.strip() for x in str(row.get("出来形マッチ","")).split("\n") if x.strip()]
        all_h = [x.strip() for x in str(row.get("品質管理マッチ","")).split("\n") if x.strip()]
        all_p = [x.strip() for x in str(row.get("撮影箇所マッチ","")).split("\n") if x.strip()]
        ld = saved["出来形"]             if saved else all_d
        lh = saved["品質管理"]           if saved else all_h
        lp = saved.get("撮影箇所",all_p) if saved else all_p
        for l in ld:
            if l and l not in seen_d: out_d.append(l); seen_d.add(l)
        for l in lh:
            if l and l not in seen_h: out_h.append(l); seen_h.add(l)
        for l in lp:
            if l and l not in seen_p: out_p.append(l); seen_p.add(l)
    return out_d,out_h,out_p

def _get_df_raw():
    df = st.session_state.df_match.copy()
    df["状態"]   = df.apply(_calc_status,axis=1)
    df["_name"]  = df.apply(_deepest_name,axis=1)
    df["_depth"] = df.apply(_depth,axis=1)
    df.insert(0,"No",range(1,len(df)+1))
    return df

# ===========================================================================
# サイドバー
# ===========================================================================
with st.sidebar:
    # ─ タイトル ──────────────────────────────────────────────
    st.markdown(
        '<div style="padding:16px 16px 14px;border-bottom:1px solid #3D3C39;'
        'margin-bottom:4px;">'
        '<div style="font-size:.90rem;font-weight:800;color:#FFFFFF;'
        'letter-spacing:.03em;">施工管理計画</div>'
        '<div style="font-size:.68rem;color:#9A9893;margin-top:3px;">'
        'Automated Planning System</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    page      = st.session_state.get("page","upload")
    has_data  = st.session_state.df_match is not None
    has_sel   = (st.session_state.selected_idx is not None
                 and has_data
                 and st.session_state.selected_idx < len(st.session_state.df_match))

    # ─ ワークフローナビ ──────────────────────────────────────
    st.markdown("### WORKFLOW")
    if st.button("① 取込", use_container_width=True,
                 type="primary" if page=="upload" else "secondary", key="nav_upload"):
        st.session_state.page = "upload"; st.rerun()
    if st.button("② 構造化", use_container_width=True,
                 type="primary" if page=="structure" else "secondary",
                 disabled=not has_data, key="nav_structure"):
        st.session_state.page = "structure"; st.rerun()
    if st.button("③ マッチング", use_container_width=True,
                 type="primary" if page=="matching" else "secondary",
                 disabled=not has_data, key="nav_matching"):
        st.session_state.page = "matching"; st.rerun()
    if st.button("④ 出力", use_container_width=True,
                 type="primary" if page=="output" else "secondary",
                 disabled=not has_data, key="nav_output"):
        st.session_state.page = "output"; st.rerun()

    st.divider()

    # ─ その他ナビ ────────────────────────────────────────────
    st.markdown("### TOOLS")
    if st.button("基準DB確認", use_container_width=True,
                 type="primary" if page=="db_view" else "secondary", key="nav_db"):
        st.session_state.page = "db_view"; st.rerun()

    if st.button("使い方", use_container_width=True,
                 type="primary" if page=="help" else "secondary", key="nav_help"):
        st.session_state.page = "help"; st.rerun()

    st.divider()

    # ─ DB情報 ────────────────────────────────────────────────
    st.markdown("### 国交省基準 DB")
    st.caption(f"Ver. {version_info.get('バージョン','不明')}  "
               f"{version_info.get('作成日時','')}")
    st.caption(f"出来形 {len(kojyo_data['出来形管理'])} 行　"
               f"品質 {len(kojyo_data['品質管理'])} 行")
    st.divider()

    if st.button("↺  リセット", use_container_width=True, key="btn_reset"):
        for k in list(st.session_state.keys()):
            if k != "page": del st.session_state[k]
        st.session_state.suryo_info = st.session_state.df_match = \
            st.session_state.selected_idx = None
        st.session_state.row_selections  = {}
        st.session_state.confirmed_keys  = set()
        st.session_state.excel_cache = st.session_state.excel_fname = None
        st.session_state.page = "upload"
        st.rerun()

# ===========================================================================
# ① 取込ページ
# ===========================================================================
def _render_upload():
    _render_step_bar("upload")
    st.markdown(
        '<div class="page-card">'
        '<div class="page-card-title">① 取込 — 数量総括表PDFのアップロード</div>'
        '<div class="page-card-sub">PDFを読み込み、工種・種別・細別の階層を抽出します</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    if st.session_state.suryo_info:
        name = st.session_state.suryo_info.get("工事名","")
        st.success(f"読込済み：{name}")
        if st.button("別のPDFを読み込む", key="re_upload"):
            st.session_state.suryo_info  = None
            st.session_state.df_match    = None
            st.session_state.selected_idx = None
            st.session_state.row_selections = {}
            st.session_state.confirmed_keys = set()
            st.session_state.excel_cache = None
            st.rerun()
        # 下部ナビ
        st.markdown('<div class="page-nav"><div></div>', unsafe_allow_html=True)
        if st.button("② 構造化を確認する →", type="primary", key="go_structure"):
            st.session_state.page = "structure"; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
        return

    uploaded = st.file_uploader("数量総括表PDFをドラッグ＆ドロップ、またはクリックで選択",
                                 type="pdf", label_visibility="visible")

    if uploaded:
        col1, col2 = st.columns([1, 3])
        with col1:
            if st.button("解析する", type="primary", use_container_width=True, key="do_parse"):
                try:
                    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
                        f.write(uploaded.read()); path_tmp = f.name
                    with st.spinner("PDF解析・マッチング中..."):
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
                    st.session_state.excluded_rows  = si.get("除外行", None)
                    st.session_state.selected_idx   = None
                    st.session_state.row_selections = {}
                    st.session_state.confirmed_keys = set()
                    st.session_state.excel_cache    = None
                    st.session_state.excel_fname    = None
                    st.session_state.page           = "structure"
                    st.rerun()
                except Exception:
                    st.error("解析中にエラーが発生しました。")
                    with st.expander("エラー詳細"): st.code(traceback.format_exc())

# ===========================================================================
# ② 構造化ページ
# ===========================================================================
def _render_structure():
    si = st.session_state.suryo_info
    _render_step_bar("structure")
    st.markdown(
        f'<div class="page-card">'
        f'<div class="page-card-title">② 構造化 — 抽出結果の確認</div>'
        f'<div class="page-card-sub">{si.get("工事名","") if si else ""} '
        f'　抽出行数：{len(st.session_state.df_match)} 行</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    df_raw = _get_df_raw()

    DOT_COLOR = {"確定":"#8E1119","要選択":"#C01820","未マッチ":"#9A9893"}

    # テーブルとして表示
    df_disp = pd.DataFrame({
        "工種・種別・細別": ["　"*row["_depth"]+row["_name"] for _,row in df_raw.iterrows()],
        "マッチ状態":       [row["状態"] for _,row in df_raw.iterrows()],
    })
    sts_idx = {i: row["状態"] for i,(_,row) in enumerate(df_raw.iterrows())}

    STATUS_BG2 = {"確定":"#FFFFFF","要選択":"#FBEBEC","未マッチ":"#F1EFE8"}
    def _rs(sts):
        def _s(row):
            bg = STATUS_BG2.get(sts.get(row.name,""),"")
            return [f"background-color:{bg}" if bg else "" for _ in row]
        return _s

    st.dataframe(
        df_disp.style.apply(_rs(sts_idx),axis=1),
        use_container_width=True, height=380, hide_index=True,
        column_config={
            "工種・種別・細別": st.column_config.TextColumn(width="large"),
            "マッチ状態":       st.column_config.TextColumn(width="small"),
        }
    )
    st.markdown(
        '<div class="legend">'
        '<span><span class="ldot" style="background:#8E1119"></span>確定</span>'
        '<span><span class="ldot" style="background:#C01820"></span>要選択</span>'
        '<span><span class="ldot" style="background:#9A9893"></span>未マッチ</span>'
        '</div>',
        unsafe_allow_html=True,
    )
    # ── 除外された行 ─────────────────────────────────────────
    df_ex = st.session_state.get("excluded_rows")
    if df_ex is not None and not df_ex.empty:
        with st.expander(f"除外された行を確認する（{len(df_ex)} 件）"):
            st.markdown("""
以下の行は **施工管理基準の照合対象外** として自動除外されました。
除外の判定は次の3条件に基づいています。

| 除外理由 | 判定条件 | 例 |
|---|---|---|
| **費用集計項目** | 「直接工事費」「共通仮設費」「現場管理費」など工事費の集計行として登録された名称と一致 | 直接工事費、純工事費、工事原価 など |
| **小計・合計行** | 行の先頭が `(` または `（` で始まる | （計）、（小計） など |
| **ヘッダー行** | 「工事区分」または「工事名」を含む | テーブルの見出し行 |

誤って除外されている項目がある場合は管理者にご連絡ください。
""")
            st.dataframe(
                df_ex,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "項目名":   st.column_config.TextColumn("項目名",   width="large"),
                    "除外理由": st.column_config.TextColumn("除外理由", width="medium"),
                },
            )

    # 下部ナビ
    st.markdown('<div class="page-nav">', unsafe_allow_html=True)
    nav_l, nav_r = st.columns(2)
    with nav_l:
        if st.button("← 取込に戻る", key="struct_back"):
            st.session_state.page = "upload"; st.rerun()
    with nav_r:
        if st.button("③ マッチングを確認する →", type="primary",
                     use_container_width=True, key="go_matching"):
            st.session_state.page = "matching"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# ===========================================================================
# ③ マッチング＋候補選択（同一ページ）
# ===========================================================================
def _render_matching():
    _render_step_bar("matching")
    df_raw = _get_df_raw()
    n_kaku = int((df_raw["状態"]=="確定").sum())
    n_yo   = int((df_raw["状態"]=="要選択").sum())
    n_mi   = int((df_raw["状態"]=="未マッチ").sum())

    # 要対応キュー情報（ページ全体で使用）
    yo_idxs = [i for i,(_,r) in enumerate(df_raw.iterrows()) if r["状態"]=="要選択"]
    confirmed_yo = sum(1 for i in yo_idxs
                       if _chain_key(df_raw.iloc[i]) in st.session_state.confirmed_keys)
    remaining = n_yo - confirmed_yo

    st.markdown(
        '<div class="page-card">'
        '<div class="page-card-title">③ マッチング — 国交省基準DBとの対応確認</div>'
        '<div class="page-card-sub">行をクリックして候補を確認・選択できます</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    # メトリクス
    st.markdown(
        f'<div class="metrics-row">'
        f'<div class="m-card kaku"><div class="m-val">{n_kaku}</div>'
        f'<div class="m-lbl">確定</div></div>'
        f'<div class="m-card yo"><div class="m-val">{n_yo}</div>'
        f'<div class="m-lbl">要選択</div></div>'
        f'<div class="m-card mi"><div class="m-val">{n_mi}</div>'
        f'<div class="m-lbl">未マッチ</div></div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # ── 全確定バナー（要選択がゼロになったとき） ─────────────
    if n_yo == 0 and (n_kaku > 0 or n_mi > 0):
        ok_col, go_col = st.columns([4, 1])
        with ok_col:
            st.markdown(
                '<div style="display:flex;align-items:center;height:38px;'
                'padding:0 14px;background:#D1FAE5;border-left:4px solid #10B981;'
                'border-radius:4px;font-size:.88rem;color:#065F46;font-weight:600;">'
                '✓ 要選択がすべて確定済みです。</div>',
                unsafe_allow_html=True,
            )
        with go_col:
            if st.button("④ 出力へ →", type="primary",
                         use_container_width=True, key="match_to_output"):
                st.session_state.page = "output"; st.rerun()

    # ── 要対応キュー進捗バナー ───────────────────────────────
    elif n_yo > 0:
        if remaining > 0:
            first_unc = next(
                (i for i in yo_idxs
                 if _chain_key(df_raw.iloc[i]) not in st.session_state.confirmed_keys),
                None,
            )
            pcol, jcol, acol = st.columns([4, 1, 1])
            with pcol:
                st.progress(
                    confirmed_yo / n_yo,
                    text=f"要選択 {n_yo} 件中 {confirmed_yo} 件確認済み　— 残り **{remaining} 件**",
                )
            with jcol:
                if first_unc is not None:
                    if st.button("未確認へ →", use_container_width=True, key="jump_unc"):
                        st.session_state.selected_idx = first_unc
                        st.rerun()
            with acol:
                if st.button("すべて確定", use_container_width=True, key="confirm_all_btn"):
                    for i in yo_idxs:
                        row = df_raw.iloc[i]
                        k = _chain_key(row)
                        st.session_state.confirmed_keys.add(k)
                        if k not in st.session_state.row_selections:
                            all_d = [x.strip() for x in str(row.get("出来形マッチ","")).split("\n") if x.strip()]
                            all_h = [x.strip() for x in str(row.get("品質管理マッチ","")).split("\n") if x.strip()]
                            all_p = [x.strip() for x in str(row.get("撮影箇所マッチ","")).split("\n") if x.strip()]
                            st.session_state.row_selections[k] = {"出来形": all_d, "品質管理": all_h, "撮影箇所": all_p}
                    st.toast("要選択をすべて確定しました")
                    st.rerun()
        else:
            ok_col, go_col = st.columns([4, 1])
            with ok_col:
                st.success(f"要選択 {n_yo} 件すべて確認済みです。")
            with go_col:
                if st.button("④ 出力へ →", type="primary",
                             use_container_width=True, key="match_to_output"):
                    st.session_state.page = "output"; st.rerun()

    # フィルタ
    sel_idx = st.session_state.selected_idx
    has_sel = sel_idx is not None and 0 <= sel_idx < len(df_raw)
    sel_ckey = _chain_key(df_raw.iloc[sel_idx]) if has_sel else None
    is_sel_confirmed = sel_ckey is not None and sel_ckey in st.session_state.confirmed_keys

    radio_col, undo_col = st.columns([4, 1])
    with radio_col:
        filter_opt = st.radio(
            "filter", ["すべて","要選択のみ","確定のみ","未マッチのみ"],
            horizontal=True, label_visibility="collapsed",
        )
    with undo_col:
        if is_sel_confirmed:
            if st.button("確定を取り消す", use_container_width=True, key="unconfirm_top"):
                st.session_state.confirmed_keys.discard(sel_ckey)
                st.toast(f"「{df_raw.iloc[sel_idx]['_name']}」の確定を解除しました")
                st.rerun()

    FM = {"確定のみ":"確定","要選択のみ":"要選択","未マッチのみ":"未マッチ"}
    df_v = (df_raw[df_raw["状態"]==FM[filter_opt]].copy()
            if filter_opt in FM else df_raw.copy())

    def _fmt(row):
        d = str(row.get("出来形マッチ","")).strip()
        h = str(row.get("品質管理マッチ","")).strip()
        p = str(row.get("撮影箇所マッチ","")).strip()
        if not (d or h or p): return "—"
        dl = [x for x in d.split("\n") if x.strip()]
        cats = (["出来形"] if d else [])+(["品質"] if h else [])+(["撮影"] if p else [])
        return f"候補{len(dl)}件（工法で分岐）" if len(dl)>=2 else "・".join(cats)

    tbl_h   = 280 if has_sel else 430

    sts_idx = {i: row["状態"] for i,(_,row) in enumerate(df_v.iterrows())}
    df_tbl = pd.DataFrame({
        "工種・項目":    ["　"*row["_depth"]+row["_name"] for _,row in df_v.iterrows()],
        "マッチした基準": [_fmt(row) for _,row in df_v.iterrows()],
        "状態":         [row["状態"] for _,row in df_v.iterrows()],
    })

    def _rs(sts):
        def _s(row):
            bg = STATUS_BG.get(sts.get(row.name,""),"")
            return [f"background-color:{bg}" if bg else "" for _ in row]
        return _s

    ev = st.dataframe(
        df_tbl.style.apply(_rs(sts_idx),axis=1),
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
        if st.session_state.selected_idx != no - 1:
            st.session_state.selected_idx = no - 1
            st.rerun()

    st.markdown(
        '<div class="legend">'
        '<span><span class="ldot" style="background:#8E1119"></span>確定</span>'
        '<span><span class="ldot" style="background:#C01820"></span>要選択（行クリックで候補展開）</span>'
        '<span><span class="ldot" style="background:#9A9893"></span>未マッチ</span>'
        '</div>',
        unsafe_allow_html=True,
    )

    # 候補パネル（行選択時のみ）
    if not has_sel:
        st.markdown(
            '<div style="margin-top:10px;padding:18px;background:#FFF;'
            'border:1px dashed #E5E3DC;border-radius:8px;text-align:center;'
            'color:#9A9893;font-size:.84rem;">'
            '行をクリックすると候補がここに展開されます</div>',
            unsafe_allow_html=True,
        )
    else:
        sel   = df_raw.iloc[sel_idx]
        ckey  = _chain_key(sel)
        chain = " › ".join(sel[c] for c in SURYO_LEVEL_COLS if sel.get(c,""))

        cur_pos = yo_idxs.index(sel_idx) if sel_idx in yo_idxs else None

        items_d = [x.strip() for x in str(sel.get("出来形マッチ","")).split("\n") if x.strip()]
        items_h = [x.strip() for x in str(sel.get("品質管理マッチ","")).split("\n") if x.strip()]
        items_p = [x.strip() for x in str(sel.get("撮影箇所マッチ","")).split("\n") if x.strip()]

        if not items_d and not items_h and not items_p:
            st.markdown(
                f'<div style="margin-top:10px;padding:14px;background:#F4F2EE;'
                f'border:1px solid #E5E3DC;border-radius:8px;font-size:.84rem;color:#9A9893;">'
                f'「{sel["_name"]}」はDBマッチなし（未マッチ）</div>',
                unsafe_allow_html=True,
            )
        else:
            saved = st.session_state.row_selections.get(ckey)

            # 進捗＋ナビ（要選択行のみ）
            if cur_pos is not None and yo_idxs:
                done  = sum(1 for i in yo_idxs
                            if _chain_key(df_raw.iloc[i]) in st.session_state.confirmed_keys)
                total = len(yo_idxs)
                prog_col, nav_col = st.columns([5, 1])
                with prog_col:
                    st.progress(done/total, text=f"要選択 {total} 件中 {done} 件確認済み")
                with nav_col:
                    if cur_pos < len(yo_idxs)-1:
                        if st.button("次へ →", key="cand_next"):
                            st.session_state.selected_idx = yo_idxs[cur_pos+1]
                            st.rerun()

            st.markdown(
                f'<div class="cand-panel">'
                f'<div class="cand-hdr">要確認: {sel["_name"]}'
                f'<span style="font-weight:400;font-size:.76rem;margin-left:8px;">{chain}</span>'
                f'</div>',
                unsafe_allow_html=True,
            )

            # ── 出来形 比較カード ──────────────────────────────────
            new_sel_d = []
            if len(items_d) >= 2:
                db_rows_d = [_lookup_db(lbl,"出来形管理") for lbl in items_d[:4]]
                diff_d    = _diff_cols(db_rows_d, _DISP_D)
                cur_d     = saved["出来形"] if saved else items_d
                cols_c    = st.columns(min(len(items_d),4))
                for i,(col,lbl) in enumerate(zip(cols_c, items_d[:4])):
                    with col:
                        parts  = [p.strip() for p in lbl.split(" / ")]
                        ctitle = " / ".join(parts[1:]) if len(parts)>1 else parts[0]
                        is_sel = lbl in cur_d
                        brd    = "border:1.5px solid #C01820;background:#FBEBEC;" if is_sel else ""
                        body   = _card_html(db_rows_d[i], _DISP_D, diff_d)
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
            elif len(items_d) == 1:
                lbl    = items_d[0]
                db_row = _lookup_db(lbl, "出来形管理")
                parts  = [p.strip() for p in lbl.split(" / ")]
                ctitle = " / ".join(parts[1:]) if len(parts)>1 else parts[0]
                body   = _card_html(db_row, _DISP_D, set())
                st.markdown(
                    f'<div class="cand-card" style="border:1.5px solid #C01820;background:#FBEBEC;">'
                    f'<div class="cand-card-title">出来形基準：{ctitle}</div>'
                    f'<div class="cand-card-body">{body}</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )
                new_sel_d = items_d

            st.markdown('</div>', unsafe_allow_html=True)

            # 品質・撮影 expander
            new_sel_h, new_sel_p = items_h, items_p
            if items_h or items_p:
                with st.expander("品質管理・撮影箇所の候補を調整"):
                    ch,cp = st.columns(2)
                    with ch:
                        new_sel_h = []
                        if items_h:
                            _sublabel("品質管理")
                            for kojyo,sub in _group_items(items_h).items():
                                if len(_group_items(items_h))>1: st.caption(kojyo)
                                for fl,dl in sub:
                                    if st.checkbox(dl,value=True,key=f"chk_h_{sel_idx}_{items_h.index(fl)}"):
                                        new_sel_h.append(fl)
                        else:
                            st.caption("該当なし")
                    with cp:
                        new_sel_p = []
                        if items_p:
                            _sublabel("撮影箇所")
                            for kojyo,sub in _group_items(items_p).items():
                                if len(_group_items(items_p))>1: st.caption(kojyo)
                                for fl,dl in sub:
                                    if st.checkbox(dl,value=True,key=f"chk_p_{sel_idx}_{items_p.index(fl)}"):
                                        new_sel_p.append(fl)
                        else:
                            st.caption("該当なし")

            st.session_state.row_selections[ckey] = {
                "出来形":  new_sel_d,
                "品質管理": new_sel_h,
                "撮影箇所": new_sel_p,
            }

            # ── 確定して次へ ──────────────────────────────────────
            if cur_pos is not None:
                btn_col, _ = st.columns([2, 5])
                with btn_col:
                    if cur_pos < len(yo_idxs) - 1:
                        if st.button("確定して次へ →", type="primary",
                                     use_container_width=True, key="confirm_next_btn"):
                            st.session_state.confirmed_keys.add(ckey)
                            st.toast(f"「{sel['_name']}」を確定しました")
                            st.session_state.selected_idx = yo_idxs[cur_pos + 1]
                            st.rerun()
                    else:
                        if st.button("確定（最終項目）✓", type="primary",
                                     use_container_width=True, key="confirm_last_btn"):
                            st.session_state.confirmed_keys.add(ckey)
                            st.toast("すべての要選択を確認しました。④出力へ進んでください。")
                            st.rerun()

            if items_d:
                fd = items_d[0].split(" / ")[0]
                dr = kojyo_data["出来形管理"][kojyo_data["出来形管理"]["工種"]==fd]
                if not dr.empty:
                    r  = dr.iloc[0]
                    bc = " › ".join(x for x in [r.get("編",""),r.get("章",""),r.get("節",""),fd] if x)
                    st.caption(f"DB 目次：{bc}")

    # ── ページ下部ナビ（常時表示） ────────────────────────────
    st.markdown('<div class="page-nav">', unsafe_allow_html=True)
    nav_l, nav_r = st.columns(2)
    with nav_l:
        if st.button("← 構造化に戻る", key="match_back"):
            st.session_state.page = "structure"; st.rerun()
    with nav_r:
        if st.button("④ 出力へ →", type="primary",
                     use_container_width=True, key="match_next"):
            st.session_state.page = "output"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)


def _render_output():
    """④ 出力ページ。"""
    si = st.session_state.suryo_info
    _render_step_bar("output")
    st.markdown(
        f'<div class="page-card">'
        f'<div class="page-card-title">④ 出力 — 施工管理計画 Excel 生成</div>'
        f'<div class="page-card-sub">{si.get("工事名","") if si else ""}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    df_tmp = _get_df_raw()
    out_d_l, out_h_l, out_p_l = _collect_labels(df_tmp)
    can_out = bool(out_d_l or out_h_l or out_p_l)

    n_kaku = int((df_tmp["状態"] == "確定").sum())
    n_yo   = int((df_tmp["状態"] == "要選択").sum())
    n_mi   = int((df_tmp["状態"] == "未マッチ").sum())

    # ── マッチング状況サマリー ────────────────────────────────
    st.markdown(
        f'<div class="metrics-row">'
        f'<div class="m-card kaku"><div class="m-val">{n_kaku}</div>'
        f'<div class="m-lbl">確定</div></div>'
        f'<div class="m-card yo"><div class="m-val">{n_yo}</div>'
        f'<div class="m-lbl">要選択</div></div>'
        f'<div class="m-card mi"><div class="m-val">{n_mi}</div>'
        f'<div class="m-lbl">未マッチ</div></div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # ── 出力内容確認 ─────────────────────────────────────────
    st.markdown(
        f'<div style="background:#F8FAFC;border:1px solid #E2E6EA;border-radius:8px;'
        f'padding:20px 24px;margin-bottom:16px;">'
        f'<div style="font-size:.88rem;font-weight:700;color:#1A2332;margin-bottom:12px;">'
        f'出力内容</div>'
        f'<table style="width:100%;font-size:.84rem;border-collapse:collapse;">'
        f'<tr><td style="padding:5px 0;color:#6B6A66;">出来形管理</td>'
        f'<td style="text-align:right;font-weight:700;color:#C01820;">{len(out_d_l)} 件</td></tr>'
        f'<tr><td style="padding:5px 0;color:#6B6A66;">品質管理</td>'
        f'<td style="text-align:right;font-weight:700;color:#C01820;">{len(out_h_l)} 件</td></tr>'
        f'<tr><td style="padding:5px 0;color:#6B6A66;">撮影箇所</td>'
        f'<td style="text-align:right;font-weight:700;color:#C01820;">{len(out_p_l)} 件</td></tr>'
        f'</table>'
        f'<div style="font-size:.73rem;color:#AAA;margin-top:10px;border-top:1px solid #EEE;padding-top:8px;">'
        f'要選択で未確認の行は全候補を自動採用します</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    if n_yo > 0:
        confirmed = sum(
            1 for i, (_, r) in enumerate(df_tmp.iterrows())
            if r["状態"] == "要選択" and _chain_key(r) in st.session_state.row_selections
        )
        st.info(f"要選択 {n_yo} 件中 {confirmed} 件が確定済みです。未確認の {n_yo - confirmed} 件は全候補を自動採用します。"
                f"　→ ③マッチングで確認できます。")

    # ── ボタン ───────────────────────────────────────────────
    col_out, col_dl, _ = st.columns([2, 1, 1])
    with col_out:
        if st.button("施工管理計画を出力", type="primary",
                     use_container_width=True, disabled=not can_out, key="btn_out"):
            try:
                with st.spinner("Excel生成中..."):
                    filtered = filter_by_row_labels(kojyo_data, out_d_l, out_h_l, out_p_l)
                    df_raw_out = _get_df_raw()
                    dmap = {}
                    for _, row in df_raw_out[df_raw_out["状態"].isin(["確定", "要選択"])].iterrows():
                        ckey = _chain_key(row)
                        suryo_kojyo = str(row.get("工種", "")).strip()
                        saved = st.session_state.row_selections.get(ckey)
                        all_d = [x.strip() for x in str(row.get("出来形マッチ", "")).split("\n") if x.strip()]
                        selected_d = saved["出来形"] if saved else all_d
                        for label in selected_d:
                            db_kojyo = label.split(" / ")[0].strip()
                            if db_kojyo and db_kojyo not in dmap:
                                dmap[db_kojyo] = suryo_kojyo
                    excel_bytes = write_excel(filtered, 工事名=si["工事名"], dekigata_kojyo_map=dmap)
                safe = re.sub(r'[\\/:*?"<>|　 ]', '_', si["工事名"])
                st.session_state.excel_cache = excel_bytes
                st.session_state.excel_fname = (f"施工管理計画_{safe}.xlsx" if safe
                                                else "施工管理計画.xlsx")
                st.success("生成完了！ダウンロードボタンからファイルを取得してください。")
                st.rerun()
            except Exception:
                st.error("生成エラー")
                with st.expander("詳細"): st.code(traceback.format_exc())

    with col_dl:
        if st.session_state.excel_cache:
            st.download_button(
                "⬇  ダウンロード",
                data=st.session_state.excel_cache,
                file_name=st.session_state.excel_fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    # 下部ナビ
    st.markdown('<div class="page-nav">', unsafe_allow_html=True)
    if st.button("← マッチングに戻る", key="output_back"):
        st.session_state.page = "matching"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)


# ===========================================================================
# 使い方ページ
# ===========================================================================
def _render_help():
    st.markdown(
        '<div class="page-card">'
        '<div class="page-card-title">使い方</div>'
        '<div class="page-card-sub">施工管理計画 自動生成アプリの操作ガイド</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    tab0, tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "概要", "① 取込", "② 構造化", "③ マッチング", "④ 出力", "よくある問題"
    ])

    # ── 概要 ──────────────────────────────────────────────────
    with tab0:
        st.markdown("""
### このアプリでできること

数量総括表（PDF）をアップロードするだけで、国交省の施工管理基準に基づいた
**施工管理計画（Excel）を自動生成**します。

手作業で基準書を調べる工程をなくし、転記ミスを防ぎます。

---

### 処理の流れ

```
数量総括表PDF
    ↓ ① 取込
工種・種別・細別の階層を自動抽出
    ↓ ② 構造化確認
抽出結果を目視で確認
    ↓ ③ マッチング
国交省基準DB（出来形・品質・撮影）と自動照合
複数候補がある場合は工法を選択
    ↓ ④ 出力
施工管理計画 Excel ファイルを生成・ダウンロード
```

---

### 画面のナビゲーション

左サイドバーは2つのセクションに分かれています。

| セクション | ボタン | 説明 |
|---|---|---|
| **WORKFLOW** | ① 取込 〜 ④ 出力 | 作業ステップを直接開く。どのページからでも移動可能。PDF未読込時は②〜④はグレーアウト |
| **TOOLS** | 基準DB確認 / 使い方 | 参照ページ。作業中でもいつでも開け、WORKFLOWボタンで作業に戻れる |

---

### 生成されるExcelのシート構成

| シート名 | 内容 |
|---|---|
| (8) 品管一覧 | 品質管理基準及び規格値（試験項目・規格値・社内規格値） |
| (8) 出来形一覧 | 出来形管理基準及び規格値（測定項目・規格値・測定箇所） |
| (8) 撮影箇所 | 撮影箇所一覧表（区分・撮影項目・撮影時期・頻度） |

---

### 基準DBについて

サイドバーの「基準DB確認」から、照合に使用している国交省基準DBの内容を
工種絞り込み・キーワード検索で確認できます。DBのExcelもダウンロード可能です。
""")

    # ── ① 取込 ───────────────────────────────────────────────
    with tab1:
        st.markdown("""
### ① 取込 — 数量総括表PDFのアップロード

#### 対応するPDFの形式

- **表形式**で工種・種別・細別・名称が記載されているもの
- 国土交通省の標準様式「数量総括表」を想定

#### 操作手順

1. 「取込」ページを開く（起動直後はここが表示されます）
2. PDFをドラッグ＆ドロップ、またはクリックしてファイルを選択
3. 「解析する」ボタンをクリック
4. 自動で構造化・マッチングが実行され、② 構造化ページへ遷移します

#### 注意事項

- **読み込み直す場合**: 「別のPDFを読み込む」ボタンで再度アップロードできます
  （マッチング結果・選択内容はリセットされます）
- スキャンされたPDF（画像PDF）は読み取れません。テキストが含まれるPDFを使用してください
- ページ数が多いPDFは解析に時間がかかる場合があります
""")

    # ── ② 構造化 ─────────────────────────────────────────────
    with tab2:
        st.markdown("""
### ② 構造化 — 抽出結果の確認

#### 画面の見方

抽出された工種・種別・細別の階層が一覧表示されます。
各行には国交省基準DBとの照合結果（マッチ状態）が色で示されます。

| 色 | 状態 | 意味 |
|---|---|---|
| 白（えんじ左線） | 確定 | 基準と1対1で照合済み |
| 淡えんじ | 要選択 | 複数候補あり・工法の選択が必要 |
| 淡グレー | 未マッチ | 対応する基準が見つからない |

#### 確認ポイント

- **未マッチ行が多い場合**: 工種名の表記が基準DBと異なる可能性があります
  「基準DB確認」ページで工種名を検索して確認してください
- **階層が崩れている場合**: PDFの表レイアウトが非標準の可能性があります

#### 除外される行について

数量総括表には施工管理基準と照合すべき工種以外に、費用の集計行やヘッダー行が含まれます。
これらは自動的に除外され、照合対象から外されます。

| 除外理由 | 判定条件 | 例 |
|---|---|---|
| **費用集計項目** | 登録済みの集計行名称と一致 | 直接工事費・共通仮設費・現場管理費・工事価格・消費税相当額 など |
| **小計・合計行** | 行の先頭が `(` または `（` で始まる | （計）・（小計） など |
| **ヘッダー行** | 「工事区分」または「工事名」を含む | テーブルの見出し行 |

除外された行は「除外された行を確認する」セクションで一覧表示されます。
本来照合すべき工種が除外されている場合は管理者にご連絡ください。

#### 次のステップ

内容を確認したら「③ マッチングを確認する →」ボタンで進みます。
""")

    # ── ③ マッチング ─────────────────────────────────────────
    with tab3:
        st.markdown("""
### ③ マッチング — 候補の確認と選択

#### 画面の構成

**上部**: 確定 / 要選択 / 未マッチの件数カード

**進捗バナー**: 要選択の残件数と進捗バーを表示。「未確認へ →」で未確認の行へジャンプ、「すべて確定」で一括確定できます

**一覧表**: 全工種とマッチ状態を表示。行をクリックすると下部に候補パネルが展開されます

**候補パネル**: 選択した行の出来形基準・品質管理・撮影箇所の候補を表示

---

#### 要選択行の処理手順

「要選択」行は工法によって適用する基準が異なるため、どの基準を使うか選択が必要です。

1. 進捗バナーの「未確認へ →」ボタン、またはフィルタで「要選択のみ」を表示
2. 一覧の行をクリックして候補パネルを開く
3. **比較カード**を見て、適用する候補に「採用」チェックを入れる
   - 差分がある項目は強調表示されます（例: 規格値・測定頻度の違い）
4. 「確定して次へ →」ボタンで確定し、次の要選択行へ進む
5. 要選択がすべて確定されると、上部に「④ 出力へ →」ボタンが表示されます

#### 一括確定

「すべて確定」ボタンを押すと、要選択の全行をまとめて確定できます（各行のデフォルト候補が採用されます）。

#### 確定の取り消し

確定済みの行を一覧でクリックすると、フィルタ横に**「確定を取り消す」**ボタンが表示されます。
クリックすると「要選択」状態に戻り、候補を選び直せます。

#### 比較カードの読み方

```
候補A: ○○工（機械掘削）     候補B: ○○工（人力掘削）
測定項目: 幅・深さ            測定項目: 幅・深さ
規格値: ±50mm               規格値: ±30mm   ← 差分
測定頻度: 200m毎             測定頻度: 100m毎 ← 差分
```

差分のある項目は赤背景のチップで強調表示されます。

#### 品質管理・撮影箇所の調整

候補パネル下部の「品質管理・撮影箇所の候補を調整」を開くと、
品質管理と撮影箇所の対応も個別に選択できます。

#### フィルタの使い方

| フィルタ | 表示内容 |
|---|---|
| すべて | 全行を表示 |
| 要選択のみ | 未確認の行に絞り込み |
| 確定のみ | 確定済みの行を確認 |
| 未マッチのみ | 基準なしの行を確認 |

ページ下部の「④ 出力へ →」ボタンはいつでも押せます。
""")

    # ── ④ 出力 ───────────────────────────────────────────────
    with tab4:
        st.markdown("""
### ④ 出力 — Excel生成とダウンロード

#### 出力前の確認

ページ上部に出力内容の件数が表示されます。

- **出来形管理 X 件**: 出来形一覧に出力される行数
- **品質管理 X 件**: 品管一覧に出力される行数
- **撮影箇所 X 件**: 撮影箇所一覧に出力される行数

要選択で未確認の行がある場合、**全候補を自動採用**して出力します。
気になる行は「③ マッチングに戻る」で確認してから出力することを推奨します。

#### 操作手順

1. 「施工管理計画を出力」ボタンをクリック
2. Excel生成が完了したら「ダウンロード」ボタンが有効になります
3. クリックしてファイルを保存してください

#### 出力ファイルの仕様

- ファイル名: `施工管理計画_（工事名）.xlsx`
- フォント: MS明朝
- 工種・種別列の連続する同じ値は縦結合されます
- 社内規格値は規格値の数値に0.8を乗じた値が自動入力されます

#### 再生成

内容を修正したい場合は「施工管理計画を出力」ボタンをもう一度クリックすると上書き生成されます。
""")

    # ── よくある問題 ─────────────────────────────────────────
    with tab5:
        st.markdown("""
### よくある問題

---

#### PDFを読み込んでも工種が抽出されない

**原因**: スキャンPDF（画像のみ）、またはPDFの表レイアウトが非標準である可能性があります。

**対処**:
- PDFをテキスト選択できるか確認してください（できなければ画像PDFです）
- 別のPDFビューアで開いて表が正しく表示されているか確認してください

---

#### マッチング結果の「未マッチ」が多い

**原因**: 数量総括表の工種名と国交省基準DBの工種名の表記が異なる場合に発生します。

**対処**:
1. 「基準DB確認」ページで工種名を検索して、DBに登録されている正式名称を確認
2. 数量総括表の工種名が基準DBと一致しているか確認してください

---

#### 生成されたExcelの社内規格値がおかしい

**仕様**: 社内規格値は規格値の数値に **0.8** を乗じた値を自動計算しています。
規格値が「±50mm」の場合、社内規格値は「±40mm」になります。

数値を含まない規格値（「設計値以上」など）はそのまま転記されます。

---

#### 国交省基準DBを更新したい

基準PDFが改訂された場合は、管理者がターミナルで以下を実行してください:

```bash
python build_db.py <施工管理基準.pdf> <写真管理基準.pdf> --version "令和○年○月版"
```

実行後、アプリを再起動するとDBが反映されます。

---

#### アプリをリセットしたい

サイドバー下部の「↺ リセット」ボタンで、取込・マッチング・選択内容をすべて初期化できます。
""")


# ===========================================================================
# 基準DB確認
# ===========================================================================
def _df_search(df: pd.DataFrame, keyword: str) -> pd.DataFrame:
    """全列を対象にキーワード絞り込み（大文字小文字無視）。"""
    if not keyword.strip():
        return df
    kw = keyword.strip().lower()
    mask = df.apply(lambda col: col.astype(str).str.lower().str.contains(kw, na=False)).any(axis=1)
    return df[mask]


def _render_db_tab(df: pd.DataFrame, tab_key: str, keyword: str) -> None:
    """DB各タブ共通：工種フィルタ＋件数表示＋dataframe。"""
    kojyo_opts = sorted(df["工種"].dropna().unique().tolist()) if "工種" in df.columns else []
    selected = st.multiselect(
        "工種で絞り込み",
        kojyo_opts,
        key=f"db_kojyo_{tab_key}",
        placeholder="すべての工種",
    )
    df_f = df[df["工種"].isin(selected)] if selected else df
    df_f = _df_search(df_f, keyword)
    st.caption(f"{len(df_f):,} / {len(df):,} 行表示")
    st.dataframe(df_f, use_container_width=True, height=520, hide_index=True)


def _render_db_view():
    st.markdown(
        '<div class="page-card">'
        '<div class="page-card-title">国交省基準 DB</div>'
        f'<div class="page-card-sub">Ver. {version_info.get("バージョン","不明")}　'
        f'{version_info.get("作成日時","")}</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    # ── 検索 & ダウンロード ────────────────────────────────────
    col_kw, col_dl = st.columns([4, 1])
    with col_kw:
        keyword = st.text_input(
            "キーワード検索",
            placeholder="工種・測定項目・規格値 などで絞り込み",
            label_visibility="collapsed",
        )
    with col_dl:
        with open(str(DB_PATH), "rb") as _f:
            st.download_button(
                "DB Excel をダウンロード",
                data=_f.read(),
                file_name=DB_PATH.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    # ── タブ ─────────────────────────────────────────────────
    tab_d, tab_h, tab_p = st.tabs(["出来形管理", "品質管理", "撮影箇所"])
    with tab_d:
        _render_db_tab(kojyo_data["出来形管理"], "d", keyword)
    with tab_h:
        _render_db_tab(kojyo_data["品質管理"], "h", keyword)
    with tab_p:
        _render_db_tab(kojyo_data["撮影箇所"], "p", keyword)

# ===========================================================================
# ルーティング
# ===========================================================================
page = st.session_state.page

if page == "help":
    _render_help()
elif page == "db_view":
    _render_db_view()
elif page == "structure" and has_data:
    _render_structure()
elif page in ("matching", "candidate") and has_data:
    _render_matching()
elif page == "output" and has_data:
    _render_output()
else:
    _render_upload()
