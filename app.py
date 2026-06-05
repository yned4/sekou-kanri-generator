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
.stApp,[data-testid="stAppViewContainer"],.main{background:#F2F4F7!important;}
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

/* ── サイドバー（ダーク紺） ───────────────────────────────── */
[data-testid="stSidebar"]{
    background:#1A2332!important;
    border-right:1px solid #263447!important;
    min-width:220px!important; max-width:220px!important;
}
[data-testid="stSidebar"] *{color:#B8CCE0!important;}
[data-testid="stSidebar"] strong,[data-testid="stSidebar"] b{color:#E2EDF8!important;}
[data-testid="stSidebar"] h3{
    color:#4E7FA8!important; font-size:.60rem!important;
    letter-spacing:.22em!important; text-transform:uppercase!important;
    font-weight:700!important; margin:12px 0 4px 16px!important;
}
[data-testid="stSidebar"] hr{border-color:#263447!important;}
[data-testid="stSidebar"] code{
    background:#111C2A!important; color:#7EC8E3!important;
    border:1px solid #2A4060!important; padding:1px 5px!important;
    border-radius:3px!important; font-size:.78rem!important;
}
[data-testid="stSidebar"] [data-testid="stExpander"]{
    background:#111C2A!important; border:1px solid #263447!important;
    border-radius:5px!important;
}
[data-testid="stSidebar"] [data-testid="stExpander"] summary{
    color:#7AAFD4!important; font-size:.83rem!important; font-weight:600!important;
}
[data-testid="stSidebar"] [data-testid="stExpander"] .stMarkdown p,
[data-testid="stSidebar"] [data-testid="stExpander"] .stMarkdown li{
    font-size:.81rem!important; line-height:1.8!important; color:#90B0C8!important;
}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"]{
    background:#111C2A!important; border:1px dashed #3A5A7A!important;
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
    background:transparent!important; color:#90A8C0!important;
    transition:background .12s,color .12s;
}
[data-testid="stSidebar"] .stButton>button:hover{
    background:rgba(100,181,246,.07)!important; color:#B8D4EC!important;
}
[data-testid="stSidebar"] .stButton>button[kind="primary"]{
    background:rgba(21,101,192,.20)!important;
    border-left:3px solid #1565C0!important;
    color:#64B5F6!important; font-weight:700!important;
}
[data-testid="stSidebar"] .stButton>button:disabled{
    opacity:.35!important; cursor:not-allowed!important;
}

/* ── メインエリア ボタン ─────────────────────────────────── */
.stButton>button[kind="primary"]{
    background:#1565C0; color:#FFF;
    border:none; border-radius:5px; font-weight:600;
}
.stButton>button[kind="primary"]:hover{background:#0D47A1;}
.stButton>button[kind="primary"]:disabled{background:#999!important;color:#CCC!important;}

/* ── ダウンロード ─────────────────────────────────────────── */
[data-testid="stDownloadButton"]>button{
    background:#FFF!important; color:#1565C0!important;
    border:1.5px solid #1565C0!important; border-radius:5px!important;
    font-weight:600!important;
}
[data-testid="stDownloadButton"]>button:hover{background:#E3F2FD!important;}

/* ── info ───────────────────────────────────────────────── */
[data-testid="stInfo"]{
    background:#E8F0FE; border-left:3px solid #1565C0;
    border-radius:4px; color:#1A2B3C;
}
[data-testid="stRadio"] label,[data-testid="stCheckbox"] label{font-size:.84rem;}
hr{border-color:#E2E6EA!important;}

/* ── ページタイトルカード ────────────────────────────────── */
.page-card{
    background:#FFF; border:1px solid #E2E6EA; border-radius:8px;
    padding:14px 20px; margin-bottom:16px;
    box-shadow:0 1px 4px rgba(0,0,0,.05);
}
.page-card-title{font-size:1.0rem;font-weight:800;color:#1A2332;margin-bottom:2px;}
.page-card-sub{font-size:.74rem;color:#888;}

/* ── メトリクスカード ────────────────────────────────────── */
.metrics-row{display:flex;gap:10px;margin-bottom:12px;}
.m-card{
    flex:1; background:#FFF; border:1px solid #E2E6EA;
    border-radius:7px; padding:10px 16px;
    box-shadow:0 1px 3px rgba(0,0,0,.04);
    display:flex; align-items:center; gap:12px;
}
.m-val{font-size:2rem;font-weight:800;line-height:1;}
.m-lbl{font-size:.70rem;font-weight:600;color:#888;
        text-transform:uppercase;letter-spacing:.07em;}
.m-card.kaku{border-left:3px solid #34A853;} .m-card.kaku .m-val{color:#1B6E2A;}
.m-card.yo  {border-left:3px solid #F59E0B;} .m-card.yo   .m-val{color:#B45309;}
.m-card.mi  {border-left:3px solid #CCC;}    .m-card.mi   .m-val{color:#9E9E9E;}

/* ── 候補パネル ──────────────────────────────────────────── */
.cand-panel{
    background:#FFFBEB; border:1.5px solid #F59E0B;
    border-radius:8px; padding:16px 20px; margin-bottom:12px;
}
.cand-hdr{font-size:.90rem;font-weight:700;color:#92400E;
          margin-bottom:14px;display:flex;align-items:center;gap:8px;}
.cand-card{background:#FFF;border:1px solid #E2E6EA;border-radius:6px;padding:14px;}
.cand-card.sel{border-color:#1565C0;background:#EEF6FF;}
.cand-card-title{font-size:.88rem;font-weight:700;color:#1A2332;margin-bottom:8px;}
.cand-card-body{font-size:.79rem;color:#555;line-height:1.9;}
.diff-chip{
    display:inline-block;background:#DBEAFE;color:#1D4ED8;
    border-radius:3px;padding:0 6px;font-size:.75rem;font-weight:600;
}
.cand-foot{font-size:.72rem;color:#AAA;margin-top:10px;}

/* ── 凡例 ────────────────────────────────────────────────── */
.legend{display:flex;gap:14px;font-size:.73rem;color:#666;
        align-items:center;margin-top:6px;}
.ldot{width:9px;height:9px;border-radius:2px;
      display:inline-block;margin-right:3px;vertical-align:middle;}

/* ── 出力サマリー（サイドバー内） ────────────────────────── */
.out-summary{
    background:#162B40; border:1px solid #2A4A60;
    border-radius:6px; padding:10px 14px; margin-bottom:8px;
    font-size:.82rem;
}
.out-summary table{width:100%;border-collapse:collapse;}
.out-summary td{padding:3px 0;color:#90B8D0;}
.out-summary td.n{text-align:right;font-weight:700;color:#64B5F6;}
.out-summary .note{font-size:.70rem;color:#5A7A90;margin-top:6px;
                   border-top:1px solid #2A4060;padding-top:5px;}

/* ── sublabel ────────────────────────────────────────────── */
.sublabel{border-bottom:1px solid #1565C0;padding-bottom:2px;
          font-weight:700;font-size:.82rem;color:#333;margin-bottom:6px;}

/* ── 構造化ツリー行 ─────────────────────────────────────── */
.tree-row{
    padding:5px 12px; border-bottom:1px solid #F0F2F5;
    font-size:.84rem; color:#333; display:flex; align-items:center; gap:6px;
}
.tree-row:hover{background:#F5F7FA;}
.status-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0;}

/* ── 候補ナビ（前/次） ───────────────────────────────────── */
.cand-nav{
    display:flex; align-items:center; justify-content:space-between;
    padding:8px 0; margin-bottom:10px; font-size:.82rem; color:#555;
}
</style>
""", unsafe_allow_html=True)

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
for _k in ["suryo_info","df_match","selected_idx"]:
    if _k not in st.session_state: st.session_state[_k] = None
if "row_selections" not in st.session_state: st.session_state["row_selections"] = {}
if "page"           not in st.session_state: st.session_state["page"]           = "upload"
if "excel_cache"    not in st.session_state: st.session_state["excel_cache"]    = None
if "excel_fname"    not in st.session_state: st.session_state["excel_fname"]    = None

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
    if _chain_key(row) in st.session_state.row_selections: return "確定"
    return "要選択"

STATUS_BG = {"確定":"#F0FDF4","要選択":"#FFFBEB","未マッチ":"#F9FAFB"}

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
        '<div style="padding:16px 16px 14px;border-bottom:1px solid #263447;'
        'margin-bottom:4px;">'
        '<div style="font-size:.90rem;font-weight:800;color:#D4E8F8;'
        'letter-spacing:.03em;">施工管理計画</div>'
        '<div style="font-size:.68rem;color:#4E7FA8;margin-top:3px;">'
        'Automated Planning System</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    # ─ ワークフロータブ ──────────────────────────────────────
    st.markdown("### WORKFLOW")

    page      = st.session_state.get("page","upload")
    has_data  = st.session_state.df_match is not None
    has_sel   = (st.session_state.selected_idx is not None
                 and has_data
                 and st.session_state.selected_idx < len(st.session_state.df_match))

    NAV = [
        ("upload",    "① 取込",      True),
        ("structure", "② 構造化",    has_data),
        ("matching",  "③ マッチング", has_data),
    ]
    for key, label, enabled in NAV:
        # candidate ページにいるときは matching をアクティブ扱い
        is_active = (page == key) or (page == "candidate" and key == "matching")
        btn_type = "primary" if is_active else "secondary"
        if st.button(label, use_container_width=True,
                     type=btn_type, disabled=not enabled, key=f"nav_{key}"):
            st.session_state.page = key
            st.rerun()

    st.divider()

    # ─ その他ナビ ────────────────────────────────────────────
    if st.button("🗄  基準DB確認", use_container_width=True,
                 type="primary" if page=="db_view" else "secondary", key="nav_db"):
        st.session_state.page = "db_view"; st.rerun()

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
        st.session_state.row_selections = {}
        st.session_state.excel_cache = st.session_state.excel_fname = None
        st.session_state.page = "upload"
        st.rerun()

    with st.expander("？  使い方"):
        st.markdown("""
**①取込** PDFをアップロードして解析

**②構造化** 抽出された工種を確認

**③マッチング** 国交省基準DBとの対応を確認・行クリックで候補選択へ

**④候補選択** 複数候補がある工種の工法を確定

最後に「施工管理計画を出力」でExcel生成

---
基準DB更新時は `build_db.py` を再実行
""")

# ===========================================================================
# ① 取込ページ
# ===========================================================================
def _render_upload():
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
            st.session_state.excel_cache = None
            st.rerun()
        if st.button("→ ② 構造化を確認する", type="primary", key="go_structure"):
            st.session_state.page = "structure"; st.rerun()
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
                    st.session_state.selected_idx   = None
                    st.session_state.row_selections = {}
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
    st.markdown(
        f'<div class="page-card">'
        f'<div class="page-card-title">② 構造化 — 抽出結果の確認</div>'
        f'<div class="page-card-sub">{si.get("工事名","") if si else ""} '
        f'　抽出行数：{len(st.session_state.df_match)} 行</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    df_raw = _get_df_raw()

    DOT_COLOR = {"確定":"#34A853","要選択":"#F59E0B","未マッチ":"#CCC"}

    # テーブルとして表示
    df_disp = pd.DataFrame({
        "工種・種別・細別": ["　"*row["_depth"]+row["_name"] for _,row in df_raw.iterrows()],
        "マッチ状態":       [row["状態"] for _,row in df_raw.iterrows()],
    })
    sts_idx = {i: row["状態"] for i,(_,row) in enumerate(df_raw.iterrows())}

    STATUS_BG2 = {"確定":"#F0FDF4","要選択":"#FFFBEB","未マッチ":"#F9FAFB"}
    def _rs(sts):
        def _s(row):
            bg = STATUS_BG2.get(sts.get(row.name,""),"")
            return [f"background-color:{bg}" if bg else "" for _ in row]
        return _s

    st.dataframe(
        df_disp.style.apply(_rs(sts_idx),axis=1),
        use_container_width=True, height=520, hide_index=True,
        column_config={
            "工種・種別・細別": st.column_config.TextColumn(width="large"),
            "マッチ状態":       st.column_config.TextColumn(width="small"),
        }
    )
    st.markdown(
        '<div class="legend">'
        '<span><span class="ldot" style="background:#1B6E2A"></span>確定</span>'
        '<span><span class="ldot" style="background:#B45309"></span>要選択</span>'
        '<span><span class="ldot" style="background:#9E9E9E"></span>未マッチ</span>'
        '</div>',
        unsafe_allow_html=True,
    )
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    if st.button("→ ③ マッチング結果を確認する", type="primary", key="go_matching"):
        st.session_state.page = "matching"; st.rerun()

# ===========================================================================
# ③ マッチング＋候補選択（同一ページ）
# ===========================================================================
def _render_matching():
    df_raw = _get_df_raw()
    n_kaku = int((df_raw["状態"]=="確定").sum())
    n_yo   = int((df_raw["状態"]=="要選択").sum())
    n_mi   = int((df_raw["状態"]=="未マッチ").sum())

    st.markdown(
        '<div class="page-card">'
        '<div class="page-card-title">③ マッチング — 国交省基準DBとの対応確認</div>'
        '<div class="page-card-sub">行をクリックすると下部に候補選択パネルが展開されます</div>'
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

    # フィルタ
    filter_opt = st.radio(
        "filter", ["すべて","要選択のみ","確定のみ","未マッチのみ"],
        horizontal=True, label_visibility="collapsed",
    )
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

    sel_idx = st.session_state.selected_idx
    has_sel = sel_idx is not None and 0 <= sel_idx < len(df_raw)
    tbl_h   = 230 if has_sel else 430   # 候補パネル表示時はテーブルを縮める

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
        '<span><span class="ldot" style="background:#1B6E2A"></span>確定</span>'
        '<span><span class="ldot" style="background:#B45309"></span>要選択（行クリックで候補展開）</span>'
        '<span><span class="ldot" style="background:#9E9E9E"></span>未マッチ</span>'
        '</div>',
        unsafe_allow_html=True,
    )

    # ── 出力セクション（テーブル直下・常に表示） ─────────────
    _render_output_section()

    st.markdown("---")

    # ── 候補選択パネル ──────────────────────────────────────
    if not has_sel:
        st.markdown(
            '<div style="margin-top:10px;padding:18px;background:#FFF;'
            'border:1px dashed #D8DCE4;border-radius:8px;text-align:center;'
            'color:#AAAAAA;font-size:.84rem;">'
            '行をクリックすると候補がここに展開されます</div>',
            unsafe_allow_html=True,
        )
        return

    sel    = df_raw.iloc[sel_idx]
    ckey   = _chain_key(sel)
    chain  = " › ".join(sel[c] for c in SURYO_LEVEL_COLS if sel.get(c,""))

    # 要選択の前後ナビ用
    yo_idxs = [i for i,(_,r) in enumerate(df_raw.iterrows()) if r["状態"]=="要選択"]
    cur_pos = yo_idxs.index(sel_idx) if sel_idx in yo_idxs else None

    items_d = [x.strip() for x in str(sel.get("出来形マッチ","")).split("\n") if x.strip()]
    items_h = [x.strip() for x in str(sel.get("品質管理マッチ","")).split("\n") if x.strip()]
    items_p = [x.strip() for x in str(sel.get("撮影箇所マッチ","")).split("\n") if x.strip()]

    if not items_d and not items_h and not items_p:
        st.markdown(
            f'<div style="margin-top:10px;padding:14px;background:#F9FAFB;'
            f'border:1px solid #DDD;border-radius:8px;font-size:.84rem;color:#888;">'
            f'「{sel["_name"]}」はDBマッチなし（未マッチ）</div>',
            unsafe_allow_html=True,
        )
        return

    saved = st.session_state.row_selections.get(ckey)

    # 進捗（要選択行のみ）
    if cur_pos is not None and yo_idxs:
        done  = sum(1 for i in yo_idxs
                    if _chain_key(df_raw.iloc[i]) in st.session_state.row_selections)
        total = len(yo_idxs)
        prog_col, nav_col = st.columns([5, 1])
        with prog_col:
            st.progress(done/total, text=f"要選択 {total} 件中 {done} 件確定済み")
        with nav_col:
            if cur_pos < len(yo_idxs)-1:
                if st.button("次へ →", key="cand_next"):
                    st.session_state.selected_idx = yo_idxs[cur_pos+1]; st.rerun()

    st.markdown(
        f'<div class="cand-panel">'
        f'<div class="cand-hdr">⚠ {sel["_name"]}'
        f'<span style="font-weight:400;font-size:.76rem;margin-left:8px;">{chain}</span>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # 出来形 比較カード
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
                brd    = "border:1.5px solid #1565C0;background:#EEF6FF;" if is_sel else ""
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
    else:
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

    if items_d:
        fd = items_d[0].split(" / ")[0]
        dr = kojyo_data["出来形管理"][kojyo_data["出来形管理"]["工種"]==fd]
        if not dr.empty:
            r = dr.iloc[0]
            bc = " › ".join(x for x in [r.get("編",""),r.get("章",""),r.get("節",""),fd] if x)
            st.caption(f"DB 目次：{bc}")


def _render_output_section():
    """施工管理計画の出力ボタン・ダウンロードボタンを描画する。"""
    if st.session_state.df_match is None:
        return

    st.divider()
    df_tmp = _get_df_raw()
    out_d_l, out_h_l, out_p_l = _collect_labels(df_tmp)
    can_out = bool(out_d_l or out_h_l or out_p_l)

    st.markdown(
        f'<div style="background:#F8FAFC;border:1px solid #E2E6EA;border-radius:8px;'
        f'padding:16px 20px;margin-top:8px;">'
        f'<div style="font-size:.85rem;font-weight:700;color:#1A2332;margin-bottom:10px;">'
        f'出力内容</div>'
        f'<div style="font-size:.82rem;color:#555;line-height:2.0;">'
        f'出来形管理：<b>{len(out_d_l)}</b> 件　'
        f'品質管理：<b>{len(out_h_l)}</b> 件　'
        f'撮影箇所：<b>{len(out_p_l)}</b> 件'
        f'</div>'
        f'<div style="font-size:.72rem;color:#AAA;margin-top:4px;">'
        f'未確認の要選択は全候補を自動採用</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    col_out, col_dl = st.columns([2, 1])
    with col_out:
        if st.button("施工管理計画を出力", type="primary",
                     use_container_width=True, disabled=not can_out, key="btn_out"):
            try:
                si = st.session_state.suryo_info
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


# ===========================================================================
# 基準DB確認
# ===========================================================================
def _render_db_view():
    st.markdown(
        '<div class="page-card">'
        '<div class="page-card-title">🗄 国交省基準 DB</div>'
        f'<div class="page-card-sub">Ver. {version_info.get("バージョン","不明")}　'
        f'{version_info.get("作成日時","")}</div>'
        '</div>',
        unsafe_allow_html=True,
    )
    tab_d,tab_h,tab_p = st.tabs(["出来形管理","品質管理","撮影箇所"])
    with tab_d:
        st.dataframe(kojyo_data["出来形管理"], use_container_width=True, height=560, hide_index=True)
    with tab_h:
        st.dataframe(kojyo_data["品質管理"], use_container_width=True, height=560, hide_index=True)
    with tab_p:
        st.dataframe(kojyo_data["撮影箇所"], use_container_width=True, height=560, hide_index=True)

# ===========================================================================
# ルーティング
# ===========================================================================
page = st.session_state.page

if page == "db_view":
    _render_db_view()
elif page == "structure" and has_data:
    _render_structure()
elif page in ("matching", "candidate") and has_data:
    _render_matching()
else:
    _render_upload()
