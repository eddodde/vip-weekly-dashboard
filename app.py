# -*- coding: utf-8 -*-
"""
VIP 주간 실적 대시보드 (LF CRM/VIP)
- 시드: data/perf_long.csv (convert.ps1 산출, tidy long)
- 사이드바에서 원본 BI export xlsx(전체관점/상품관점 · 일/주/월) 업로드 시 즉시 재파싱
- 섹션: 핵심 KPI / 월별 / 주차별 / 채널별 / 상품별  (Summary 시트 2.실적 미러링)
"""
import io
import re
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from plotly.subplots import make_subplots

st.set_page_config(page_title="VIP 주간 실적 대시보드", page_icon="📈", layout="wide")

SEED_CSV = "data/perf_long.csv"

LEVEL_METRICS = ["일평균거래액", "일평균고객수", "DAU", "유효회원수", "일평균객단가"]
RATE_METRICS = ["유입율", "CR"]
SHARE_METRICS = ["거래액비중", "고객비중"]
CHANNEL_ORDER = ["직접", "광고", "EP", "PUSH", "제휴", "브랜드광고", "미디어커머스"]
PRODUCT_METRICS = ["일평균거래액", "일평균고객수", "일평균객단가", "상품UV", "상품CR"]
CH_COLORS = {"직접": "#2E5EAA", "광고": "#5B9BD5", "EP": "#70AD47",
             "PUSH": "#ED7D31", "제휴": "#A5A5A5", "브랜드광고": "#7030A0", "미디어커머스": "#C00000"}

# ----------------------------------------------------------------------------- parsing
def _year_of(x):
    if x is None:
        return None
    m = re.match(r"\s*(20\d{2})", str(x))
    if m:
        y = int(m.group(1))
        if 2000 <= y <= 2100:
            return y
    return None


def _is_num(x):
    return isinstance(x, (int, float)) and not isinstance(x, bool)


def _grain(label):
    if label is None:
        return "unknown"
    s = str(label)
    if "/" in s:
        return "day"
    if re.search(r"\d+\D+\d+", s):
        return "week"
    if re.search(r"\d+", s):
        return "month"
    return "unknown"


def _sort_key(grain, year, label):
    s = str(label)
    if grain == "day":
        m = re.search(r"(\d{1,2})\s*/\s*(\d{1,2})", s)
        if m:
            return f"{year}{int(m.group(1)):02d}{int(m.group(2)):02d}"
    elif grain == "week":
        m = re.search(r"(\d{1,2})\D+(\d{1,2})", s)
        if m:
            return f"{year}{int(m.group(1)):02d}{int(m.group(2)):02d}"
    elif grain == "month":
        m = re.search(r"(\d{1,2})", s)
        if m:
            return f"{year}{int(m.group(1)):02d}00"
    return f"{year}9999"


def _norm_seg(x):
    if x is None:
        return ""
    s = str(x).strip()
    if s in ("-", ""):
        return ""
    if "TOTAL" in s.upper():
        return "TOTAL"
    return s


def parse_workbook(name, data):
    """Parse one BI-export xlsx (bytes) into a list of tidy records. Mirrors convert.ps1."""
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    ws = wb.worksheets[0]
    grid = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    if not grid:
        return []
    nrows = len(grid)
    ncols = max(len(r) for r in grid)

    def cell(r, c):  # 1-based
        row = grid[r - 1] if r - 1 < nrows else []
        return row[c - 1] if c - 1 < len(row) else None

    # value-start col = first col on row1 with a year marker
    val_start = 0
    for c in range(1, ncols + 1):
        if _year_of(cell(1, c)) is not None:
            val_start = c
            break
    if val_start == 0:
        return []

    # year forward-fill + labels(row3)
    year_of, label_of = {}, {}
    cur = None
    for c in range(val_start, ncols + 1):
        y = _year_of(cell(1, c))
        if y is not None:
            cur = y
        year_of[c] = cur
        label_of[c] = cell(3, c)

    grain = "unknown"
    for c in range(val_start, ncols + 1):
        if label_of[c] is not None:
            grain = _grain(label_of[c])
            break

    perspective = "product" if (val_start - 1) >= 7 else "overall"

    out = []
    cur_metric, cur_seg1 = None, None
    for r in range(2, nrows + 1):
        if not any(_is_num(cell(r, c)) for c in range(val_start, ncols + 1)):
            continue
        m = cell(r, 1)
        if m is not None:
            cur_metric = str(m).strip()
        if cur_metric is None:
            continue
        if perspective == "product":
            s1 = cell(r, 2)
            if s1 is not None:
                cur_seg1 = _norm_seg(s1)
            seg1 = cur_seg1 or ""
            seg2 = _norm_seg(cell(r, 3))
        else:
            seg1 = _norm_seg(cell(r, 2))
            seg2 = ""
        for c in range(val_start, ncols + 1):
            val = cell(r, c)
            if not _is_num(val):
                continue
            yr, lab = year_of[c], label_of[c]
            if yr is None or lab is None:
                continue
            labs = str(lab).strip()
            out.append(dict(grain=grain, perspective=perspective, year=yr,
                            period=labs, period_sort=_sort_key(grain, yr, labs),
                            seg1=seg1, seg2=seg2, metric=cur_metric,
                            value=float(val), source=name))
    return out


TIDY_COLS = ["grain", "perspective", "year", "period", "period_sort",
             "seg1", "seg2", "metric", "value", "source"]
DEDUP_KEY = ["grain", "perspective", "year", "period", "seg1", "seg2", "metric"]


def finalize(df):
    """Normalize types and drop duplicate period/seg/metric rows (keep last vintage)."""
    if df is None or df.empty:
        return df
    df = df.copy()
    df["year"] = df["year"].astype(int)
    for c in ("seg1", "seg2"):
        df[c] = df[c].fillna("")
    # overall: a blank channel means the whole aggregate — unify with "TOTAL" so the
    # total-only files (seg1="") and the channel file (seg1="TOTAL") collapse to one row.
    ov = df["perspective"] == "overall"
    df.loc[ov & (df["seg1"] == ""), "seg1"] = "TOTAL"
    df = df.drop_duplicates(DEDUP_KEY, keep="last").reset_index(drop=True)
    return df


def _read_tidy_csv(name, data):
    """Read an already-decrypted tidy long CSV (convert.ps1 산출)."""
    df = pd.read_csv(io.BytesIO(data), encoding="utf-8-sig")
    missing = [c for c in ("grain", "perspective", "year", "period", "metric", "value") if c not in df.columns]
    if missing:
        raise ValueError(f"tidy CSV 컬럼 누락: {missing}")
    for c in TIDY_COLS:
        if c not in df.columns:
            df[c] = ""
    df["seg1"] = df["seg1"].fillna("")
    df["seg2"] = df["seg2"].fillna("")
    return df[TIDY_COLS]


def parse_uploads(files):
    """Accept tidy CSV (primary, DRM-safe) and/or non-DRM xlsx. Returns combined DataFrame."""
    frames, recs, drm = [], [], []
    for f in files:
        low = f.name.lower()
        try:
            if low.endswith(".csv"):
                frames.append(_read_tidy_csv(f.name, f.getvalue()))
            else:
                recs.extend(parse_workbook(f.name, f.getvalue()))
        except Exception as e:  # noqa
            if "not a zip" in str(e).lower() or "badzip" in str(e).lower():
                drm.append(f.name)
            else:
                st.sidebar.error(f"파싱 실패: {f.name} — {e}")
    if drm:
        st.sidebar.warning(
            "🔒 DRM 보호 파일이라 앱에서 직접 열 수 없습니다: "
            + ", ".join(drm)
            + "\n\n로컬에서 `update.ps1`(또는 convert.ps1)을 실행해 만든 "
            + "`data/perf_long.csv`를 올려주세요. (Excel COM이 DRM을 복호화합니다)"
        )
    if recs:
        frames.append(pd.DataFrame(recs))
    if not frames:
        return None
    return finalize(pd.concat(frames, ignore_index=True))


@st.cache_data(show_spinner=False)
def load_seed():
    try:
        return finalize(pd.read_csv(SEED_CSV, encoding="utf-8-sig"))
    except Exception:
        return pd.DataFrame(columns=["grain", "perspective", "year", "period",
                                     "period_sort", "seg1", "seg2", "metric", "value", "source"])


# ----------------------------------------------------------------------------- helpers
def fmt(metric, v):
    if v is None or pd.isna(v):
        return "-"
    if metric in RATE_METRICS or metric in SHARE_METRICS or metric == "상품CR":
        return f"{v*100:.2f}%"
    if metric == "일평균거래액":
        return f"{v/1e8:.2f}억"
    if metric == "일평균객단가":
        return f"{v:,.0f}원"
    if metric in ("일평균고객수", "DAU", "유효회원수", "상품UV"):
        return f"{v:,.0f}"
    return f"{v:,.1f}"


def yoy(cur, prev):
    if prev in (None, 0) or pd.isna(prev) or cur is None or pd.isna(cur):
        return None
    return cur / prev - 1.0


def yoy_str(r):
    if r is None:
        return "—"
    return f"{r*100:+.1f}%"


def ordered_periods(sub):
    """unique period labels ordered by chronological sort (year-agnostic label)."""
    tmp = sub.sort_values("period_sort")
    seen, order = set(), []
    for p in tmp["period"]:
        if p not in seen:
            seen.add(p)
            order.append(p)
    return order


# ----------------------------------------------------------------------------- data load
seed = load_seed()
if "df" not in st.session_state:
    st.session_state.df = seed

st.sidebar.title("📈 VIP 주간 실적")
st.sidebar.caption("Summary 시트 2.실적 기반 · 전체/상품관점")

with st.sidebar.expander("🔄 데이터 업데이트", expanded=False):
    st.markdown(
        "**주간 갱신**: 로컬에서 `update.ps1` 실행 → `data/perf_long.csv` 생성 → 아래에 업로드.\n\n"
        "원본 BI export는 **DRM 보호(SCDSA)** 라 앱에서 직접 못 엽니다. "
        "(비DRM xlsx는 바로 업로드 가능)"
    )
    ups = st.file_uploader("CSV(권장) 또는 xlsx", type=["csv", "xlsx"], accept_multiple_files=True)
    if ups:
        newdf = parse_uploads(ups)
        if newdf is not None and not newdf.empty:
            st.session_state.df = newdf
            load_seed.clear()
            st.success(f"{len(ups)}개 파일 · {len(newdf):,}행 반영")
    if st.button("시드 데이터로 되돌리기"):
        st.session_state.df = load_seed()

df = st.session_state.df
if df is None or df.empty:
    st.warning("데이터가 없습니다. 사이드바에서 원본 xlsx를 업로드하세요.")
    st.stop()

CUR_YEAR = int(df["year"].max())
PREV_YEAR = CUR_YEAR - 1

st.sidebar.markdown("---")
page = st.sidebar.radio("메뉴", ["🏠 핵심 KPI", "🗓️ 월별", "📅 주차별", "🔀 채널별", "🛍️ 상품별"])
st.sidebar.markdown("---")
st.sidebar.caption(f"기준연도 {CUR_YEAR} · 전년 {PREV_YEAR}")


def get(grain, metric, seg1="TOTAL", seg2="", perspective="overall"):
    return df[(df.grain == grain) & (df.perspective == perspective) &
              (df.metric == metric) & (df.seg1 == seg1) & (df.seg2 == seg2)]


def how_to(text):
    with st.expander("ℹ️ 읽는 법"):
        st.markdown(text)


# ============================================================================= KPI
if page.startswith("🏠"):
    st.title("🏠 핵심 KPI")
    wk = get("week", "일평균거래액", "TOTAL")
    if wk.empty:
        st.info("주차 데이터가 없습니다.")
        st.stop()
    periods = ordered_periods(wk[wk.year == CUR_YEAR])
    latest = periods[-1] if periods else None
    st.subheader(f"최신 주차: {CUR_YEAR}년 {latest} (전년 동주 대비)")

    kpi_metrics = ["일평균거래액", "일평균고객수", "DAU", "유효회원수", "유입율", "CR", "일평균객단가"]
    cols = st.columns(len(kpi_metrics))
    for col, met in zip(cols, kpi_metrics):
        sub = get("week", met, "TOTAL")
        cur = sub[(sub.year == CUR_YEAR) & (sub.period == latest)]["value"]
        prev = sub[(sub.year == PREV_YEAR) & (sub.period == latest)]["value"]
        curv = cur.iloc[0] if len(cur) else None
        prevv = prev.iloc[0] if len(prev) else None
        col.metric(met, fmt(met, curv), yoy_str(yoy(curv, prevv)))

    st.markdown("---")
    st.subheader("최근 12주 추세 (거래액 · DAU)")
    c1, c2 = st.columns(2)
    for cc, met in ((c1, "일평균거래액"), (c2, "DAU")):
        sub = get("week", met, "TOTAL")
        order = ordered_periods(sub[sub.year == CUR_YEAR])[-12:]
        fig = go.Figure()
        for yr, color in ((PREV_YEAR, "#BBBBBB"), (CUR_YEAR, "#2E5EAA")):
            s = sub[sub.year == yr].set_index("period")["value"].reindex(order)
            fig.add_trace(go.Scatter(x=order, y=s.values, name=f"{yr}", mode="lines+markers",
                                     line=dict(color=color, width=3 if yr == CUR_YEAR else 2,
                                               dash="solid" if yr == CUR_YEAR else "dot")))
        fig.update_layout(title=met, height=340, margin=dict(t=40, b=40),
                          legend=dict(orientation="h", y=1.12))
        cc.plotly_chart(fig, use_container_width=True)
    how_to("- 실선=올해, 점선=전년. 같은 '월 N주차' 라벨끼리 전년비를 계산합니다.\n"
           "- 거래액=일평균거래액, DAU=일평균 방문자수.")

# ============================================================================= 월별
elif page.startswith("🗓️"):
    st.title("🗓️ 월별 트렌드")
    met = st.selectbox("지표", LEVEL_METRICS + RATE_METRICS, index=0)
    sub = get("month", met, "TOTAL")
    if sub.empty:
        st.info("월별 데이터가 없습니다.")
        st.stop()
    order = ordered_periods(sub)
    cur = sub[sub.year == CUR_YEAR].set_index("period")["value"].reindex(order)
    prev = sub[sub.year == PREV_YEAR].set_index("period")["value"].reindex(order)
    yy = [yoy(c, p) for c, p in zip(cur.values, prev.values)]

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(x=order, y=prev.values, name=f"{PREV_YEAR}", marker_color="#CFCFCF"), secondary_y=False)
    fig.add_trace(go.Bar(x=order, y=cur.values, name=f"{CUR_YEAR}", marker_color="#2E5EAA"), secondary_y=False)
    fig.add_trace(go.Scatter(x=order, y=[None if v is None else v*100 for v in yy], name="전년비(%)",
                             mode="lines+markers", line=dict(color="#ED7D31", width=2)), secondary_y=True)
    fig.update_layout(barmode="group", height=460, legend=dict(orientation="h", y=1.1),
                      title=f"{met} — {CUR_YEAR} vs {PREV_YEAR}")
    fig.update_yaxes(title_text=met, secondary_y=False)
    fig.update_yaxes(title_text="전년비(%)", secondary_y=True, zeroline=True, zerolinecolor="#ED7D31")
    st.plotly_chart(fig, use_container_width=True)

    tbl = pd.DataFrame({"월": order,
                        f"{CUR_YEAR}": [fmt(met, v) for v in cur.values],
                        f"{PREV_YEAR}": [fmt(met, v) for v in prev.values],
                        "전년비": [yoy_str(v) for v in yy]})
    st.dataframe(tbl, use_container_width=True, hide_index=True)
    how_to("- 막대=월별 값(회색 전년/파랑 올해), 주황선=전년비(우축, %).\n- 지표는 상단에서 선택.")

# ============================================================================= 주차별
elif page.startswith("📅"):
    st.title("📅 주차별 트렌드")
    c0, c1 = st.columns([2, 1])
    met = c0.selectbox("지표", LEVEL_METRICS + RATE_METRICS, index=0)
    nwk = c1.slider("최근 주차 수", 6, 26, 12)
    sub = get("week", met, "TOTAL")
    if sub.empty:
        st.info("주차 데이터가 없습니다.")
        st.stop()
    order = ordered_periods(sub[sub.year == CUR_YEAR])[-nwk:]
    cur = sub[sub.year == CUR_YEAR].set_index("period")["value"].reindex(order)
    prev = sub[sub.year == PREV_YEAR].set_index("period")["value"].reindex(order)
    yy = [yoy(c, p) for c, p in zip(cur.values, prev.values)]

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Scatter(x=order, y=prev.values, name=f"{PREV_YEAR}", mode="lines+markers",
                             line=dict(color="#BBBBBB", dash="dot")), secondary_y=False)
    fig.add_trace(go.Scatter(x=order, y=cur.values, name=f"{CUR_YEAR}", mode="lines+markers",
                             line=dict(color="#2E5EAA", width=3)), secondary_y=False)
    fig.add_trace(go.Bar(x=order, y=[None if v is None else v*100 for v in yy], name="전년비(%)",
                         marker_color="rgba(237,125,49,0.35)"), secondary_y=True)
    fig.update_layout(height=460, legend=dict(orientation="h", y=1.1),
                      title=f"{met} — 최근 {nwk}주 ({CUR_YEAR} vs {PREV_YEAR})")
    fig.update_yaxes(title_text=met, secondary_y=False)
    fig.update_yaxes(title_text="전년비(%)", secondary_y=True)
    st.plotly_chart(fig, use_container_width=True)

    tbl = pd.DataFrame({"주차": order,
                        f"{CUR_YEAR}": [fmt(met, v) for v in cur.values],
                        f"{PREV_YEAR}": [fmt(met, v) for v in prev.values],
                        "전년비": [yoy_str(v) for v in yy]})
    st.dataframe(tbl, use_container_width=True, hide_index=True)
    how_to("- 실선=올해, 점선=전년(같은 '월 N주차'), 막대=전년비(우축).")

# ============================================================================= 채널별
elif page.startswith("🔀"):
    st.title("🔀 채널별")
    c0, c1 = st.columns(2)
    grain = c0.radio("주기", ["week", "month"], format_func=lambda x: "주별" if x == "week" else "월별", horizontal=True)
    met = c1.selectbox("지표", ["일평균거래액", "일평균고객수", "DAU", "일평균객단가", "CR"], index=0)
    channels = [ch for ch in CHANNEL_ORDER if not get(grain, met, ch).empty]

    sub_all = df[(df.grain == grain) & (df.perspective == "overall") & (df.metric == met) & (df.seg2 == "")]
    order = ordered_periods(sub_all[sub_all.year == CUR_YEAR])[-(12 if grain == "week" else 12):]

    st.subheader(f"채널 구성 추세 ({CUR_YEAR})")
    stack = met in ("일평균거래액", "일평균고객수", "DAU")  # additive metrics → stacked
    fig = go.Figure()
    for ch in channels:
        s = get(grain, met, ch)
        s = s[s.year == CUR_YEAR].set_index("period")["value"].reindex(order)
        if stack:
            fig.add_trace(go.Bar(x=order, y=s.values, name=ch, marker_color=CH_COLORS.get(ch)))
        else:
            fig.add_trace(go.Scatter(x=order, y=s.values, name=ch, mode="lines+markers",
                                     line=dict(color=CH_COLORS.get(ch))))
    fig.update_layout(barmode="stack" if stack else "overlay", height=440,
                      legend=dict(orientation="h", y=1.1), title=f"{met} — 채널별")
    st.plotly_chart(fig, use_container_width=True)

    latest = order[-1] if order else None
    st.subheader(f"채널별 전년비 — 최신 {'주차' if grain=='week' else '월'} ({latest})")
    rows = []
    for ch in channels:
        s = get(grain, met, ch)
        cv = s[(s.year == CUR_YEAR) & (s.period == latest)]["value"]
        pv = s[(s.year == PREV_YEAR) & (s.period == latest)]["value"]
        cvv = cv.iloc[0] if len(cv) else None
        pvv = pv.iloc[0] if len(pv) else None
        rows.append({"채널": ch, f"{CUR_YEAR}": fmt(met, cvv), f"{PREV_YEAR}": fmt(met, pvv),
                     "전년비": yoy_str(yoy(cvv, pvv))})
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    how_to("- 가산 지표(거래액·고객수·DAU)는 누적 막대로 채널 기여를 봅니다.\n"
           "- 비율 지표(CR·객단가)는 채널별 선그래프.\n- 표는 최신 시점 채널별 전년비.")

# ============================================================================= 상품별
elif page.startswith("🛍️"):
    st.title("🛍️ 상품별 (e-영업 × 카테고리)")
    prod = df[df.perspective == "product"]
    if prod.empty:
        st.info("상품관점 데이터가 없습니다. 사이드바에서 '상품관점' xlsx를 올리세요.")
        st.stop()
    c0, c1 = st.columns(2)
    met = c0.selectbox("지표", ["일평균거래액", "상품UV", "상품CR", "일평균객단가", "일평균고객수"], index=0)
    sub = prod[prod.metric == met]
    order = ordered_periods(sub[sub.year == CUR_YEAR])
    latest = c1.selectbox("주차", order[::-1], index=0) if order else None

    yeongs = ["e-영업1", "e-영업2", "e-영업3", "e-영업4"]
    st.subheader(f"{met} 전년비 — {latest}")
    # heatmap: rows=영업, cols=category, value=전년비
    cats = [c for c in ["골프", "남성", "여성", "슈즈", "잡화", "스포츠", "명품", "아웃도어", "리빙", "뷰티", "키즈"]
            if c in sub["seg2"].unique()]
    z, text = [], []
    for ye in yeongs:
        zr, tr = [], []
        for ca in cats:
            cv = sub[(sub.year == CUR_YEAR) & (sub.period == latest) & (sub.seg1 == ye) & (sub.seg2 == ca)]["value"]
            pv = sub[(sub.year == PREV_YEAR) & (sub.period == latest) & (sub.seg1 == ye) & (sub.seg2 == ca)]["value"]
            r = yoy(cv.iloc[0] if len(cv) else None, pv.iloc[0] if len(pv) else None)
            zr.append(None if r is None else r*100)
            tr.append("" if r is None else f"{r*100:+.0f}%")
        z.append(zr)
        text.append(tr)
    fig = go.Figure(go.Heatmap(z=z, x=cats, y=yeongs, text=text, texttemplate="%{text}",
                               colorscale="RdBu", zmid=0, reversescale=True,
                               colorbar=dict(title="전년비%")))
    fig.update_layout(height=360, title=f"{met} 전년비 (%)")
    st.plotly_chart(fig, use_container_width=True)

    # top movers (absolute 전년비 on 거래액-like)
    st.subheader("전년비 급변 카테고리 TOP")
    movers = []
    for ye in yeongs:
        for ca in cats:
            cv = sub[(sub.year == CUR_YEAR) & (sub.period == latest) & (sub.seg1 == ye) & (sub.seg2 == ca)]["value"]
            pv = sub[(sub.year == PREV_YEAR) & (sub.period == latest) & (sub.seg1 == ye) & (sub.seg2 == ca)]["value"]
            r = yoy(cv.iloc[0] if len(cv) else None, pv.iloc[0] if len(pv) else None)
            if r is not None:
                movers.append({"영업": ye, "카테고리": ca, f"{CUR_YEAR}": fmt(met, cv.iloc[0]),
                               "전년비": r})
    md = pd.DataFrame(movers)
    if not md.empty:
        md = md.sort_values("전년비")
        cL, cR = st.columns(2)
        top = md.tail(5).iloc[::-1].copy(); top["전년비"] = top["전년비"].map(yoy_str)
        bot = md.head(5).copy(); bot["전년비"] = bot["전년비"].map(yoy_str)
        cL.caption("📈 신장 TOP5"); cL.dataframe(top, use_container_width=True, hide_index=True)
        cR.caption("📉 역신장 TOP5"); cR.dataframe(bot, use_container_width=True, hide_index=True)
    how_to("- 히트맵: 파랑=신장/빨강=역신장 (전년비 %).\n- 영업×카테고리별 전년비를 한눈에, 아래 표는 급변 TOP.")
