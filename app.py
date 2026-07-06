# -*- coding: utf-8 -*-
"""
VIP 주간 실적 대시보드 (LF CRM/VIP)
주간회의 엑셀 'Summary' 시트 2.실적 양식을 그대로 재현 — 수기 입력만 자동화.
표 구조: 구분(지표/채널/카테고리) × [2026년 | 전년비 | 2025년]
- 시드: data/perf_long.csv (convert.ps1 산출, tidy long)
- 사이드바에서 원본(또는 update.ps1이 만든 CSV) 업로드 시 즉시 재계산
"""
import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="VIP 주간 실적", page_icon="📊", layout="wide")

SEED_CSV = "data/perf_long.csv"

# ----------------------------------------------------------------------------- parsing
def _year_of(x):
    if x is None:
        return None
    m = re.match(r"\s*(20\d{2})", str(x))
    if m and 2000 <= int(m.group(1)) <= 2100:
        return int(m.group(1))
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
    """Parse one BI-export xlsx (bytes) into tidy records. Mirrors convert.ps1."""
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    ws = wb.worksheets[0]
    grid = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    if not grid:
        return []
    nrows = len(grid)
    ncols = max(len(r) for r in grid)

    def cell(r, c):
        row = grid[r - 1] if r - 1 < nrows else []
        return row[c - 1] if c - 1 < len(row) else None

    val_start = 0
    for c in range(1, ncols + 1):
        if _year_of(cell(1, c)) is not None:
            val_start = c
            break
    if val_start == 0:
        return []

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

# 유효회원수는 55~59k로 안정적인 스냅샷 지표. 원본(일·주·월 전 export)의 2025-09말~10월
# 구간은 이 값이 ~112k로 배증하는 데이터 손상이 있고, 같은 구간의 거래액/DAU도 함께 부풀려짐.
# → 유효회원수(TOTAL)가 연중앙값의 1.5배를 넘는 (year,period)를 손상으로 보고 전 지표에서 제외.
EFF_METRIC = "유효회원수"
# 월별은 손상 많은 월별 export 대신 깨끗한 일자별에서 집계(레벨=일평균, 비율=구성요소로 재계산).
MONTHLY_LEVEL = ["일평균거래액", "일평균고객수", "DAU", "유효회원수", "유입율"]


def _corrupt_keys(df):
    """손상 (grain,year,period): 유효회원수 또는 DAU(TOTAL)가 연중앙값의 1.5배를 넘는 구간.
    유효회원수는 55~59k 고정이라 배증이 확실한 신호, DAU는 손상 구간 경계 주차 보완용."""
    ck = set()
    for met in (EFF_METRIC, "DAU"):
        s = df[(df.perspective == "overall") & (df.metric == met) & (df.seg1 == "TOTAL")].copy()
        if s.empty:
            continue
        s["med"] = s.groupby(["grain", "year"])["value"].transform("median")
        bad = s[s["value"] > s["med"] * 1.5]
        ck |= set(zip(bad.grain, bad.year, bad.period))
    return ck


def _drop_corrupt(df):
    ck = _corrupt_keys(df)
    if not ck:
        return df, ck
    keys = zip(df["grain"], df["year"], df["period"])
    keep = [k not in ck for k in keys]
    return df[keep].copy(), ck


def derive_monthly(df):
    """깨끗한 일자별(overall)에서 월별을 집계. 손상일 제외 후 유효일<20인 월은 생략(→ '—')."""
    day = df[(df.grain == "day") & (df.perspective == "overall")].copy()
    if day.empty:
        return None
    day["ym"] = day["period_sort"].astype(str).str[:6]
    rows = []
    for (yr, ym), g in day.groupby(["year", "ym"]):
        if g[g.metric == "DAU"]["value"].count() < 20:
            continue
        mo = int(ym[4:6])
        vals = {m: g[g.metric == m]["value"].mean() for m in MONTHLY_LEVEL}
        cust, dau, sales = vals.get("일평균고객수"), vals.get("DAU"), vals.get("일평균거래액")
        vals["일평균객단가"] = sales / cust if cust else None
        vals["CR"] = cust / dau if dau else None
        for m, v in vals.items():
            if v is None or pd.isna(v):
                continue
            rows.append(dict(grain="month", perspective="overall", year=int(yr), period=f"{mo}월",
                             period_sort=f"{yr}{mo:02d}00", seg1="TOTAL", seg2="",
                             metric=m, value=float(v), source="derived_from_daily"))
    return pd.DataFrame(rows) if rows else None


def finalize(df):
    if df is None or df.empty:
        return df
    df = df.copy()
    df["year"] = df["year"].astype(int)
    for c in ("seg1", "seg2"):
        df[c] = df[c].fillna("")
    ov = df["perspective"] == "overall"
    df.loc[ov & (df["seg1"] == ""), "seg1"] = "TOTAL"
    df = df.drop_duplicates(DEDUP_KEY, keep="last").reset_index(drop=True)
    df, _ = _drop_corrupt(df)
    dm = derive_monthly(df)
    if dm is not None and not dm.empty:
        df = df[~((df.grain == "month") & (df.perspective == "overall"))]
        df = pd.concat([df, dm], ignore_index=True)
    return df.reset_index(drop=True)


def _read_tidy_csv(name, data):
    df = pd.read_csv(io.BytesIO(data), encoding="utf-8-sig")
    need = [c for c in ("grain", "perspective", "year", "period", "metric", "value") if c not in df.columns]
    if need:
        raise ValueError(f"tidy CSV 컬럼 누락: {need}")
    for c in TIDY_COLS:
        if c not in df.columns:
            df[c] = ""
    return df[TIDY_COLS]


def parse_uploads(files):
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
            "🔒 DRM 보호(SCDSA) 파일이라 앱에서 직접 못 엽니다: " + ", ".join(drm)
            + "\n\n로컬에서 `update.ps1` 실행 → `data/perf_long.csv` 를 올려주세요.")
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
        return pd.DataFrame(columns=TIDY_COLS)


# ----------------------------------------------------------------------------- formatting
def fmt(metric, v):
    if v is None or pd.isna(v):
        return "-"
    if metric in ("유입율", "CR", "상품CR"):
        return f"{v*100:.1f}%"
    if metric == "일평균거래액":
        return f"{v/1e8:.2f}억"
    if metric == "일평균객단가":
        return f"{v:,.0f}"
    if metric in ("일평균고객수", "DAU", "유효회원수", "상품UV"):
        return f"{v:,.0f}"
    return f"{v:,.1f}"


def fmt_delta(metric, v):
    if v is None or pd.isna(v):
        return "-"
    if metric == "일평균거래액":
        return f"{v/1e8:+.2f}억"
    if metric in ("유입율", "CR", "상품CR"):
        return f"{v*100:+.1f}%p"
    return f"{v:+,.0f}"


def yoy(cur, prev):
    if prev in (None, 0) or pd.isna(prev) or cur is None or pd.isna(cur):
        return None
    return cur / prev - 1.0


def yoy_str(r):
    return "—" if r is None or pd.isna(r) else f"{r*100:+.1f}%"


def week_pretty(lbl):
    m = re.match(r"0?(\d{1,2})\D+?(\d{1,2})", str(lbl))
    return f"{int(m.group(1))}월 {int(m.group(2))}주" if m else str(lbl)


def month_pretty(lbl):
    m = re.match(r"0?(\d{1,2})", str(lbl))
    return f"{int(m.group(1))}월" if m else str(lbl)


# ----------------------------------------------------------------------------- data load
if "df" not in st.session_state:
    st.session_state.df = load_seed()

st.sidebar.title("📊 VIP 주간 실적")
st.sidebar.caption("주간회의 'Summary' 시트 2.실적 양식")

with st.sidebar.expander("🔄 데이터 업데이트", expanded=False):
    st.markdown(
        "**주간 갱신**: 로컬에서 `update.ps1` 실행 → `data/perf_long.csv` 생성 → 아래 업로드.\n\n"
        "원본 BI export는 **DRM(SCDSA)** 이라 앱에서 직접 못 엽니다. (비DRM xlsx는 바로 가능)")
    ups = st.file_uploader("CSV(권장) 또는 xlsx", type=["csv", "xlsx"], accept_multiple_files=True)
    if ups:
        newdf = parse_uploads(ups)
        if newdf is not None and not newdf.empty:
            st.session_state.df = newdf
            st.success(f"{len(ups)}개 파일 · {len(newdf):,}행 반영")
    if st.button("시드 데이터로 되돌리기"):
        load_seed.clear()
        st.session_state.df = load_seed()

df = st.session_state.df
if df is None or df.empty:
    st.warning("데이터가 없습니다. 사이드바에서 CSV/xlsx를 업로드하세요.")
    st.stop()

CUR = int(df["year"].max())
PREV = CUR - 1


# ----------------------------------------------------------------------------- accessors
def V(grain, perspective, metric, seg1, seg2, year, period):
    q = df[(df.grain == grain) & (df.perspective == perspective) & (df.metric == metric)
           & (df.seg1 == seg1) & (df.seg2 == seg2) & (df.year == year) & (df.period == period)]
    return q["value"].iloc[0] if len(q) else None


def periods(grain, perspective, metric, seg1, seg2, year):
    q = df[(df.grain == grain) & (df.perspective == perspective) & (df.metric == metric)
           & (df.seg1 == seg1) & (df.seg2 == seg2) & (df.year == year)].sort_values("period_sort")
    seen, out = set(), []
    for p in q["period"]:
        if p not in seen:
            seen.add(p)
            out.append(p)
    return out


YOY_POS = "color:#1f5fbf;font-weight:600"   # 신장(+) 파랑
YOY_NEG = "color:#c0392b;font-weight:600"   # 역신장(−) 빨강

TABLE_CSS = """
<style>
.sumtbl{border-collapse:collapse;font-size:12.5px;white-space:nowrap;}
.sumtbl th,.sumtbl td{border:1px solid #d9d9d9;padding:4px 8px;text-align:right;}
.sumtbl th{background:#f2f5fa;color:#222;text-align:center;font-weight:600;}
.sumtbl td.rowh,.sumtbl th.rowh{position:sticky;left:0;background:#fafafa;text-align:left;font-weight:600;z-index:1;}
.sumtbl .grp2026{background:#eaf1fb;}
.sumtbl .grpyoy{background:#fff4ec;}
.sumtbl .grp2025{background:#f4f4f4;}
.sumwrap{overflow-x:auto;border:1px solid #e6e6e6;border-radius:6px;}
</style>
"""


def render_block_table(row_labels, blocks):
    """blocks = list of (group_title, group_css, [(col_label, [cell_html per row]) ...])."""
    html = ['<div class="sumwrap"><table class="sumtbl">']
    # header row 1: group titles
    html.append('<tr><th class="rowh" rowspan="2">구분</th>')
    for title, gcss, cols in blocks:
        html.append(f'<th colspan="{len(cols)}" class="{gcss}">{title}</th>')
    html.append('</tr>')
    # header row 2: column labels
    html.append('<tr>')
    for title, gcss, cols in blocks:
        for cl, _ in cols:
            html.append(f'<th class="{gcss}">{cl}</th>')
    html.append('</tr>')
    # body
    for i, rl in enumerate(row_labels):
        html.append(f'<tr><td class="rowh">{rl}</td>')
        for title, gcss, cols in blocks:
            for _, cells in cols:
                html.append(cells[i])
        html.append('</tr>')
    html.append('</table></div>')
    return "".join(html)


PERF_ROWS = [("거래액", "일평균거래액"), ("고객수", "일평균고객수"), ("DAU", "DAU"),
             ("유입률", "유입율"), ("CR", "CR"), ("객단가", "일평균객단가")]


def perf_table(grain, cur_periods, prev_periods, pretty):
    """월별/주차별 실적표: rows=지표, blocks=[2026 | 전년비 | 2025]. 전년비는 cur_periods 기준."""
    b2026, byoy, b2025 = [], [], []
    for p in cur_periods:
        cells = []
        for _, met in PERF_ROWS:
            cells.append(f'<td class="grp2026">{fmt(met, V(grain,"overall",met,"TOTAL","",CUR,p))}</td>')
        b2026.append((pretty(p), cells))
    for p in cur_periods:
        cells = []
        for _, met in PERF_ROWS:
            r = yoy(V(grain, "overall", met, "TOTAL", "", CUR, p), V(grain, "overall", met, "TOTAL", "", PREV, p))
            sty = "" if r is None else (YOY_NEG if r < 0 else YOY_POS)
            cells.append(f'<td class="grpyoy" style="{sty}">{yoy_str(r)}</td>')
        byoy.append((pretty(p), cells))
    for p in prev_periods:
        cells = []
        for _, met in PERF_ROWS:
            cells.append(f'<td class="grp2025">{fmt(met, V(grain,"overall",met,"TOTAL","",PREV,p))}</td>')
        b2025.append((pretty(p), cells))
    blocks = [(f"{CUR}년", "grp2026", b2026), ("전년비", "grpyoy", byoy), (f"{PREV}년", "grp2025", b2025)]
    return render_block_table([r for r, _ in PERF_ROWS], blocks)


CH_ROWS = [("TTL", "TOTAL"), ("직접", "직접"), ("광고", "광고"), ("EP", "EP"), ("PUSH", "PUSH"), ("제휴", "제휴")]


def channel_table(metric, wk_periods):
    """채널별 표(단일 지표): rows=채널, blocks=[2026 | 전년비 | 2025], 3블록 동일 주차 라벨."""
    b2026, byoy, b2025 = [], [], []
    for p in wk_periods:
        b2026.append((week_pretty(p), [f'<td class="grp2026">{fmt(metric, V("week","overall",metric,s1,"",CUR,p))}</td>' for _, s1 in CH_ROWS]))
    for p in wk_periods:
        cells = []
        for _, s1 in CH_ROWS:
            r = yoy(V("week", "overall", metric, s1, "", CUR, p), V("week", "overall", metric, s1, "", PREV, p))
            sty = "" if r is None else (YOY_NEG if r < 0 else YOY_POS)
            cells.append(f'<td class="grpyoy" style="{sty}">{yoy_str(r)}</td>')
        byoy.append((week_pretty(p), cells))
    for p in wk_periods:
        b2025.append((week_pretty(p), [f'<td class="grp2025">{fmt(metric, V("week","overall",metric,s1,"",PREV,p))}</td>' for _, s1 in CH_ROWS]))
    blocks = [(f"{CUR}년", "grp2026", b2026), ("전년비", "grpyoy", byoy), (f"{PREV}년", "grp2025", b2025)]
    return render_block_table([r for r, _ in CH_ROWS], blocks)


CATS_ORDER = ["골프", "남성", "여성", "슈즈", "잡화", "스포츠", "명품", "아웃도어", "리빙", "뷰티", "키즈"]
YEONG = ["e-영업1", "e-영업2", "e-영업3", "e-영업4"]


def product_table(metric, wk):
    """상품별 표: rows=영업>카테고리, cols=[26년 | 25년 | 전년비 | 증감]. 단일 주차."""
    rows, r26, r25, ryoy, rdlt = [], [], [], [], []
    for ye in YEONG:
        # 영업 소계
        c = V("week", "product", metric, ye, "TOTAL", CUR, wk)
        p = V("week", "product", metric, ye, "TOTAL", PREV, wk)
        rows.append(f"<b>{ye}</b>")
        r26.append(("<b>" + fmt(metric, c) + "</b>") if c is not None else "-")
        r25.append(("<b>" + fmt(metric, p) + "</b>") if p is not None else "-")
        rr = yoy(c, p)
        ryoy.append((rr, True))
        rdlt.append(fmt_delta(metric, (c - p) if (c is not None and p is not None) else None))
        cats = [ca for ca in CATS_ORDER
                if V("week", "product", metric, ye, ca, CUR, wk) is not None
                or V("week", "product", metric, ye, ca, PREV, wk) is not None]
        for ca in cats:
            c = V("week", "product", metric, ye, ca, CUR, wk)
            p = V("week", "product", metric, ye, ca, PREV, wk)
            rows.append("&nbsp;&nbsp;" + ca)
            r26.append(fmt(metric, c))
            r25.append(fmt(metric, p))
            ryoy.append((yoy(c, p), False))
            rdlt.append(fmt_delta(metric, (c - p) if (c is not None and p is not None) else None))
    # build table
    html = [TABLE_CSS, '<div class="sumwrap"><table class="sumtbl">']
    html.append(f'<tr><th class="rowh">구분</th><th class="grp2026">{CUR}년</th>'
                f'<th class="grp2025">{PREV}년</th><th class="grpyoy">전년비</th><th>증감</th></tr>')
    for i, rl in enumerate(rows):
        rr, bold = ryoy[i]
        sty = "" if rr is None else (YOY_NEG if rr < 0 else YOY_POS)
        ytxt = yoy_str(rr)
        if bold:
            ytxt = f"<b>{ytxt}</b>"
        html.append(f'<tr><td class="rowh">{rl}</td><td class="grp2026">{r26[i]}</td>'
                    f'<td class="grp2025">{r25[i]}</td><td class="grpyoy" style="{sty}">{ytxt}</td>'
                    f'<td>{rdlt[i]}</td></tr>')
    html.append('</table></div>')
    return "".join(html)


# ============================================================================= RENDER
st.markdown(TABLE_CSS, unsafe_allow_html=True)

wk_all = periods("week", "overall", "일평균거래액", "TOTAL", "", CUR)
latest_wk = wk_all[-1] if wk_all else None
st.title(f"■ {week_pretty(latest_wk) if latest_wk else ''} 마감 CRM_VIP 실적")
st.caption(f"기준연도 {CUR} · 전년 {PREV}  |  주간회의 Summary 시트 2.실적 양식 · 자동 집계")

nwk = st.slider("주차 표시 개수", 5, 16, 5, help="주차별·채널별 표에 보여줄 최근 주차 수 (엑셀 기본 5주)")
wk_periods = wk_all[-nwk:]

# ---- 2) 월별 ----
st.header("2) 월별")
mcur = periods("month", "overall", "일평균거래액", "TOTAL", "", CUR)
mprev = [f"{m}월" for m in range(1, 13)]  # 전년은 항상 1~12월(손상/결측 월은 '-')
st.markdown(perf_table("month", mcur, mprev, month_pretty), unsafe_allow_html=True)

# ---- 3) 주차별 ----
st.header("3) 주차별")
st.markdown(perf_table("week", wk_periods, wk_periods, week_pretty), unsafe_allow_html=True)

# ---- 4) 주차별·채널별 ----
st.header("4) 주차별·채널별")
for tag, met in [("① 거래액", "일평균거래액"), ("② DAU", "DAU"), ("③ CR", "CR")]:
    st.subheader(tag)
    st.markdown(channel_table(met, wk_periods), unsafe_allow_html=True)

# ---- 5) 상품별 ----
if not df[df.perspective == "product"].empty:
    st.header("5) 상품별 (e-영업 × 카테고리)")
    pwk_all = periods("week", "product", "일평균거래액", "e-영업1", "TOTAL", CUR)
    default_ix = len(pwk_all) - 1
    sel = st.selectbox("주차 선택", pwk_all[::-1], index=0,
                       format_func=week_pretty) if pwk_all else None
    if sel:
        st.subheader("① 거래액")
        st.markdown(product_table("일평균거래액", sel), unsafe_allow_html=True)
        st.subheader("② 상품UV")
        st.markdown(product_table("상품UV", sel), unsafe_allow_html=True)

with st.expander("ℹ️ 표 읽는 법 / 데이터"):
    st.markdown(
        f"- 각 표는 **구분 × [{CUR}년 | 전년비 | {PREV}년]** 구조로 엑셀 Summary 2.실적과 동일합니다.\n"
        "- **전년비**: 같은 월/주차 라벨을 올해 vs 전년으로 계산. <span style='color:#1f5fbf'>파랑=신장(+)</span> / "
        "<span style='color:#c0392b'>빨강=역신장(−)</span>.\n"
        "- 거래액=일평균거래액(억), 고객수·DAU=일평균, 유입률·CR=%. (유효회원수는 Summary처럼 숨김)\n"
        "- **월별은 깨끗한 일자별에서 집계**(손상 많은 월별 export 대체).\n"
        "- ⚠️ 원본 데이터에 **2025년 9월말~10월** 손상 구간(유효회원수·DAU·거래액이 ~2배)이 있어 자동 제외했습니다. "
        "해당 구간·전년비는 '—'로 표시됩니다(현재연도 2026은 정상).\n"
        "- 목표 대비 달성율(2-1)·행사별(2-6)은 원본 6종에 데이터가 없어 제외(목표 파일 주시면 추가).",
        unsafe_allow_html=True)
