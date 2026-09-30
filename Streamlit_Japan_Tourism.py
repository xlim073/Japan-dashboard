"""
Japan resort-hotel destination screen (portfolio project, Module 1).

Question: which regional Japanese prefectures have growing, diversified, ski-season-strong
foreign demand that could support a new resort hotel?

Data: Japan Tourism Agency (観光庁), Overnight Travel Statistics Survey (宿泊旅行統計調査),
annual confirmed workbooks for 2025 and 2019, downloaded live from
https://www.mlit.go.jp/kankocho/tokei_hakusyo/shukuhakutokei.html

Run:  /usr/local/bin/python3 -m streamlit run "Streamlit_Japan_Tourism.py"
"""
import io

import numpy as np
import openpyxl
import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st

if __name__ == "__main__" and not st.runtime.exists():
    # Started with plain `python` (e.g. the VS Code ▶ button): relaunch properly with `streamlit run`.
    import os
    import sys
    os.execv(sys.executable, [sys.executable, "-m", "streamlit", "run", __file__])

st.set_page_config(page_title="Japan resort destination screen", page_icon="🏔️", layout="wide")

# ---------------------------------------------------------------------------
# Settings (edit these in the code; the dashboard has no controls on purpose)
# ---------------------------------------------------------------------------
LATEST, BASE = 2025, 2019          # 2019 = last year before the pandemic
URLS = {
    2025: "https://www.mlit.go.jp/kankocho/content/002010340.xlsx",
    2019: "https://www.mlit.go.jp/kankocho/tokei_hakusyo/content/001350484.xlsx",
}
# Prefecture outlines for the map (based on MLIT's National Land Numerical Information, 国土数値情報)
BOUNDARIES_URL = "https://raw.githubusercontent.com/dataofjapan/land/master/japan.geojson"
MIN_NIGHTS = 500_000              # prefectures smaller than this are too small for a luxury resort
METRO = {"Tokyo", "Kanagawa", "Chiba", "Saitama", "Aichi", "Osaka", "Kyoto", "Hyogo"}  # JTA's 3 metro areas
LONG_HAUL = ["米国", "アメリカ", "カナダ", "英国", "イギリス", "ドイツ", "フランス", "イタリア", "スペイン", "オーストラリア"]
SKI_MONTHS = [12, 1, 2, 3]
INDIGO, GREY = "#2a4f8f", "#9aa6b5"

# Japanese name, English name, capital city, latitude, longitude (capital city, used for the map)
PREFECTURES = [
    ("北海道", "Hokkaido", "Sapporo", 43.06, 141.35), ("青森県", "Aomori", "Aomori", 40.82, 140.74),
    ("岩手県", "Iwate", "Morioka", 39.70, 141.15), ("宮城県", "Miyagi", "Sendai", 38.27, 140.87),
    ("秋田県", "Akita", "Akita", 39.72, 140.10), ("山形県", "Yamagata", "Yamagata", 38.24, 140.36),
    ("福島県", "Fukushima", "Fukushima", 37.75, 140.47), ("茨城県", "Ibaraki", "Mito", 36.34, 140.45),
    ("栃木県", "Tochigi", "Utsunomiya", 36.57, 139.88), ("群馬県", "Gunma", "Maebashi", 36.39, 139.06),
    ("埼玉県", "Saitama", "Saitama", 35.86, 139.65), ("千葉県", "Chiba", "Chiba", 35.61, 140.12),
    ("東京都", "Tokyo", "Tokyo", 35.69, 139.69), ("神奈川県", "Kanagawa", "Yokohama", 35.45, 139.64),
    ("新潟県", "Niigata", "Niigata", 37.90, 139.02), ("富山県", "Toyama", "Toyama", 36.70, 137.21),
    ("石川県", "Ishikawa", "Kanazawa", 36.59, 136.63), ("福井県", "Fukui", "Fukui", 36.07, 136.22),
    ("山梨県", "Yamanashi", "Kofu", 35.66, 138.57), ("長野県", "Nagano", "Nagano", 36.65, 138.18),
    ("岐阜県", "Gifu", "Gifu", 35.39, 136.72), ("静岡県", "Shizuoka", "Shizuoka", 34.98, 138.38),
    ("愛知県", "Aichi", "Nagoya", 35.18, 136.91), ("三重県", "Mie", "Tsu", 34.73, 136.51),
    ("滋賀県", "Shiga", "Otsu", 35.00, 135.87), ("京都府", "Kyoto", "Kyoto", 35.02, 135.76),
    ("大阪府", "Osaka", "Osaka", 34.69, 135.52), ("兵庫県", "Hyogo", "Kobe", 34.69, 135.18),
    ("奈良県", "Nara", "Nara", 34.69, 135.83), ("和歌山県", "Wakayama", "Wakayama", 34.23, 135.17),
    ("鳥取県", "Tottori", "Tottori", 35.50, 134.24), ("島根県", "Shimane", "Matsue", 35.47, 133.05),
    ("岡山県", "Okayama", "Okayama", 34.66, 133.93), ("広島県", "Hiroshima", "Hiroshima", 34.40, 132.46),
    ("山口県", "Yamaguchi", "Yamaguchi", 34.19, 131.47), ("徳島県", "Tokushima", "Tokushima", 34.07, 134.56),
    ("香川県", "Kagawa", "Takamatsu", 34.34, 134.04), ("愛媛県", "Ehime", "Matsuyama", 33.84, 132.77),
    ("高知県", "Kochi", "Kochi", 33.56, 133.53), ("福岡県", "Fukuoka", "Fukuoka", 33.61, 130.42),
    ("佐賀県", "Saga", "Saga", 33.25, 130.30), ("長崎県", "Nagasaki", "Nagasaki", 32.74, 129.87),
    ("熊本県", "Kumamoto", "Kumamoto", 32.79, 130.74), ("大分県", "Oita", "Oita", 33.24, 131.61),
    ("宮崎県", "Miyazaki", "Miyazaki", 31.91, 131.42), ("鹿児島県", "Kagoshima", "Kagoshima", 31.56, 130.56),
    ("沖縄県", "Okinawa", "Naha", 26.21, 127.68),
]
JP_TO_EN = {p[0]: p[1] for p in PREFECTURES}

# ---------------------------------------------------------------------------
# Reading the JTA workbooks
# Every table has the same layout: prefecture names in column A (e.g. "01北海道"),
# column headers in rows 4-5. We pick columns by their header text.
# ---------------------------------------------------------------------------

def num(v) -> float:
    """Cell -> number. JTA marks some small values with '*' and missing ones with '-'."""
    try:
        return float(str(v).replace(",", "").lstrip("*"))
    except ValueError:
        return np.nan


def sum_columns(wb, sheet: str, keywords: list[str]) -> pd.Series:
    """Add up every column whose header contains one of the keywords, for each prefecture."""
    rows = [list(r) + [None] * 30 for r in wb[sheet].iter_rows(values_only=True)]  # pad short rows
    header = [f"{rows[3][c] or ''}{rows[4][c] or ''}".replace("\n", "") for c in range(len(rows[3]))]
    cols = [c for c, h in enumerate(header) if any(k in h for k in keywords)]
    out = {}
    for r in rows:
        name = str(r[0] or "").strip("　 ").lstrip("0123456789")
        if name in JP_TO_EN:
            out[JP_TO_EN[name]] = sum(num(r[c]) for c in cols)
    return pd.Series(out)


@st.cache_data(ttl=7 * 86400, show_spinner="Downloading Japan Tourism Agency data…")
def load_year(year: int) -> pd.DataFrame:
    raw = requests.get(URLS[year], headers={"User-Agent": "Mozilla/5.0"}, timeout=120).content
    wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
    return pd.DataFrame({
        "foreign nights": sum_columns(wb, "第2表(年計)", ["外国人"]),          # all foreign guest nights
        "by nationality": sum_columns(wb, "参考第1表(年計)", ["外国人"]),      # total of the nationality table
        "long haul": sum_columns(wb, "参考第1表(年計)", LONG_HAUL),
        "ski occupancy": pd.concat(  # resort-hotel room occupancy, average of Dec-Mar
            [sum_columns(wb, f"第8表({m}月)", ["リゾート"]) for m in SKI_MONTHS], axis=1).mean(axis=1),
    })


def simplify_ring(ring: list) -> list:
    """Round an outline's points to ~1 km and drop repeats: the file shrinks from 13 MB to 0.6 MB."""
    out = []
    for x, y in ring:
        p = [round(x, 2), round(y, 2)]
        if not out or out[-1] != p:
            out.append(p)
    return out


@st.cache_data(ttl=30 * 86400, show_spinner="Downloading prefecture boundaries…")
def load_boundaries() -> dict:
    """Prefecture outlines for the map, with each shape's id set to the English prefecture name."""
    geo = requests.get(BOUNDARIES_URL, timeout=120).json()
    for f in geo["features"]:
        f["id"] = JP_TO_EN[f["properties"]["nam_ja"]]
        g = f["geometry"]
        polygons = g["coordinates"] if g["type"] == "MultiPolygon" else [g["coordinates"]]
        polygons = [[r for r in map(simplify_ring, p) if len(r) >= 4] for p in polygons]
        g["type"], g["coordinates"] = "MultiPolygon", [p for p in polygons if p]
    return geo

# ---------------------------------------------------------------------------
# Metrics and score
# ---------------------------------------------------------------------------
GROWTH = f"Growth since {BASE} %"
# Higher is better for all three; the score treats them equally.
SCORED = [GROWTH, "Ski-season resort occupancy %", "Long-haul share %"]

try:
    now, then = load_year(LATEST), load_year(BASE)
    boundaries = load_boundaries()
except Exception as e:
    st.error(f"Could not download the data ({e}). Check your internet connection and reload.")
    st.stop()

df = pd.DataFrame(
    {"Capital": [p[2] for p in PREFECTURES], "lat": [p[3] for p in PREFECTURES], "lon": [p[4] for p in PREFECTURES]},
    index=[p[1] for p in PREFECTURES])
df["Foreign nights"] = now["foreign nights"]
df[GROWTH] = 100 * (now["foreign nights"] / then["foreign nights"] - 1)
df["Ski-season resort occupancy %"] = now["ski occupancy"]
df["Long-haul share %"] = 100 * now["long haul"] / now["by nationality"]

# Screen: regional prefectures with enough foreign demand. Score = average z-score of the three metrics.
screen = df[~df.index.isin(METRO) & (df["Foreign nights"] >= MIN_NIGHTS)].copy()
z = (screen[SCORED] - screen[SCORED].mean()) / screen[SCORED].std()
screen["Score"] = z.fillna(0).mean(axis=1)
screen = screen.sort_values("Score", ascending=False)
screen.insert(0, "Rank", range(1, len(screen) + 1))

# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------
st.title("Where should the next Japan resort hotel go?")
st.caption("Module 1 · Destination screen using Japan Tourism Agency Overnight Travel Statistics (宿泊旅行統計調査).")

c1, c2, c3 = st.columns(3)
c1.metric("Top-ranked prefecture", screen.index[0], f"then {', '.join(screen.index[1:3])}", delta_color="off")
c2.metric(f"Foreign guest nights in screened group, {LATEST}", f"{screen['Foreign nights'].sum() / 1e6:.1f}M")
c3.metric("Prefectures screened", len(screen))
st.info(f"{LATEST} data vs {BASE}. Screened: regional prefectures (excluding the Tokyo, Nagoya and Osaka metro "
        f"areas) with at least {MIN_NIGHTS:,} foreign guest nights.")

# ---- 1) Map ----
st.markdown("## 1) Map of Japan")
others = df[~df.index.isin(screen.index)]
hover = "<b>%{location}</b> (capital: %{customdata[0]})<br>%{customdata[1]:,.0f} foreign guest nights"
fig = go.Figure([
    # prefectures outside the screen: plain grey
    go.Choroplethmap(geojson=boundaries, locations=others.index, z=[0] * len(others), name="Not screened",
                     colorscale=[[0, GREY], [1, GREY]], showscale=False, showlegend=True, marker_opacity=0.2,
                     customdata=np.c_[others["Capital"], others["Foreign nights"]],
                     hovertemplate=hover + "<br>not screened<extra></extra>"),
    # screened prefectures: darker = higher score
    go.Choroplethmap(geojson=boundaries, locations=screen.index, z=screen["Score"], name="Screened",
                     colorscale=[[0, "#a9bcdf"], [1, INDIGO]], marker_opacity=0.85, marker_line_color="white",
                     colorbar=dict(title="Score", thickness=12),
                     customdata=np.c_[screen["Capital"], screen["Foreign nights"], screen["Rank"]],
                     hovertemplate=hover + "<br>rank %{customdata[2]} of " + f"{len(screen)}<extra></extra>"),
    # rank labels for the top 10 (labelling all 21 overlaps in the south-west)
    go.Scattermap(lat=screen["lat"][:10], lon=screen["lon"][:10], mode="text", text=[str(r) for r in screen["Rank"][:10]],
                  textfont=dict(size=13, color="black"), hoverinfo="skip", showlegend=False),
])
fig.update_layout(map=dict(style="carto-positron", center=dict(lat=36.0, lon=136.0), zoom=4.1), height=650,
                  margin=dict(t=10, b=10, l=10, r=10), legend=dict(orientation="h", y=-0.02))
st.plotly_chart(fig, width="stretch")
st.caption("Significance: screened prefectures are shaded by score (darker = better); numbers mark the top 10. "
           "Grey prefectures are metro areas or have under "
           f"{MIN_NIGHTS:,} foreign guest nights. Hover over any prefecture for details; scroll to zoom.")

# ---- 2) Ranking ----
st.markdown("## 2) Ranking")
left, right = st.columns(2)
with left:
    b = screen.iloc[::-1]
    fig2 = go.Figure(go.Bar(x=b["Score"], y=b.index, orientation="h",
                            marker_color=INDIGO))
    fig2.update_layout(title="Score (0 = average of the screened group)", height=480, margin=dict(t=40, l=10))
    st.plotly_chart(fig2, width="stretch")
    st.caption("Significance: the score averages three equally weighted measures (higher is better for each): "
               "foreign-demand growth, ski-season resort occupancy and long-haul guest share.")
with right:
    cols = ["Rank", "Capital", "Foreign nights", *SCORED, "Score"]
    st.dataframe(screen[cols].style.format({"Foreign nights": "{:,.0f}", **{c: "{:.1f}" for c in SCORED},
                                            "Score": "{:+.2f}"}, na_rep="–"),
                 width="stretch", height=480)
    st.caption("Significance: the raw numbers behind the score, so each ranking can be checked.")

# ---- Sources ----
with st.expander("Data sources and checks"):
    st.write("Japan Tourism Agency, Overnight Travel Statistics Survey, annual confirmed values: "
             "https://www.mlit.go.jp/kankocho/tokei_hakusyo/shukuhakutokei.html")
    st.write("Tables used: 第2表 (foreign guest nights), 参考第1表 (nationality, facilities with 10+ staff), "
             "第8表 Dec–Mar sheets (resort-hotel room occupancy).")
    st.write(f"Map outlines: {BOUNDARIES_URL} (derived from MLIT National Land Numerical Information), "
             "simplified to ~1 km detail.")
    st.write(f"Check: total foreign guest nights {LATEST} = {now['foreign nights'].sum():,.0f} "
             f"(JTA's published 2025 total is 179,922,130).")
st.caption("Source: Japan Tourism Agency. Screening tool, not investment advice.")
