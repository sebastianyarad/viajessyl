"""
Dashboard de Viajes · Sebastián Yarad Luco
Fuente: Google Drive (xlsx via API)   |   Acceso: st.secrets
"""
import streamlit as st
import pandas as pd
import datetime
import io
from googleapiclient.discovery import build
from google.oauth2 import service_account
import plotly.graph_objects as go
import plotly.express as px

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Viajes · SYL",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

BG     = "#0D1117"
CARD   = "#161B22"
BORD   = "#30363D"
BLUE   = "#58A6FF"
GREEN  = "#3FB950"
PURPLE = "#A371F7"
ORANGE = "#FFA657"
RED    = "#F85149"
TEXT   = "#E6EDF3"
MUTED  = "#8B949E"

REGION_MAP = [
    ("Piura",        "388BFD", ["PIU"]),
    ("USA / Canadá", "3FB950", ["USA", "PHL", "CAN", "CPMA", "HONEY", "BKSFLD", "IFPA"]),
    ("Europa",       "A371F7", ["BERLIN", "EUR", "ITGS"]),
    ("Asia / HK",    "F78166", ["AFL", "HK"]),
    ("Chile",        "FFA657", ["BBA", "SIX", "BALMACEDA"]),
]
REGION_COLORS = {label: "#" + col for label, col, _ in REGION_MAP}

st.markdown(f"""
<style>
[data-testid="stAppViewContainer"] {{ background: {BG}; }}
[data-testid="stHeader"] {{ background: transparent !important; }}
[data-testid="stMain"] .block-container {{ padding-top: 1.2rem; max-width: 1400px; }}

.hero {{
    background: linear-gradient(135deg, {CARD} 0%, #0D1117 100%);
    border: 1px solid {BORD};
    border-radius: 16px;
    padding: 28px 36px;
    margin-bottom: 6px;
    position: relative;
    overflow: hidden;
}}
.hero::before {{
    content: '✈';
    position: absolute; right: 36px; top: 50%;
    transform: translateY(-50%) rotate(15deg);
    font-size: 5rem; opacity: 0.05;
    pointer-events: none;
}}
.hero-name  {{ font-size: 1.7rem; font-weight: 700; color: {TEXT}; margin: 0; line-height: 1.2; }}
.hero-sub   {{ font-size: 0.85rem; color: {MUTED}; margin-top: 5px; }}
.hero-badge {{
    display: inline-block;
    background: rgba(88,166,255,.12);
    border: 1px solid rgba(88,166,255,.3);
    border-radius: 20px;
    padding: 3px 14px;
    font-size: 0.73rem; font-weight: 600; color: {BLUE};
    margin-top: 10px; letter-spacing: .04em;
}}

.kpi {{
    background: {CARD};
    border: 1px solid {BORD};
    border-radius: 14px;
    padding: 20px 18px;
    height: 100%;
}}
.kpi-icon {{ font-size: 1.3rem; }}
.kpi-val  {{ font-size: 2.3rem; font-weight: 700; color: {BLUE}; line-height: 1.05; margin-top: 4px; }}
.kpi-lbl  {{ font-size: 0.68rem; color: {MUTED}; text-transform: uppercase;
             letter-spacing: .1em; margin-top: 3px; }}
.kpi-sub  {{ font-size: 0.78rem; color: {MUTED}; margin-top: 10px;
             border-top: 1px solid {BORD}; padding-top: 8px; }}
.kpi-green .kpi-val  {{ color: {GREEN}; }}
.kpi-purple .kpi-val {{ color: {PURPLE}; }}
.kpi-orange .kpi-val {{ color: {ORANGE}; }}
.kpi-red .kpi-val    {{ color: {RED}; }}

.sec {{
    font-size: 0.68rem; font-weight: 600; color: {MUTED};
    text-transform: uppercase; letter-spacing: .14em;
    margin: 26px 0 14px;
    display: flex; align-items: center; gap: 10px;
}}
.sec::after {{ content: ''; flex: 1; height: 1px; background: {BORD}; }}

.next-card {{
    background: linear-gradient(135deg, #0D2818 0%, #0F3020 100%);
    border: 1px solid rgba(63,185,80,.3);
    border-radius: 14px; padding: 20px 22px;
    height: 100%;
}}
.next-label {{ font-size: 0.65rem; color: {GREEN}; text-transform: uppercase;
               letter-spacing: .1em; font-weight: 700; }}
.next-name  {{ font-size: 1.15rem; font-weight: 700; color: {TEXT}; margin-top: 6px; }}
.next-dest  {{ font-size: 0.82rem; color: {MUTED}; margin-top: 3px; }}
.next-dates {{ font-size: 0.78rem; color: {MUTED}; margin-top: 6px; }}
.next-tag {{
    display: inline-block;
    background: rgba(63,185,80,.15);
    border: 1px solid rgba(63,185,80,.3);
    border-radius: 20px; padding: 3px 12px;
    font-size: 0.73rem; color: {GREEN}; font-weight: 600; margin-top: 10px;
}}
.active-card {{
    background: linear-gradient(135deg, #1A0D2E 0%, #1E0F35 100%);
    border: 1px solid rgba(163,113,247,.35);
    border-radius: 14px; padding: 20px 22px;
    height: 100%;
}}
.active-label {{ font-size: 0.65rem; color: {PURPLE}; text-transform: uppercase;
                 letter-spacing: .1em; font-weight: 700; }}
.active-name  {{ font-size: 1.15rem; font-weight: 700; color: {TEXT}; margin-top: 6px; }}
.active-dest  {{ font-size: 0.82rem; color: {MUTED}; margin-top: 3px; }}
.active-tag {{
    display: inline-block;
    background: rgba(163,113,247,.15);
    border: 1px solid rgba(163,113,247,.3);
    border-radius: 20px; padding: 3px 12px;
    font-size: 0.73rem; color: {PURPLE}; font-weight: 600; margin-top: 10px;
}}
.no-trips {{
    background: {CARD}; border: 1px solid {BORD};
    border-radius: 14px; padding: 20px 22px;
    text-align: center; color: {MUTED}; font-size: 0.85rem;
    height: 100%;
}}

[data-testid="stTabs"] [role="tablist"] {{
    background: {CARD}; border-radius: 10px;
    padding: 4px; border: 1px solid {BORD}; gap: 2px;
}}
[data-testid="stTabs"] button {{
    border-radius: 8px !important;
    color: {MUTED} !important; font-size: 0.83rem !important;
}}
[data-testid="stTabs"] button[aria-selected="true"] {{
    background: rgba(88,166,255,.12) !important; color: {BLUE} !important;
}}
[data-testid="stTabs"] [role="tabpanel"] {{ padding-top: 20px; }}

.login-wrap {{
    max-width: 360px; margin: 80px auto 0;
    background: {CARD}; border: 1px solid {BORD};
    border-radius: 18px; padding: 40px 36px;
}}
.login-icon  {{ text-align: center; font-size: 2.6rem; margin-bottom: 6px; }}
.login-title {{ text-align: center; font-size: 1.25rem; font-weight: 700; color: {TEXT}; }}
.login-sub   {{ text-align: center; font-size: 0.8rem; color: {MUTED}; margin-top: 4px; margin-bottom: 28px; }}

.chip {{
    display: inline-flex; align-items: center; gap: 6px;
    font-size: 0.75rem; color: {MUTED};
    background: {CARD}; border: 1px solid {BORD};
    border-radius: 20px; padding: 3px 10px;
}}
.chip-dot {{ width: 9px; height: 9px; border-radius: 3px; flex-shrink: 0; }}

hr {{ border-color: {BORD} !important; }}
[data-testid="stCaptionContainer"] {{ color: {MUTED} !important; }}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# 1. AUTENTICACIÓN
# ─────────────────────────────────────────────────────────────────────────────
def check_password():
    if st.session_state.get("authenticated"):
        return True
    st.markdown("""
    <div class='login-wrap'>
        <div class='login-icon'>✈️</div>
        <div class='login-title'>Sebastián Yarad Luco</div>
        <div class='login-sub'>Dashboard de Viajes — acceso privado</div>
    </div>
    """, unsafe_allow_html=True)
    col_l, col_c, col_r = st.columns([1, 1.4, 1])
    with col_c:
        pwd = st.text_input("", type="password", placeholder="Contraseña de acceso",
                            label_visibility="collapsed")
        if st.button("Ingresar →", use_container_width=True, type="primary"):
            if pwd == st.secrets["ACCESS_PASSWORD"]:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")
    return False

if not check_password():
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# 2. CARGA DESDE GOOGLE DRIVE
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(ttl=300, show_spinner="Cargando datos desde Google Drive…")
def load_excel_from_drive() -> bytes:
    import json
    creds_info = json.loads(st.secrets["GCP_SERVICE_ACCOUNT_JSON"])
    creds = service_account.Credentials.from_service_account_info(
        creds_info,
        scopes=["https://www.googleapis.com/auth/drive.readonly"],
    )
    service = build("drive", "v3", credentials=creds, cache_discovery=False)
    file_id = st.secrets["GDRIVE_FILE_ID"]
    try:
        return service.files().export(
            fileId=file_id,
            mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ).execute()
    except Exception:
        from googleapiclient.http import MediaIoBaseDownload
        request = service.files().get_media(fileId=file_id)
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        buf.seek(0)
        return buf.read()

# ─────────────────────────────────────────────────────────────────────────────
# 3. PARSEO
# ─────────────────────────────────────────────────────────────────────────────
def get_region(nombre):
    nu = nombre.upper()
    for label, col, kws in REGION_MAP:
        if any(k in nu for k in kws):
            return label, "#" + col
    return "Otros", "#8B949E"

def to_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    return None

@st.cache_data(ttl=300, show_spinner=False)
def parse_trips(raw_bytes: bytes) -> pd.DataFrame:
    wb = pd.read_excel(io.BytesIO(raw_bytes), sheet_name="DATA", header=0)
    col_n      = wb.columns[0]
    col_nombre = wb.columns[2]
    col_sal    = wb.columns[7]
    col_reg    = wb.columns[10]
    col_est    = wb.columns[12]
    col_ciu    = wb.columns[15]
    col_pai    = wb.columns[16]
    trips_raw = {}
    for _, row in wb.iterrows():
        n = row[col_n]
        if pd.isna(n):
            continue
        n = int(n)
        if n not in trips_raw:
            trips_raw[n] = {
                "nombre": str(row[col_nombre] or ""),
                "estado": str(row[col_est] or ""),
                "tramos": [],
            }
        fs = to_date(row[col_sal])
        if fs:
            fl = to_date(row[col_reg]) or fs
            trips_raw[n]["tramos"].append({
                "fs": fs, "fl": fl,
                "ciudad": str(row[col_ciu] or ""),
                "pais":   str(row[col_pai] or ""),
            })
    rows = []
    for n, data in sorted(trips_raw.items()):
        tr = data["tramos"]
        if not tr:
            continue
        inicio = min(t["fs"] for t in tr)
        fin    = max(t["fl"] for t in tr)
        region, color = get_region(data["nombre"])
        dest = tr[0]["ciudad"] if tr else ""
        pais = tr[0]["pais"]   if tr else ""
        rows.append({
            "n":       n,
            "nombre":  data["nombre"],
            "estado":  data["estado"],
            "inicio":  inicio,
            "fin":     fin,
            "dias":    (fin - inicio).days + 1,
            "ciudad":  dest,
            "pais":    pais,
            "destino": f"{dest}, {pais}" if pais else dest,
            "region":  region,
            "color":   color,
            "sem_ini": inicio.isocalendar()[1],
            "sem_fin": fin.isocalendar()[1],
        })
    return pd.DataFrame(rows)

# ─────────────────────────────────────────────────────────────────────────────
# 4. CARGA
# ─────────────────────────────────────────────────────────────────────────────
try:
    raw = load_excel_from_drive()
except Exception as e:
    st.error(f"Error al conectar con Google Drive: {e}")
    st.stop()

df = parse_trips(raw)
if df.empty:
    st.warning("No se encontraron viajes en el archivo.")
    st.stop()

today = datetime.date.today()
year  = df["inicio"].apply(lambda x: x.year).value_counts().idxmax()
df["inicio_dt"] = pd.to_datetime(df["inicio"])
df["fin_dt"]    = pd.to_datetime(df["fin"])
df_year = df[df["inicio"].apply(lambda x: x.year) == year].copy()

# ─────────────────────────────────────────────────────────────────────────────
# 5. CABECERA
# ─────────────────────────────────────────────────────────────────────────────
col_hd, col_btn = st.columns([9, 1])
with col_hd:
    st.markdown(f"""
    <div class='hero'>
        <div class='hero-name'>Sebastián Yarad Luco</div>
        <div class='hero-sub'>Dashboard de Viajes · registro personal y profesional</div>
        <div class='hero-badge'>✈ Temporada {year}</div>
    </div>
    """, unsafe_allow_html=True)
with col_btn:
    st.markdown("<div style='height:30px'></div>", unsafe_allow_html=True)
    if st.button("↻ Recargar", help="Fuerza descarga fresca desde Google Drive"):
        st.cache_data.clear()
        st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
# 6. KPIs
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<div class='sec'>Resumen general</div>", unsafe_allow_html=True)

total_viajes = len(df_year)
total_dias   = int(df_year["dias"].sum())
pct_year     = round(total_dias / 365 * 100)
fin_count    = (df_year["estado"].str.upper() == "FINALIZADO").sum()
n_paises     = df_year["pais"].nunique()
n_regiones   = df_year["region"].nunique()

upcoming = df_year[df_year["inicio"].apply(lambda x: x) > today].sort_values("inicio")
en_curso = df_year[df_year.apply(lambda r: r["inicio"] <= today <= r["fin"], axis=1)]

dias_hasta_prox = "—"
prox_info = "Sin viajes futuros registrados"
if not upcoming.empty:
    prox = upcoming.iloc[0]
    diff = (prox["inicio"] - today).days
    dias_hasta_prox = f"en {diff}d"
    prox_info = prox["nombre"][:28]

k1, k2, k3, k4, k5 = st.columns(5)
cards_data = [
    (k1, "✈️", total_viajes,          "Viajes",
     f"{fin_count} finalizados · {total_viajes - fin_count} pendientes", ""),
    (k2, "📅", f"{total_dias}d",       "Días en ruta",
     f"{pct_year}% del año fuera de casa", "kpi-green"),
    (k3, "🌍", n_paises,               "Países",
     "  ·  ".join(df_year["pais"].dropna().unique()[:4]), "kpi-purple"),
    (k4, "📍", n_regiones,             "Regiones",
     "  ·  ".join(df_year["region"].unique()[:4]), "kpi-orange"),
    (k5, "⏰", dias_hasta_prox,        "Próximo viaje",
     prox_info, "kpi-red"),
]
for col, icon, val, lbl, sub, cls in cards_data:
    col.markdown(
        f"<div class='kpi {cls}'>"
        f"<div class='kpi-icon'>{icon}</div>"
        f"<div class='kpi-val'>{val}</div>"
        f"<div class='kpi-lbl'>{lbl}</div>"
        f"<div class='kpi-sub'>{sub}</div>"
        f"</div>",
        unsafe_allow_html=True,
    )

# ─────────────────────────────────────────────────────────────────────────────
# 7. ESTADO ACTUAL + DONUT
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<div class='sec'>Estado actual</div>", unsafe_allow_html=True)
c_act, c_prox, c_chart = st.columns([1.4, 1.4, 3.2])

with c_act:
    if not en_curso.empty:
        t = en_curso.iloc[0]
        dias_rest = (t["fin"] - today).days
        st.markdown(f"""
        <div class='active-card'>
            <div class='active-label'>🟣 En curso ahora</div>
            <div class='active-name'>{t['nombre']}</div>
            <div class='active-dest'>📍 {t['destino']}</div>
            <div style='font-size:.75rem;color:{MUTED};margin-top:5px'>
                {t['inicio'].strftime('%d %b')} → {t['fin'].strftime('%d %b')}
            </div>
            <div class='active-tag'>Quedan {dias_rest} días</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div class='no-trips'>
            <div style='font-size:1.6rem;margin-bottom:8px'>🏡</div>
            <div style='color:{TEXT};font-weight:600;font-size:.9rem'>Sin viaje activo</div>
            <div style='margin-top:4px'>Estás en casa hoy</div>
        </div>
        """, unsafe_allow_html=True)

with c_prox:
    if not upcoming.empty:
        prox = upcoming.iloc[0]
        diff = (prox["inicio"] - today).days
        st.markdown(f"""
        <div class='next-card'>
            <div class='next-label'>🟢 Próximo viaje</div>
            <div class='next-name'>{prox['nombre']}</div>
            <div class='next-dest'>📍 {prox['destino']}</div>
            <div class='next-dates'>
                {prox['inicio'].strftime('%d %b')} → {prox['fin'].strftime('%d %b')} · {prox['dias']} días
            </div>
            <div class='next-tag'>en {diff} día{'s' if diff != 1 else ''}</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div class='no-trips'>
            <div style='font-size:1.6rem;margin-bottom:8px'>🔜</div>
            <div style='color:{TEXT};font-weight:600;font-size:.9rem'>Sin próximos viajes</div>
            <div style='margin-top:4px'>No hay viajes futuros registrados</div>
        </div>
        """, unsafe_allow_html=True)

with c_chart:
    dias_region = df_year.groupby("region")["dias"].sum().reset_index()
    dias_region["color"] = dias_region["region"].map(REGION_COLORS)
    fig_donut = go.Figure(go.Pie(
        labels=dias_region["region"],
        values=dias_region["dias"],
        hole=0.62,
        marker=dict(
            colors=dias_region["color"].tolist(),
            line=dict(color=BG, width=3),
        ),
        textinfo="label+percent",
        textfont=dict(size=11, color=TEXT),
        hovertemplate="<b>%{label}</b><br>%{value} días<extra></extra>",
    ))
    fig_donut.add_annotation(
        text=f"<b>{total_dias}</b><br><span style='font-size:11px'>días totales</span>",
        x=0.5, y=0.5, showarrow=False,
        font=dict(size=15, color=TEXT),
    )
    fig_donut.update_layout(
        height=200,
        margin=dict(l=10, r=10, t=10, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        showlegend=False,
        font=dict(color=TEXT),
    )
    st.plotly_chart(fig_donut, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# 8. TABS
# ─────────────────────────────────────────────────────────────────────────────
tab_cal, tab_tl, tab_ag = st.tabs(["📅  Calendario anual", "🗓  Timeline", "📋  Agenda"])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — CALENDARIO ANUAL
# ══════════════════════════════════════════════════════════════════════════════
with tab_cal:
    travel_days = {}
    for _, trip in df_year.iterrows():
        d = trip["inicio"]
        while d <= trip["fin"]:
            travel_days[d] = trip
            d += datetime.timedelta(days=1)

    jan1  = datetime.date(year, 1, 1)
    dec31 = datetime.date(year, 12, 31)
    all_days = [jan1 + datetime.timedelta(days=i) for i in range((dec31 - jan1).days + 1)]

    week_nums, day_nums, colors_cal, texts_cal = [], [], [], []
    for d in all_days:
        iso = d.isocalendar()
        week_nums.append(iso[1])
        day_nums.append(d.weekday())
        if d in travel_days:
            t = travel_days[d]
            colors_cal.append(t["color"])
            texts_cal.append(f"{d.strftime('%d %b')} · {t['nombre']}<br>📍 {t['destino']}")
        elif d == today:
            colors_cal.append(RED)
            texts_cal.append(f"{d.strftime('%d %b')} · Hoy")
        elif d.weekday() >= 5:
            colors_cal.append("#1C2128")
            texts_cal.append(d.strftime("%d %b"))
        else:
            colors_cal.append("#21262D")
            texts_cal.append(d.strftime("%d %b"))

    DAYS_ES   = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]
    MONTHS_ES = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]

    fig_cal = go.Figure()
    fig_cal.add_trace(go.Scatter(
        x=week_nums, y=day_nums,
        mode="markers",
        marker=dict(
            symbol="square", size=17,
            color=colors_cal,
            line=dict(width=1.5, color=BG),
        ),
        text=texts_cal,
        hoverinfo="text",
    ))

    month_week = {}
    for d in all_days:
        if d.day <= 7 and d.month not in month_week:
            month_week[d.month] = d.isocalendar()[1]

    for m, w in month_week.items():
        fig_cal.add_annotation(
            x=w, y=-1.2, text=MONTHS_ES[m - 1],
            showarrow=False, font=dict(size=10, color=MUTED), yref="y",
        )

    fig_cal.update_layout(
        height=210,
        margin=dict(l=40, r=10, t=10, b=30),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(showgrid=False, zeroline=False, showticklabels=False, range=[0, 54]),
        yaxis=dict(
            showgrid=False, zeroline=False,
            ticktext=DAYS_ES, tickvals=list(range(7)),
            autorange="reversed", tickfont=dict(size=9, color=MUTED),
        ),
        showlegend=False,
    )
    st.plotly_chart(fig_cal, use_container_width=True)

    # Leyenda
    chips_html = ""
    for label, col, _ in REGION_MAP:
        chips_html += (
            f"<span class='chip'>"
            f"<span class='chip-dot' style='background:#{col}'></span>"
            f"{label}</span>&nbsp;"
        )
    chips_html += f"<span class='chip'><span class='chip-dot' style='background:{RED}'></span>Hoy</span>"
    st.markdown(f"<div style='margin-top:6px'>{chips_html}</div>", unsafe_allow_html=True)

    # Actividad mensual
    st.markdown("<div class='sec'>Días de viaje por mes</div>", unsafe_allow_html=True)
    month_dias = {m: 0 for m in range(1, 13)}
    for d in travel_days:
        month_dias[d.month] = month_dias.get(d.month, 0) + 1

    fig_bar = go.Figure(go.Bar(
        x=MONTHS_ES,
        y=[month_dias[m] for m in range(1, 13)],
        marker=dict(
            color=[month_dias[m] for m in range(1, 13)],
            colorscale=[[0, "#21262D"], [1, BLUE]],
            line=dict(width=0),
        ),
        hovertemplate="<b>%{x}</b><br>%{y} días de viaje<extra></extra>",
    ))
    fig_bar.update_layout(
        height=200,
        margin=dict(l=30, r=10, t=10, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(showgrid=False, tickfont=dict(color=MUTED, size=10)),
        yaxis=dict(
            showgrid=True, gridcolor="#21262D",
            tickfont=dict(color=MUTED, size=9),
            title=dict(text="días", font=dict(color=MUTED, size=9)),
        ),
        bargap=0.35,
    )
    st.plotly_chart(fig_bar, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — TIMELINE
# ══════════════════════════════════════════════════════════════════════════════
with tab_tl:
    gantt_rows = []
    for _, trip in df_year.iterrows():
        gantt_rows.append(dict(
            Task     = trip["region"],
            Start    = trip["inicio_dt"].strftime("%Y-%m-%d"),
            Finish   = trip["fin_dt"].strftime("%Y-%m-%d"),
            Resource = trip["region"],
            Viaje    = trip["nombre"],
            Destino  = trip["destino"],
            Dias     = trip["dias"],
        ))

    df_gantt = pd.DataFrame(gantt_rows)
    region_order = [r[0] for r in REGION_MAP if r[0] in df_gantt["Task"].values]

    fig_gantt = px.timeline(
        df_gantt,
        x_start="Start", x_end="Finish",
        y="Task",
        color="Resource",
        color_discrete_map=REGION_COLORS,
        hover_name="Viaje",
        hover_data={"Destino": True, "Dias": True,
                    "Start": True, "Finish": True,
                    "Task": False, "Resource": False},
        labels={"Task": "", "Resource": "Región"},
        category_orders={"Task": region_order},
    )
    fig_gantt.add_vline(
        x=today.strftime("%Y-%m-%d"),
        line_dash="dot", line_color=RED, line_width=1.5,
        annotation_text="Hoy",
        annotation_font=dict(color=RED, size=10),
    )
    fig_gantt.update_layout(
        height=280,
        margin=dict(l=120, r=20, t=16, b=30),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        showlegend=False,
        xaxis=dict(
            showgrid=True, gridcolor="#21262D",
            tickformat="%b %Y", tickfont=dict(size=10, color=MUTED),
            range=[f"{year}-01-01", f"{year}-12-31"],
            zeroline=False,
        ),
        yaxis=dict(tickfont=dict(size=10, color=MUTED), showgrid=False),
        bargap=0.35,
        font=dict(color=TEXT),
    )
    fig_gantt.update_traces(marker_line_width=0, opacity=0.92)
    st.plotly_chart(fig_gantt, use_container_width=True)

    chips_html2 = ""
    for label, col, _ in REGION_MAP:
        if label in df_gantt["Task"].values:
            chips_html2 += (
                f"<span class='chip'>"
                f"<span class='chip-dot' style='background:#{col}'></span>"
                f"{label}</span>&nbsp;"
            )
    st.markdown(f"<div>{chips_html2}</div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — AGENDA
# ══════════════════════════════════════════════════════════════════════════════
with tab_ag:
    fc1, fc2, fc3 = st.columns([2, 2, 3])
    with fc1:
        regiones_sel = st.multiselect(
            "Región", df_year["region"].unique().tolist(),
            default=df_year["region"].unique().tolist(),
        )
    with fc2:
        estados_sel = st.multiselect(
            "Estado", df_year["estado"].dropna().unique().tolist(),
            default=df_year["estado"].dropna().unique().tolist(),
        )
    with fc3:
        search = st.text_input("🔍 Buscar", placeholder="nombre, destino, país…")

    df_filt = df_year[
        df_year["region"].isin(regiones_sel) &
        df_year["estado"].isin(estados_sel)
    ]
    if search:
        mask = (
            df_filt["nombre"].str.contains(search, case=False, na=False) |
            df_filt["destino"].str.contains(search, case=False, na=False) |
            df_filt["pais"].str.contains(search, case=False, na=False)
        )
        df_filt = df_filt[mask]

    df_show = df_filt[[
        "n", "estado", "nombre", "destino",
        "inicio", "fin", "dias", "sem_ini", "sem_fin", "region",
    ]].copy()
    df_show["inicio"] = df_show["inicio"].apply(lambda x: x.strftime("%d %b"))
    df_show["fin"]    = df_show["fin"].apply(lambda x: x.strftime("%d %b"))
    df_show.columns = ["N°", "Estado", "Viaje", "Destino",
                        "Salida", "Regreso", "Días", "Sem sal.", "Sem reg.", "Región"]

    st.dataframe(
        df_show,
        use_container_width=True,
        hide_index=True,
        height=min(50 + len(df_show) * 36, 520),
        column_config={
            "N°":       st.column_config.NumberColumn(width="small"),
            "Estado":   st.column_config.TextColumn(width="medium"),
            "Viaje":    st.column_config.TextColumn(width="large"),
            "Destino":  st.column_config.TextColumn(width="large"),
            "Salida":   st.column_config.TextColumn(width="small"),
            "Regreso":  st.column_config.TextColumn(width="small"),
            "Días":     st.column_config.NumberColumn(width="small"),
            "Sem sal.": st.column_config.NumberColumn(width="small"),
            "Sem reg.": st.column_config.NumberColumn(width="small"),
            "Región":   st.column_config.TextColumn(width="medium"),
        },
    )
    st.caption(
        f"**{len(df_filt)} viajes** · **{int(df_filt['dias'].sum())} días** "
        f"| Mostrando {len(df_show)} de {total_viajes} en {year}"
    )

# ─────────────────────────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    f"<div style='text-align:center;color:{MUTED};font-size:.72rem'>"
    f"Sebastián Yarad Luco · Dashboard privado · Datos desde Google Drive"
    f"</div>",
    unsafe_allow_html=True,
)
      background: white; border-radius: 12px; padding: 20px 24px;
      box-shadow: 0 1px 6px rgba(0,0,0,0.07); text-align: center;
  }
  .kpi-val  { font-size: 2.2rem; font-weight: 700; color: #1B4F72; line-height: 1.1; }
  .kpi-lbl  { font-size: 0.78rem; color: #888; margin-top: 4px; text-transform: uppercase; letter-spacing: .05em; }
  .kpi-sub  { font-size: 0.82rem; color: #555; margin-top: 6px; }
  .section-title {
      font-size: 1rem; font-weight: 700; color: #1A1A2E;
      border-left: 4px solid #1B4F72; padding-left: 10px; margin: 24px 0 12px;
  }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# 1. AUTENTICACIÓN
# ─────────────────────────────────────────────────────────────────────────────
def check_password():
    if st.session_state.get("authenticated"):
        return True
    st.markdown("""
    <div style='max-width:380px;margin:80px auto 0;background:white;
         border-radius:16px;padding:40px;box-shadow:0 4px 24px rgba(0,0,0,0.1)'>
    <div style='text-align:center;margin-bottom:28px'>
      <div style='font-size:2.5rem'>✈</div>
      <div style='font-size:1.3rem;font-weight:700;color:#1A1A2E'>Pura Fruit</div>
      <div style='font-size:0.85rem;color:#888'>Dashboard de Viajes</div>
    </div>
    """, unsafe_allow_html=True)
    pwd = st.text_input("Contraseña", type="password", placeholder="Ingresá la clave de acceso")
    btn = st.button("Ingresar", use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)
    if btn:
        if pwd == st.secrets["ACCESS_PASSWORD"]:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Contraseña incorrecta.")
    return False

if not check_password():
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# 2. CARGA DESDE GOOGLE DRIVE
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(ttl=300, show_spinner="Cargando datos desde Google Sheets…")
def load_excel_from_drive() -> bytes:
    import json
    # Leer el JSON completo como string — evita problemas de TOML con private_key
    creds_info = json.loads(st.secrets["GCP_SERVICE_ACCOUNT_JSON"])
    creds = service_account.Credentials.from_service_account_info(
        creds_info,
        scopes=["https://www.googleapis.com/auth/drive.readonly"],
    )
    service = build("drive", "v3", credentials=creds, cache_discovery=False)
    file_id = st.secrets["GDRIVE_FILE_ID"]
    # Intentar export (Google Sheet nativo) y si falla, usar get_media (xlsx subido)
    try:
        return service.files().export(
            fileId=file_id,
            mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ).execute()
    except Exception:
        import io
        from googleapiclient.http import MediaIoBaseDownload
        request = service.files().get_media(fileId=file_id)
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        buf.seek(0)
        return buf.read()

# ─────────────────────────────────────────────────────────────────────────────
# 3. PARSEO DE DATOS
# ─────────────────────────────────────────────────────────────────────────────
REGION_MAP = [
    ("Piura",        "1B4F72", ["PIU"]),
    ("USA / Canadá", "1F618D", ["USA","PHL","CAN","CPMA","HONEY","BKSFLD","IFPA"]),
    ("Europa",       "154360", ["BERLIN","EUR","ITGS"]),
    ("Asia / HK",    "1B6CA8", ["AFL","HK"]),
    ("Chile",        "5D6D7E", ["BBA","SIX","BALMACEDA"]),
]
REGION_COLORS = {label: "#" + col for label, col, _ in REGION_MAP}

def get_region(nombre):
    nu = nombre.upper()
    for label, col, kws in REGION_MAP:
        if any(k in nu for k in kws):
            return label, "#" + col
    return "Otros", "#808080"

def to_date(v):
    if isinstance(v, datetime.datetime): return v.date()
    if isinstance(v, datetime.date):     return v
    return None

@st.cache_data(ttl=300, show_spinner=False)
def parse_trips(raw_bytes: bytes) -> pd.DataFrame:
    wb = pd.read_excel(io.BytesIO(raw_bytes), sheet_name="DATA", header=0)
    col_n      = wb.columns[0]
    col_nombre = wb.columns[2]
    col_sal    = wb.columns[7]
    col_reg    = wb.columns[10]
    col_est    = wb.columns[12]
    col_ciu    = wb.columns[15]
    col_pai    = wb.columns[16]

    trips_raw = {}
    for _, row in wb.iterrows():
        n = row[col_n]
        if pd.isna(n): continue
        n = int(n)
        if n not in trips_raw:
            trips_raw[n] = {"nombre": str(row[col_nombre] or ""), "estado": str(row[col_est] or ""), "tramos": []}
        fs = to_date(row[col_sal])
        if fs:
            fl = to_date(row[col_reg]) or fs
            trips_raw[n]["tramos"].append({"fs": fs, "fl": fl, "ciudad": str(row[col_ciu] or ""), "pais": str(row[col_pai] or "")})

    rows = []
    for n, data in sorted(trips_raw.items()):
        tr = data["tramos"]
        if not tr: continue
        inicio = min(t["fs"] for t in tr)
        fin    = max(t["fl"] for t in tr)
        region, color = get_region(data["nombre"])
        rows.append({
            "n": n, "nombre": data["nombre"], "estado": data["estado"],
            "inicio": inicio, "fin": fin, "dias": (fin - inicio).days + 1,
            "ciudad": tr[0]["ciudad"], "pais": tr[0]["pais"],
            "destino": f"{tr[0]['ciudad']}, {tr[0]['pais']}" if tr[0]["pais"] else tr[0]["ciudad"],
            "region": region, "color": color,
            "sem_ini": inicio.isocalendar()[1], "sem_fin": fin.isocalendar()[1],
        })
    return pd.DataFrame(rows)

# ─────────────────────────────────────────────────────────────────────────────
# 4. CARGA + CABECERA
# ─────────────────────────────────────────────────────────────────────────────
try:
    raw = load_excel_from_drive()
except Exception as e:
    st.error(f"Error al conectar con Google Drive: {e}")
    st.stop()

df = parse_trips(raw)
if df.empty:
    st.warning("No se encontraron viajes en el archivo.")
    st.stop()

today = datetime.date.today()
year  = df["inicio"].apply(lambda x: x.year).value_counts().idxmax()
df["inicio"] = pd.to_datetime(df["inicio"])
df["fin"]    = pd.to_datetime(df["fin"])

col_h1, col_h2 = st.columns([8, 1])
with col_h1:
    st.markdown(
        f"<h2 style='color:#1A1A2E;margin:0;padding:16px 0 4px'>✈ Dashboard de Viajes · {year}</h2>"
        f"<div style='color:#888;font-size:.85rem;padding-bottom:12px'>"
        f"Pura Fruit Company &nbsp;·&nbsp; Actualizado: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}</div>",
        unsafe_allow_html=True,
    )
with col_h2:
    if st.button("↻ Recargar", help="Fuerza descarga fresca desde Google Sheets"):
        st.cache_data.clear()
        st.rerun()

st.divider()

# ─────────────────────────────────────────────────────────────────────────────
# 5. KPI CARDS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<div class='section-title'>Resumen</div>", unsafe_allow_html=True)

total_trips = len(df)
total_days  = df["dias"].sum()
fin_count   = (df["estado"].str.upper() == "FINALIZADO").sum()
upcoming    = df[df["inicio"].dt.date >= today].sort_values("inicio")
next_str    = ""
if not upcoming.empty:
    nx = upcoming.iloc[0]
    diff = (nx["inicio"].date() - today).days
    next_str = f"{nx['nombre'][:22]} · en {diff}d"

k1, k2, k3, k4, k5 = st.columns(5)
for col, val, lbl, sub in [
    (k1, total_trips,            "Viajes totales",    f"{fin_count} finalizados"),
    (k2, f"{total_days}d",       "Días fuera",        f"~{total_days/365*100:.0f}% del año"),
    (k3, df["region"].nunique(), "Regiones",          ", ".join(df["region"].unique()[:3])),
    (k4, df["pais"].nunique(),   "Países",            ", ".join(df["pais"].dropna().unique()[:3])),
    (k5, upcoming.shape[0],      "Próximos viajes",   next_str or "—"),
]:
    col.markdown(
        f"<div class='kpi-card'><div class='kpi-val'>{val}</div>"
        f"<div class='kpi-lbl'>{lbl}</div><div class='kpi-sub'>{sub}</div></div>",
        unsafe_allow_html=True,
    )

# ─────────────────────────────────────────────────────────────────────────────
# 6. CALENDARIO ANUAL
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<div class='section-title'>Calendario anual</div>", unsafe_allow_html=True)

travel_days = {}
for _, trip in df.iterrows():
    d = trip["inicio"].date()
    while d <= trip["fin"].date():
        travel_days[d] = trip
        d += datetime.timedelta(days=1)

jan1  = datetime.date(year, 1, 1)
dec31 = datetime.date(year, 12, 31)
all_days = [jan1 + datetime.timedelta(days=i) for i in range((dec31 - jan1).days + 1)]

week_nums, day_nums, colors, texts = [], [], [], []
for d in all_days:
    iso = d.isocalendar()
    week_nums.append(iso[1])
    day_nums.append(d.weekday())
    if d in travel_days:
        t = travel_days[d]
        colors.append(t["color"])
        texts.append(f"{d.strftime('%d %b')} · {t['nombre']}<br>{t['destino']}")
    elif d == today:
        colors.append("#E74C3C")
        texts.append(f"{d.strftime('%d %b')} · Hoy")
    elif d.weekday() >= 5:
        colors.append("#EAECEE")
        texts.append(d.strftime('%d %b'))
    else:
        colors.append("#F4F6F7")
        texts.append(d.strftime('%d %b'))

DAYS_LABEL = ["Lun","Mar","Mié","Jue","Vie","Sáb","Dom"]
fig_cal = go.Figure()
fig_cal.add_trace(go.Scatter(
    x=week_nums, y=day_nums, mode="markers",
    marker=dict(symbol="square", size=18, color=colors, line=dict(width=1, color="white")),
    text=texts, hoverinfo="text",
))
MONTHS_ES = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
month_week = {}
for d in all_days:
    if d.day == 1:
        month_week[d.month] = d.isocalendar()[1]
for m, w in month_week.items():
    fig_cal.add_annotation(x=w, y=-1, text=MONTHS_ES[m-1], showarrow=False,
                           font=dict(size=10, color="#888"), yref="y")
fig_cal.update_layout(
    height=200, margin=dict(l=40, r=20, t=10, b=30),
    paper_bgcolor="white", plot_bgcolor="white",
    xaxis=dict(showgrid=False, zeroline=False, showticklabels=False, range=[0, 54]),
    yaxis=dict(showgrid=False, zeroline=False, ticktext=DAYS_LABEL, tickvals=list(range(7)),
               autorange="reversed", tickfont=dict(size=9, color="#aaa")),
    showlegend=False,
)
st.plotly_chart(fig_cal, use_container_width=True)

leg_cols = st.columns(len(REGION_MAP) + 1)
for i, (label, col, _) in enumerate(REGION_MAP):
    leg_cols[i].markdown(
        f"<div style='display:flex;align-items:center;gap:6px;font-size:.8rem;color:#555'>"
        f"<span style='width:12px;height:12px;border-radius:3px;background:#{col};display:inline-block'></span>"
        f"{label}</div>", unsafe_allow_html=True)
leg_cols[-1].markdown(
    "<div style='display:flex;align-items:center;gap:6px;font-size:.8rem;color:#555'>"
    "<span style='width:12px;height:12px;border-radius:3px;background:#E74C3C;display:inline-block'></span>"
    "Hoy</div>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# 7. SWIM-LANES
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<div class='section-title'>Swim-lanes por región</div>", unsafe_allow_html=True)

gantt_rows = []
for _, trip in df.iterrows():
    gantt_rows.append(dict(
        Task=trip["region"], Start=trip["inicio"].strftime("%Y-%m-%d"),
        Finish=trip["fin"].strftime("%Y-%m-%d"), Resource=trip["region"],
        Label=f"{trip['nombre']} · {trip['destino']}", Dias=trip["dias"],
    ))
df_gantt = pd.DataFrame(gantt_rows)
region_order = [r[0] for r in REGION_MAP if r[0] in df_gantt["Task"].values]

fig_gantt = px.timeline(
    df_gantt, x_start="Start", x_end="Finish", y="Task", color="Resource",
    color_discrete_map=REGION_COLORS, hover_name="Label",
    hover_data={"Dias": True, "Start": True, "Finish": True, "Task": False, "Resource": False},
    category_orders={"Task": region_order},
)
fig_gantt.add_vline(x=today.strftime("%Y-%m-%d"), line_dash="dot", line_color="#E74C3C",
                    line_width=2, annotation_text="Hoy", annotation_font_color="#E74C3C")
fig_gantt.update_layout(
    height=260, margin=dict(l=120, r=20, t=10, b=30),
    paper_bgcolor="white", plot_bgcolor="white", showlegend=False,
    xaxis=dict(showgrid=True, gridcolor="#F0F0F0", tickformat="%b %Y",
               tickfont=dict(size=10), range=[f"{year}-01-01", f"{year}-12-31"]),
    yaxis=dict(tickfont=dict(size=10)), bargap=0.35,
)
st.plotly_chart(fig_gantt, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# 8. TABLA AGENDA
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<div class='section-title'>Agenda detallada</div>", unsafe_allow_html=True)

fc1, fc2, fc3 = st.columns([2, 2, 3])
with fc1:
    regiones_sel = st.multiselect("Región", df["region"].unique().tolist(), default=df["region"].unique().tolist())
with fc2:
    estados_sel = st.multiselect("Estado", df["estado"].dropna().unique().tolist(), default=df["estado"].dropna().unique().tolist())
with fc3:
    search = st.text_input("Buscar viaje", placeholder="nombre, destino…")

df_filt = df[df["region"].isin(regiones_sel) & df["estado"].isin(estados_sel)]
if search:
    mask = (df_filt["nombre"].str.contains(search, case=False, na=False) |
            df_filt["destino"].str.contains(search, case=False, na=False))
    df_filt = df_filt[mask]

df_show = df_filt[["n","estado","nombre","destino","inicio","fin","dias","sem_ini","sem_fin","region"]].copy()
df_show.columns = ["N°","Estado","Viaje","Destino","Salida","Regreso","Días","Sem sal.","Sem reg.","Región"]
df_show["Salida"]  = df_show["Salida"].dt.strftime("%d-%b")
df_show["Regreso"] = df_show["Regreso"].dt.strftime("%d-%b")

st.dataframe(
    df_show, use_container_width=True, hide_index=True,
    height=min(40 + len(df_show) * 36, 500),
    column_config={
        "N°":      st.column_config.NumberColumn(width="small"),
        "Estado":  st.column_config.TextColumn(width="medium"),
        "Viaje":   st.column_config.TextColumn(width="large"),
        "Destino": st.column_config.TextColumn(width="large"),
        "Salida":  st.column_config.TextColumn(width="small"),
        "Regreso": st.column_config.TextColumn(width="small"),
        "Días":    st.column_config.NumberColumn(width="small"),
        "Sem sal.":st.column_config.NumberColumn(width="small"),
        "Sem reg.":st.column_config.NumberColumn(width="small"),
        "Región":  st.column_config.TextColumn(width="medium"),
    }
)
st.caption(f"**{len(df_filt)} viajes** · **{df_filt['dias'].sum()} días** | Mostrando {len(df_show)} de {total_trips} viajes")

st.markdown("---")
st.markdown(
    "<div style='text-align:center;color:#ccc;font-size:.75rem'>"
    "Pura Fruit Company · Dashboard generado por Claude · Datos desde Google Sheets</div>",
    unsafe_allow_html=True,
)
