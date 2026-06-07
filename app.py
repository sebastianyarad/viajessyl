"""
Pura Fruit — Dashboard de Viajes
Fuente de datos: Google Sheets (export as xlsx via Drive API)
Acceso: contraseña simple via st.secrets
"""

import streamlit as st
import pandas as pd
import datetime
import calendar
import io
from collections import Counter

# ── Google Drive ──────────────────────────────────────────────────────────────
from googleapiclient.discovery import build
from google.oauth2 import service_account

# ── Plotly ────────────────────────────────────────────────────────────────────
import plotly.graph_objects as go
import plotly.express as px

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Pura Fruit · Viajes",
    page_icon="✈",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
  [data-testid="stAppViewContainer"] { background: #F7F9FC; }
  [data-testid="stHeader"] { background: transparent; }
  .kpi-card {
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
