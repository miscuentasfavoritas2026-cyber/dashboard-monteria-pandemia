# Dashboard educativo Monteria - version v8 corregida
# Mapa con Folium visible + exportacion horizontal PNG/JPG/PDF + informe PDF con graficos.
# Ejecutar:
#   python -m pip install -r requirements_monteria_v8.txt
#   python -m streamlit run dashboard_monteria_v8_corregido.py

from __future__ import annotations

from pathlib import Path
from io import BytesIO
from datetime import datetime
import re

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

import folium
from streamlit_folium import st_folium

from PIL import Image, ImageDraw, ImageFont
import matplotlib.pyplot as plt
from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas as pdfcanvas
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak

st.set_page_config(
    page_title="Monitoreo del aprendizaje remoto - Montería",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

DATA_PATH = Path("data_moodle_monteria_dashboard_mapa.xlsx")
APP_TITLE = "Monitoreo del aprendizaje remoto durante la pandemia en instituciones educativas urbanas de Montería"
APP_SUBTITLE = "Participación, rendimiento y riesgo escolar"
AUTHOR = "Mónica Fernanda Ojeda González"
CACHE_VERSION = "v8_titulo_mapa_nombres_2026_05_12"

# =============================================================================
# ESTILOS
# =============================================================================
st.markdown(
    """
    <style>
        :root {
            --azul: #0f3d73;
            --turquesa: #0f766e;
            --verde: #22c55e;
            --ambar: #f59e0b;
            --rojo: #ef4444;
            --gris: #64748b;
            --borde: rgba(148, 163, 184, .32);
        }
        .stApp {
            background: radial-gradient(circle at top left, #dbeafe 0, #f8fafc 34%, #eef2ff 100%);
        }
        .block-container {
            padding-top: .65rem;
            padding-left: 1.1rem;
            padding-right: 1.1rem;
            padding-bottom: .6rem;
            max-width: 100%;
        }
        div[data-testid="stVerticalBlock"] { gap: .58rem; }
        div[data-testid="column"] { padding-left: .18rem; padding-right: .18rem; }
        .hero {
            padding: .9rem 1.15rem;
            border-radius: 24px;
            color: white;
            background: linear-gradient(120deg, #071a33 0%, #0f3d73 45%, #0f766e 100%);
            box-shadow: 0 18px 45px rgba(15, 23, 42, .22);
            border: 1px solid rgba(255,255,255,.18);
        }
        .hero-title { font-size: 1.65rem; font-weight: 950; margin: 0; letter-spacing: -.02em; }
        .hero-subtitle { margin-top: .25rem; font-size: .9rem; opacity: .94; }
        .hero-author { margin-top: .42rem; font-size: .84rem; font-weight: 850; opacity: .98; }
        .hero-author span {
            display: inline-block; padding: .18rem .52rem; border-radius: 999px;
            background: rgba(255,255,255,.18); border: 1px solid rgba(255,255,255,.24);
        }
        .chip {
            display: inline-block; padding: .2rem .55rem; border-radius: 999px;
            background: rgba(255,255,255,.16); border: 1px solid rgba(255,255,255,.22);
            font-size: .72rem; font-weight: 750; margin-right: .3rem;
        }
        .panel {
            background: rgba(255,255,255,.94); border-radius: 22px; padding: .75rem .82rem;
            box-shadow: 0 12px 30px rgba(15, 23, 42, .08); border: 1px solid var(--borde); height: 100%;
        }
        .panel-title { font-size: .95rem; font-weight: 950; color: #0f172a; margin-bottom: .25rem; }
        .panel-help { font-size: .72rem; color: #64748b; margin-top: -.12rem; margin-bottom: .25rem; }
        .metric-card {
            background: rgba(255,255,255,.98); border-radius: 18px; padding: .62rem .72rem;
            border: 1px solid var(--borde); box-shadow: 0 8px 20px rgba(15, 23, 42, .07); min-height: 86px;
        }
        .metric-label {
            color: #64748b; font-size: .68rem; font-weight: 900; text-transform: uppercase;
            letter-spacing: .055em; line-height: 1.05;
        }
        .metric-value { color: #0f172a; font-size: 1.28rem; font-weight: 950; margin-top: .16rem; line-height: 1.12; }
        .metric-note { color: #64748b; font-size: .68rem; margin-top: .1rem; }
        .institution-card {
            border-radius: 20px; padding: .72rem .78rem;
            background: linear-gradient(135deg, rgba(37,99,235,.08), rgba(15,118,110,.08));
            border: 1px solid rgba(37,99,235,.16);
        }
        .institution-name { color: #0f172a; font-size: 1.02rem; font-weight: 950; line-height: 1.13; margin-bottom: .25rem; }
        .institution-meta { color: #475569; font-size: .75rem; line-height: 1.28; }
        .alert-high, .alert-ok {
            border-radius: 16px; padding: .55rem .65rem; font-size: .78rem; font-weight: 750;
        }
        .alert-high { background: #fff1f2; color: #991b1b; border: 1px solid #fecdd3; }
        .alert-ok { background: #ecfdf5; color: #166534; border: 1px solid #bbf7d0; }
        .stDataFrame { border-radius: 18px; overflow: hidden; }
        header[data-testid="stHeader"] { background: rgba(0,0,0,0); }
        #MainMenu, footer { visibility: hidden; }
    </style>
    """,
    unsafe_allow_html=True,
)

# =============================================================================
# FUNCIONES DE DATOS
# =============================================================================
def clean_text_value(valor):
    """Limpia errores de escritura visibles en nombres institucionales."""
    if pd.isna(valor):
        return valor
    texto = re.sub(r"\s+", " ", str(valor).strip())
    texto = re.sub(r"^\s*(?i:es\s+decir)\s*[,;:\-–—]?\s*", "IE ", texto).strip()
    texto = re.sub(r"^(?i:instituci[oó]n\s+educativa)\s+", "IE ", texto).strip()
    texto = re.sub(r"^(?i:i\.\s*e\.)\s+", "IE ", texto).strip()
    texto = re.sub(r"^(?i:ie)\s+(?i:ie)\s+", "IE ", texto).strip()
    reemplazos = {
        "Mogambo": "IE Mogambo", "IE Mogambo": "IE Mogambo",
        "La Inmaculada": "IE La Inmaculada", "IE La Inmaculada": "IE La Inmaculada",
        "Santa Maria Goretti": "IE Santa María Goretti", "Santa María Goretti": "IE Santa María Goretti",
        "IE Santa Maria Goretti": "IE Santa María Goretti", "IE Santa María Goretti": "IE Santa María Goretti",
        "INEM Lorenzo Maria Lleras": "IE INEM Lorenzo María Lleras", "INEM Lorenzo María Lleras": "IE INEM Lorenzo María Lleras",
        "IE INEM Lorenzo Maria Lleras": "IE INEM Lorenzo María Lleras", "IE INEM Lorenzo María Lleras": "IE INEM Lorenzo María Lleras",
        "Liceo Guillermo Valencia": "IE Liceo Guillermo Valencia", "IE Liceo Guillermo Valencia": "IE Liceo Guillermo Valencia",
        "San Jose": "IE San José", "San José": "IE San José", "IE San Jose": "IE San José", "IE San José": "IE San José",
        "Manuel Ruiz Alvarez": "IE Manuel Ruiz Álvarez", "Manuel Ruiz Álvarez": "IE Manuel Ruiz Álvarez",
        "IE Manuel Ruiz Alvarez": "IE Manuel Ruiz Álvarez", "IE Manuel Ruiz Álvarez": "IE Manuel Ruiz Álvarez",
        "Mercedes Abrego": "IE Mercedes Ábrego", "Mercedes Ábrego": "IE Mercedes Ábrego",
        "IE Mercedes Abrego": "IE Mercedes Ábrego", "IE Mercedes Ábrego": "IE Mercedes Ábrego",
        "Normal Superior": "IE Normal Superior", "IE Normal Superior": "IE Normal Superior",
        "Antonio Narino": "IE Antonio Nariño", "Antonio Nariño": "IE Antonio Nariño",
        "IE Antonio Narino": "IE Antonio Nariño", "IE Antonio Nariño": "IE Antonio Nariño",
        "Cantaclaro": "IE Cantaclaro", "IE Cantaclaro": "IE Cantaclaro",
    }
    return reemplazos.get(texto, texto)


def limpiar_nombres_instituciones(tabla: pd.DataFrame) -> pd.DataFrame:
    tabla = tabla.copy()
    for columna in tabla.select_dtypes(include=["object"]).columns:
        tabla[columna] = tabla[columna].apply(clean_text_value)
    return tabla

@st.cache_data(show_spinner=False)
def load_data(path: Path, cache_version: str):
    estudiantes = pd.read_excel(path, sheet_name="Datos_estudiantes_Moodle")
    escuelas = pd.read_excel(path, sheet_name="Escuelas_mapa")
    logs = pd.read_excel(path, sheet_name="Logs_Moodle_muestra")
    estudiantes = limpiar_nombres_instituciones(estudiantes)
    escuelas = limpiar_nombres_instituciones(escuelas)
    logs = limpiar_nombres_instituciones(logs)
    estudiantes["ultimo_acceso"] = pd.to_datetime(estudiantes["ultimo_acceso"])
    logs["fecha"] = pd.to_datetime(logs["fecha"])
    return estudiantes, escuelas, logs


def pct(num: float, den: float) -> float:
    return 0 if den == 0 or pd.isna(den) else (num / den) * 100


def format_int(n: float) -> str:
    return f"{n:,.0f}".replace(",", ".")


def metric(label: str, value: str, note: str = "") -> str:
    return f"""
    <div class="metric-card">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        <div class="metric-note">{note}</div>
    </div>
    """


def build_school_summary(data: pd.DataFrame) -> pd.DataFrame:
    summary = (
        data.groupby(["codigo_escuela", "institucion", "nombre_oficial", "barrio", "direccion", "latitud", "longitud"], as_index=False)
        .agg(
            estudiantes=("id_estudiante", "count"),
            promedio=("calificacion_final", "mean"),
            accesos_promedio=("accesos_moodle", "mean"),
            dias_activos_promedio=("dias_activos", "mean"),
            minutos_promedio=("minutos_plataforma", "mean"),
            avance_promedio=("avance_curso_pct", "mean"),
            tareas_entregadas=("tareas_entregadas", "sum"),
            tareas_asignadas=("tareas_asignadas", "sum"),
            foros_promedio=("participaciones_foro", "mean"),
            recursos_promedio=("recursos_visualizados", "mean"),
            riesgo_alto=("riesgo_academico", lambda s: (s == "Alto").sum()),
            riesgo_medio=("riesgo_academico", lambda s: (s == "Medio").sum()),
            riesgo_bajo=("riesgo_academico", lambda s: (s == "Bajo").sum()),
        )
    )
    summary["cumplimiento_pct"] = summary.apply(lambda r: pct(r["tareas_entregadas"], r["tareas_asignadas"]), axis=1)
    summary["riesgo_alto_pct"] = summary.apply(lambda r: pct(r["riesgo_alto"], r["estudiantes"]), axis=1)
    summary["horas_total"] = summary["minutos_promedio"] * summary["estudiantes"] / 60
    return summary.round({
        "promedio": 2,
        "accesos_promedio": 1,
        "dias_activos_promedio": 1,
        "minutos_promedio": 0,
        "avance_promedio": 1,
        "cumplimiento_pct": 1,
        "riesgo_alto_pct": 1,
        "foros_promedio": 1,
        "recursos_promedio": 1,
        "horas_total": 0,
    })


def dataframe_to_csv_bytes(data: pd.DataFrame) -> bytes:
    return data.to_csv(index=False).encode("utf-8-sig")

# =============================================================================
# GRAFICOS INTERACTIVOS
# =============================================================================
def build_folium_map(summary: pd.DataFrame, selected_school: str) -> folium.Map:
    """Construye un mapa con marcadores visibles y clicables."""
    center = [float(summary["latitud"].mean()), float(summary["longitud"].mean())]
    fmap = folium.Map(location=center, zoom_start=12, tiles="OpenStreetMap", control_scale=True)
    max_est = max(1, float(summary["estudiantes"].max()))
    bounds = []
    for _, r in summary.iterrows():
        lat, lon = float(r["latitud"]), float(r["longitud"])
        bounds.append([lat, lon])
        risk = float(r["riesgo_alto_pct"])
        if risk >= 30:
            color = icon_color = "red"
        elif risk >= 18:
            color = icon_color = "orange"
        else:
            color = icon_color = "green"
        if r["institucion"] == selected_school:
            color = icon_color = "blue"
        radius = 8 + min(20, float(r["estudiantes"]) / max_est * 20)
        popup_html = f"""
        <div style='font-family: Arial; width: 235px;'>
            <b style='font-size:14px'>{r['institucion']}</b><br>
            <hr style='margin:6px 0'>
            <b>Estudiantes:</b> {format_int(r['estudiantes'])}<br>
            <b>Promedio:</b> {r['promedio']:.2f}<br>
            <b>Cumplimiento:</b> {r['cumplimiento_pct']:.1f}%<br>
            <b>Riesgo alto:</b> {r['riesgo_alto_pct']:.1f}%<br>
            <b>Accesos:</b> {r['accesos_promedio']:.1f}
        </div>
        """
        folium.CircleMarker(
            location=[lat, lon], radius=radius, color=color, fill=True, fill_color=color,
            fill_opacity=0.40, weight=3, tooltip=str(r["institucion"]),
            popup=folium.Popup(popup_html, max_width=280),
        ).add_to(fmap)
        folium.Marker(
            location=[lat, lon], tooltip=str(r["institucion"]),
            popup=folium.Popup(popup_html, max_width=280),
            icon=folium.Icon(color=icon_color, icon="graduation-cap", prefix="fa"),
        ).add_to(fmap)
    if bounds:
        fmap.fit_bounds(bounds, padding=(35, 35))
    legend_html = """
    <div style="position: fixed; bottom: 24px; left: 24px; z-index:9999; background:white; padding:8px 10px; border:1px solid #cbd5e1; border-radius:10px; font-size:12px; font-family:Arial; box-shadow:0 2px 8px rgba(0,0,0,.15);">
        <b>Riesgo alto</b><br>
        <span style='color:green'>●</span> Bajo &nbsp;
        <span style='color:orange'>●</span> Medio &nbsp;
        <span style='color:red'>●</span> Alto<br>
        <span style='color:blue'>●</span> Institución seleccionada
    </div>
    """
    fmap.get_root().html.add_child(folium.Element(legend_html))
    return fmap

def plotly_figures(summary, selected_data, selected_logs, vista_ranking):
    if vista_ranking == "Riesgo alto":
        rank_df = summary.sort_values("riesgo_alto_pct", ascending=True)
        x_col, label, scale = "riesgo_alto_pct", "% riesgo alto", ["#22c55e", "#f59e0b", "#ef4444"]
    elif vista_ranking == "Promedio":
        rank_df = summary.sort_values("promedio", ascending=True)
        x_col, label, scale = "promedio", "promedio final", "Blues"
    elif vista_ranking == "Cumplimiento":
        rank_df = summary.sort_values("cumplimiento_pct", ascending=True)
        x_col, label, scale = "cumplimiento_pct", "% cumplimiento", "Teal"
    else:
        rank_df = summary.sort_values("accesos_promedio", ascending=True)
        x_col, label, scale = "accesos_promedio", "accesos promedio", "Purples"

    rank_fig = px.bar(
        rank_df,
        x=x_col,
        y="institucion",
        orientation="h",
        color=x_col,
        text=x_col,
        color_continuous_scale=scale,
        labels={x_col: label, "institucion": ""},
        height=432,
    )
    rank_fig.update_traces(texttemplate="%{text:.1f}", textposition="outside", marker_line_width=0)
    rank_fig.update_layout(
        margin=dict(l=0, r=18, t=0, b=0),
        coloraxis_showscale=False,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(showgrid=True, gridcolor="rgba(148,163,184,.25)"),
        yaxis=dict(tickfont=dict(size=10)),
    )

    risk_order = ["Alto", "Medio", "Bajo"]
    riesgo_counts = selected_data["riesgo_academico"].value_counts().reindex(risk_order, fill_value=0).reset_index()
    riesgo_counts.columns = ["riesgo", "estudiantes"]
    fig_risk = px.pie(
        riesgo_counts,
        names="riesgo",
        values="estudiantes",
        hole=.58,
        color="riesgo",
        color_discrete_map={"Alto": "#ef4444", "Medio": "#f59e0b", "Bajo": "#22c55e"},
        height=280,
    )
    fig_risk.update_layout(margin=dict(l=0, r=0, t=0, b=0), legend=dict(orientation="h", y=-.08), paper_bgcolor="rgba(0,0,0,0)")

    scatter_data = selected_data.copy()
    scatter_data["cumplimiento_estudiante"] = scatter_data.apply(lambda r: pct(r["tareas_entregadas"], r["tareas_asignadas"]), axis=1)
    fig_scatter = px.scatter(
        scatter_data.sample(min(len(scatter_data), 900), random_state=7),
        x="cumplimiento_estudiante",
        y="calificacion_final",
        color="riesgo_academico",
        size="accesos_moodle",
        hover_name="id_estudiante",
        hover_data=["grado", "grupo", "conectividad", "dispositivo"],
        color_discrete_map={"Alto": "#ef4444", "Medio": "#f59e0b", "Bajo": "#22c55e"},
        labels={"cumplimiento_estudiante": "% tareas", "calificacion_final": "nota"},
        height=280,
    )
    fig_scatter.update_layout(margin=dict(l=0, r=0, t=0, b=0), legend=dict(orientation="h", y=-.22), paper_bgcolor="rgba(0,0,0,0)")

    grade_summary = selected_data.groupby("grado", as_index=False).agg(promedio=("calificacion_final", "mean"), estudiantes=("id_estudiante", "count"))
    fig_grade = px.bar(
        grade_summary,
        x="grado",
        y="promedio",
        color="promedio",
        text="promedio",
        color_continuous_scale="Viridis",
        hover_data=["estudiantes"],
        labels={"grado": "grado", "promedio": "promedio"},
        height=280,
    )
    fig_grade.update_traces(texttemplate="%{text:.2f}", textposition="outside")
    fig_grade.update_layout(margin=dict(l=0, r=0, t=0, b=0), coloraxis_showscale=False, yaxis_range=[0, 5], paper_bgcolor="rgba(0,0,0,0)")

    if selected_logs.empty:
        eventos = pd.DataFrame({"tipo_evento_moodle": [], "registros": []})
    else:
        eventos = selected_logs.groupby("tipo_evento_moodle", as_index=False).size().rename(columns={"size": "registros"})
    fig_events = px.bar(
        eventos.sort_values("registros", ascending=True),
        x="registros",
        y="tipo_evento_moodle",
        orientation="h",
        color="registros",
        color_continuous_scale="Blues",
        labels={"registros": "registros", "tipo_evento_moodle": ""},
        height=280,
    )
    fig_events.update_layout(margin=dict(l=0, r=0, t=0, b=0), coloraxis_showscale=False, paper_bgcolor="rgba(0,0,0,0)")

    return rank_fig, fig_risk, fig_scatter, fig_grade, fig_events

# =============================================================================
# EXPORTACION ROBUSTA PNG/JPG/PDF SIN KALEIDO
# =============================================================================
def get_font(size: int, bold: bool = False):
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf",
        "/Library/Fonts/Arial Bold.ttf" if bold else "/Library/Fonts/Arial.ttf",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            continue
    return ImageFont.load_default()


def wrap_text(draw, text: str, font, max_width: int):
    words = str(text).split()
    lines, current = [], ""
    for word in words:
        test = word if not current else current + " " + word
        if draw.textlength(test, font=font) <= max_width:
            current = test
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines


def draw_panel(draw, box, fill="#ffffff", outline="#d7deea", radius=22, width=2):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def fig_to_pil_mpl(fig) -> Image.Image:
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return Image.open(buf).convert("RGB")


def chart_risk(selected_data):
    counts = selected_data["riesgo_academico"].value_counts().reindex(["Alto", "Medio", "Bajo"], fill_value=0)
    fig, ax = plt.subplots(figsize=(5.2, 3.1))
    colors_m = ["#ef4444", "#f59e0b", "#22c55e"]
    ax.pie(counts.values, labels=counts.index, autopct="%1.1f%%", startangle=90, colors=colors_m, wedgeprops=dict(width=0.45))
    ax.set_title("Riesgo académico", fontsize=12, fontweight="bold")
    return fig_to_pil_mpl(fig)


def chart_scatter(selected_data):
    sample = selected_data.sample(min(len(selected_data), 800), random_state=7).copy()
    sample["cumplimiento"] = sample.apply(lambda r: pct(r["tareas_entregadas"], r["tareas_asignadas"]), axis=1)
    fig, ax = plt.subplots(figsize=(5.2, 3.1))
    color_map = {"Alto": "#ef4444", "Medio": "#f59e0b", "Bajo": "#22c55e"}
    for risk, group in sample.groupby("riesgo_academico"):
        ax.scatter(group["cumplimiento"], group["calificacion_final"], s=18, alpha=0.62, label=risk, c=color_map.get(risk, "#64748b"))
    ax.set_xlabel("% tareas")
    ax.set_ylabel("nota")
    ax.set_title("Cumplimiento vs nota", fontsize=12, fontweight="bold")
    ax.grid(alpha=0.25)
    ax.legend(fontsize=8, loc="lower right")
    return fig_to_pil_mpl(fig)


def chart_grade(selected_data):
    data = selected_data.groupby("grado", as_index=False)["calificacion_final"].mean().sort_values("grado")
    fig, ax = plt.subplots(figsize=(5.2, 3.1))
    bars = ax.bar(data["grado"].astype(str), data["calificacion_final"], color="#2563eb")
    ax.set_ylim(0, 5)
    ax.set_xlabel("grado")
    ax.set_ylabel("promedio")
    ax.set_title("Promedio por grado", fontsize=12, fontweight="bold")
    ax.grid(axis="y", alpha=0.25)
    for b in bars:
        ax.text(b.get_x() + b.get_width()/2, b.get_height()+0.05, f"{b.get_height():.2f}", ha="center", va="bottom", fontsize=8)
    return fig_to_pil_mpl(fig)


def chart_events(selected_logs):
    if selected_logs.empty:
        data = pd.Series(dtype=float)
    else:
        data = selected_logs["tipo_evento_moodle"].value_counts().sort_values()
    fig, ax = plt.subplots(figsize=(5.2, 3.1))
    if len(data) == 0:
        ax.text(0.5, 0.5, "Sin eventos", ha="center", va="center")
        ax.axis("off")
    else:
        ax.barh(data.index.astype(str), data.values, color="#0f766e")
        ax.set_xlabel("registros")
        ax.set_title("Eventos de plataforma", fontsize=12, fontweight="bold")
        ax.grid(axis="x", alpha=0.25)
    return fig_to_pil_mpl(fig)


def chart_ranking(summary, vista_ranking):
    col_map = {"Riesgo alto": "riesgo_alto_pct", "Promedio": "promedio", "Cumplimiento": "cumplimiento_pct", "Accesos": "accesos_promedio"}
    xcol = col_map.get(vista_ranking, "riesgo_alto_pct")
    data = summary.sort_values(xcol, ascending=True)
    fig, ax = plt.subplots(figsize=(5.6, 4.0))
    ax.barh(data["institucion"], data[xcol], color="#0f3d73")
    ax.set_title(f"Ranking: {vista_ranking}", fontsize=12, fontweight="bold")
    ax.grid(axis="x", alpha=0.25)
    ax.tick_params(axis="y", labelsize=8)
    return fig_to_pil_mpl(fig)


def chart_heat(filtered):
    heat = filtered.pivot_table(index="institucion", columns="grado", values="calificacion_final", aggfunc="mean").round(2)
    fig, ax = plt.subplots(figsize=(9.0, 2.8))
    im = ax.imshow(heat.values, aspect="auto", cmap="YlGnBu", vmin=0, vmax=5)
    ax.set_xticks(range(len(heat.columns)))
    ax.set_xticklabels(heat.columns)
    ax.set_yticks(range(len(heat.index)))
    ax.set_yticklabels(heat.index, fontsize=7)
    ax.set_title("Mapa de calor: calificación por grado e institución", fontsize=12, fontweight="bold")
    for i in range(heat.shape[0]):
        for j in range(heat.shape[1]):
            val = heat.iloc[i, j]
            if not pd.isna(val):
                ax.text(j, i, f"{val:.1f}", ha="center", va="center", fontsize=7, color="#0f172a")
    fig.colorbar(im, ax=ax, fraction=0.02, pad=0.02)
    return fig_to_pil_mpl(fig)


def draw_static_map(summary, selected_school):
    W, H = 860, 385
    img = Image.new("RGB", (W, H), "#f8fafc")
    draw = ImageDraw.Draw(img)
    draw.rounded_rectangle((2, 2, W-3, H-3), radius=18, outline="#cbd5e1", width=2, fill="#f8fafc")
    font = get_font(14)
    title = get_font(18, True)
    draw.text((18, 14), "Ubicación de instituciones - Montería", fill="#0f172a", font=title)
    min_lat, max_lat = summary["latitud"].min(), summary["latitud"].max()
    min_lon, max_lon = summary["longitud"].min(), summary["longitud"].max()
    pad = 52
    for _, r in summary.iterrows():
        x = int(pad + (r["longitud"] - min_lon) / (max_lon - min_lon + 1e-9) * (W - 2*pad))
        y = int(H - pad - (r["latitud"] - min_lat) / (max_lat - min_lat + 1e-9) * (H - 2*pad))
        risk = r["riesgo_alto_pct"]
        color = "#ef4444" if risk >= 30 else "#f59e0b" if risk >= 18 else "#22c55e"
        if r["institucion"] == selected_school:
            color = "#2563eb"
            rr = 18
        else:
            rr = 12
        draw.ellipse((x-rr, y-rr, x+rr, y+rr), fill=color, outline="white", width=4)
        label = r["institucion"].replace("IE ", "")[:18]
        draw.text((x+rr+4, y-8), label, fill="#0f172a", font=font)
    draw.text((18, H-28), "Azul: institución seleccionada | Verde/ámbar/rojo: riesgo alto", fill="#475569", font=font)
    return img


def create_dashboard_export_png(summary, filtered, selected_school, selected_row, selected_data, selected_logs, vista_ranking, total_est, prom_final, tareas_pct, riesgo_alto_pct, accesos_prom, horas_total, avance_prom, filtros_resumen):
    W, H = 2400, 1500
    pad = 24
    gap = 16
    canvas = Image.new("RGB", (W, H), "#eef4ff")
    draw = ImageDraw.Draw(canvas)
    title_font = get_font(42, True)
    subtitle_font = get_font(22)
    author_font = get_font(20, True)
    kpi_label_font = get_font(19, True)
    kpi_value_font = get_font(34, True)
    small_font = get_font(17)
    panel_title_font = get_font(24, True)
    body_font = get_font(18)
    body_bold_font = get_font(18, True)
    table_font = get_font(16)
    table_font_bold = get_font(16, True)

    draw.rounded_rectangle((pad, pad, W-pad, 170), radius=30, fill="#0f3d73")
    draw.text((pad+28, pad+22), APP_TITLE, fill="white", font=title_font)
    draw.text((pad+28, pad+76), APP_SUBTITLE, fill="white", font=subtitle_font)
    draw.text((pad+28, pad+108), f"Creado por: {AUTHOR}", fill="#dbeafe", font=author_font)
    y = pad+24
    for line in wrap_text(draw, f"Filtros aplicados: {filtros_resumen}", small_font, W-330)[:3]:
        draw.text((W-760, y), line, fill="#e2e8f0", font=small_font)
        y += 24

    kpis = [
        ("Estudiantes", format_int(total_est), "Muestra filtrada"),
        ("Promedio", f"{prom_final:.2f}", "Escala 0 a 5"),
        ("Cumplimiento", f"{tareas_pct:.1f}%", "Tareas entregadas"),
        ("Riesgo alto", f"{riesgo_alto_pct:.1f}%", "Alerta global"),
        ("Accesos", f"{accesos_prom:.1f}", "Promedio"),
        ("Horas plataforma", format_int(horas_total), "Acumuladas"),
        ("Avance", f"{avance_prom:.1f}%", "Curso virtual"),
    ]
    x0 = pad
    top = 188
    card_w = int((W - 2*pad - gap*6) / 7)
    for label, val, note in kpis:
        draw_panel(draw, (x0, top, x0+card_w, top+118), fill="#ffffff")
        draw.text((x0+16, top+14), label.upper(), fill="#64748b", font=kpi_label_font)
        draw.text((x0+16, top+46), str(val), fill="#0f172a", font=kpi_value_font)
        draw.text((x0+16, top+88), note, fill="#64748b", font=small_font)
        x0 += card_w + gap

    y1, panel_h1 = 328, 450
    left_w, center_w = 900, 560
    right_w = W - 2*pad - left_w - center_w - 2*gap
    left_box = (pad, y1, pad+left_w, y1+panel_h1)
    center_box = (pad+left_w+gap, y1, pad+left_w+gap+center_w, y1+panel_h1)
    right_box = (center_box[2]+gap, y1, center_box[2]+gap+right_w, y1+panel_h1)

    draw_panel(draw, left_box)
    draw.text((left_box[0]+16, left_box[1]+12), "Mapa de instituciones", fill="#0f172a", font=panel_title_font)
    map_img = draw_static_map(summary, selected_school)
    canvas.paste(map_img.resize((860, 385)), (left_box[0]+20, left_box[1]+48))

    draw_panel(draw, center_box)
    draw.text((center_box[0]+16, center_box[1]+12), "Detalle institucional", fill="#0f172a", font=panel_title_font)
    draw.rounded_rectangle((center_box[0]+16, center_box[1]+48, center_box[2]-16, center_box[1]+172), radius=18, fill="#eff6ff", outline="#bfd5ff")
    draw.text((center_box[0]+30, center_box[1]+64), str(selected_row["institucion"]), fill="#111827", font=get_font(28, True))
    info_lines = [
        f"Nombre oficial: {selected_row['nombre_oficial']}",
        f"Barrio: {selected_row['barrio']}",
        f"Dirección: {selected_row['direccion']}",
        f"Coordenadas: {selected_row['latitud']:.5f}, {selected_row['longitud']:.5f}",
    ]
    yy = center_box[1]+102
    for line in info_lines:
        draw.text((center_box[0]+30, yy), line, fill="#475569", font=body_font)
        yy += 22
    cards = [("ESTUDIANTES", format_int(selected_row["estudiantes"])), ("RIESGO ALTO", f"{selected_row['riesgo_alto_pct']:.1f}%"), ("PROMEDIO", f"{selected_row['promedio']:.2f}"), ("CUMPLIMIENTO", f"{selected_row['cumplimiento_pct']:.1f}%")]
    sx, sy = center_box[0]+16, center_box[1]+194
    sw, sh = int((center_w - 48) / 2), 96
    for i, (lab, val) in enumerate(cards):
        row, col = divmod(i, 2)
        bx, by = sx + col*(sw+16), sy + row*(sh+12)
        draw_panel(draw, (bx, by, bx+sw, by+sh), fill="#ffffff")
        draw.text((bx+14, by+12), lab, fill="#64748b", font=kpi_label_font)
        draw.text((bx+14, by+42), str(val), fill="#0f172a", font=get_font(26, True))
    risk_val = float(selected_row["riesgo_alto_pct"])
    fill, outline, txt_color = ("#fff1f2", "#fecdd3", "#7f1d1d") if risk_val >= 25 else ("#ecfdf5", "#bbf7d0", "#166534")
    text = f"Alerta: riesgo alto de {risk_val:.1f}%. Priorizar acompañamiento y recuperación." if risk_val >= 25 else f"Condición controlada: riesgo alto de {risk_val:.1f}%. Mantener seguimiento semanal."
    draw.rounded_rectangle((center_box[0]+16, center_box[1]+408, center_box[2]-16, center_box[3]-16), radius=16, fill=fill, outline=outline)
    ty = center_box[1] + 424
    for line in wrap_text(draw, text, body_bold_font, center_w-60)[:3]:
        draw.text((center_box[0]+28, ty), line, fill=txt_color, font=body_bold_font)
        ty += 22

    draw_panel(draw, right_box)
    draw.text((right_box[0]+16, right_box[1]+12), "Ranking comparativo", fill="#0f172a", font=panel_title_font)
    canvas.paste(chart_ranking(summary, vista_ranking).resize((int(right_w-36), 390)), (right_box[0]+18, right_box[1]+44))

    y2, panel_h2 = 798, 320
    box_w = int((W - 2*pad - 3*gap)/4)
    chart_imgs = [chart_risk(selected_data), chart_scatter(selected_data), chart_grade(selected_data), chart_events(selected_logs)]
    titles = ["Riesgo académico", "Cumplimiento vs nota", "Promedio por grado", "Eventos de plataforma"]
    for i in range(4):
        box = (pad + i*(box_w+gap), y2, pad + i*(box_w+gap)+box_w, y2+panel_h2)
        draw_panel(draw, box)
        draw.text((box[0]+16, box[1]+12), titles[i], fill="#0f172a", font=panel_title_font)
        canvas.paste(chart_imgs[i].resize((box_w-28, 258)), (box[0]+14, box[1]+46))

    y3 = 1138
    box1 = (pad, y3, pad+1220, H-pad)
    box2 = (pad+1220+gap, y3, W-pad, H-pad)
    draw_panel(draw, box1)
    draw.text((box1[0]+16, box1[1]+12), "Mapa de calor: calificación por grado e institución", fill="#0f172a", font=panel_title_font)
    canvas.paste(chart_heat(filtered).resize((1188, 290)), (box1[0]+16, box1[1]+42))

    draw_panel(draw, box2)
    draw.text((box2[0]+16, box2[1]+12), "Estudiantes priorizados", fill="#0f172a", font=panel_title_font)
    draw.text((box2[0]+16, box2[1]+42), f"Institución seleccionada: {selected_school}", fill="#475569", font=body_font)
    priorizados = selected_data[selected_data["riesgo_academico"] == "Alto"].copy()
    priorizados = priorizados.sort_values(["calificacion_final", "accesos_moodle", "avance_curso_pct"]).head(8)
    table_top = box2[1] + 72
    cols = ["Estudiante", "Grado", "Grupo", "Accesos", "Nota", "% avance"]
    col_widths = [150, 80, 80, 90, 80, 100]
    x = box2[0]+16
    draw.rounded_rectangle((x, table_top, box2[2]-16, table_top+28), radius=10, fill="#0f3d73")
    xx = x+8
    for c, cw in zip(cols, col_widths):
        draw.text((xx, table_top+6), c, fill="white", font=table_font_bold)
        xx += cw
    yy = table_top + 34
    if priorizados.empty:
        draw.text((x+8, yy+20), "No hay estudiantes con riesgo alto.", fill="#166534", font=body_bold_font)
    else:
        for idx, (_, r) in enumerate(priorizados.iterrows()):
            fill_row = "#f8fafc" if idx % 2 == 0 else "#ffffff"
            draw.rectangle((x, yy, box2[2]-16, yy+26), fill=fill_row, outline="#e2e8f0")
            vals = [str(r["id_estudiante"]), str(r["grado"]), str(r["grupo"]), f"{r['accesos_moodle']:.0f}", f"{r['calificacion_final']:.2f}", f"{r['avance_curso_pct']:.1f}%"]
            xx = x+8
            for v, cw in zip(vals, col_widths):
                draw.text((xx, yy+5), v, fill="#0f172a", font=table_font)
                xx += cw
            yy += 26

    out = BytesIO()
    canvas.save(out, format="PNG")
    out.seek(0)
    return out.getvalue()


def png_to_jpg(png_bytes: bytes) -> bytes:
    image = Image.open(BytesIO(png_bytes)).convert("RGB")
    output = BytesIO()
    image.save(output, format="JPEG", quality=95, optimize=True)
    output.seek(0)
    return output.getvalue()


def png_to_pdf(png_bytes: bytes) -> bytes:
    buffer = BytesIO()
    page_w, page_h = landscape(A4)
    c = pdfcanvas.Canvas(buffer, pagesize=(page_w, page_h))
    img = Image.open(BytesIO(png_bytes))
    iw, ih = img.size
    margin = 18
    scale = min((page_w - 2*margin)/iw, (page_h - 2*margin)/ih)
    draw_w, draw_h = iw * scale, ih * scale
    x, y = (page_w - draw_w) / 2, (page_h - draw_h) / 2
    c.drawImage(ImageReader(BytesIO(png_bytes)), x, y, width=draw_w, height=draw_h)
    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer.getvalue()


def build_pdf_report(filtered, summary, selected_school, selected_row, selected_data, selected_logs, dashboard_png, total_est, prom_final, tareas_pct, riesgo_alto_pct, accesos_prom, horas_total, avance_prom):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=1.35*cm, leftMargin=1.35*cm, topMargin=1.1*cm, bottomMargin=1.1*cm)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("TitleX", parent=styles["Title"], alignment=TA_CENTER, fontName="Helvetica-Bold", fontSize=16, leading=20, textColor=colors.HexColor("#0f3d73"), spaceAfter=6)
    subtitle_style = ParagraphStyle("SubX", parent=styles["Normal"], alignment=TA_CENTER, fontSize=9, leading=12, textColor=colors.HexColor("#475569"), spaceAfter=8)
    h2 = ParagraphStyle("H2X", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=12, leading=15, textColor=colors.HexColor("#0f172a"), spaceBefore=8, spaceAfter=5)
    body = ParagraphStyle("BodyX", parent=styles["BodyText"], fontSize=9, leading=12, textColor=colors.HexColor("#1f2937"), spaceAfter=6)
    story = []
    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")
    story += [Paragraph(APP_TITLE, title_style), Paragraph(APP_SUBTITLE, subtitle_style), Paragraph(f"Creadora: {AUTHOR}", subtitle_style), Paragraph(f"Informe generado el {fecha}", subtitle_style)]
    story.append(Paragraph("1. Contexto de la situación", h2))
    story.append(Paragraph("El informe resume una situación educativa simulada durante la pandemia en diez instituciones educativas urbanas de Montería, Córdoba. Los estudiantes desarrollaron actividades en una plataforma virtual de aprendizaje y el tablero consolida indicadores de acceso, cumplimiento, rendimiento, avance y riesgo académico para apoyar decisiones pedagógicas.", body))
    story.append(Paragraph("2. Vista general del dashboard", h2))
    # Insertar dashboard completo como imagen reducida
    img_buf = BytesIO(dashboard_png)
    story.append(RLImage(img_buf, width=17.2*cm, height=10.75*cm))
    story.append(Paragraph("3. Indicadores generales", h2))
    rows = [["Indicador", "Valor"], ["Estudiantes analizados", format_int(total_est)], ["Promedio académico", f"{prom_final:.2f}"], ["Cumplimiento de tareas", f"{tareas_pct:.1f}%"], ["Riesgo alto", f"{riesgo_alto_pct:.1f}%"], ["Accesos promedio", f"{accesos_prom:.1f}"], ["Horas acumuladas en plataforma", format_int(horas_total)], ["Avance promedio", f"{avance_prom:.1f}%"]]
    table = Table(rows, colWidths=[8.0*cm, 7.8*cm])
    table.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.HexColor("#0f3d73")), ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("FONTSIZE", (0,0), (-1,-1), 8.5), ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#cbd5e1")), ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f8fafc")])]))
    story.append(table)
    story.append(Paragraph("4. Situación de la institución seleccionada", h2))
    story.append(Paragraph(f"Institución seleccionada: <b>{selected_school}</b>. Promedio: <b>{selected_row['promedio']:.2f}</b>. Cumplimiento: <b>{selected_row['cumplimiento_pct']:.1f}%</b>. Riesgo alto: <b>{selected_row['riesgo_alto_pct']:.1f}%</b>. Accesos promedio: <b>{selected_row['accesos_promedio']:.1f}</b>.", body))
    story.append(Paragraph("5. Ranking de instituciones prioritarias", h2))
    ranking = summary.sort_values("riesgo_alto_pct", ascending=False).head(10)
    rank_rows = [["Institución", "Est.", "Prom.", "% riesgo", "% cumpl."]]
    for _, r in ranking.iterrows():
        rank_rows.append([str(r["institucion"]), format_int(r["estudiantes"]), f"{r['promedio']:.2f}", f"{r['riesgo_alto_pct']:.1f}%", f"{r['cumplimiento_pct']:.1f}%"])
    rank_table = Table(rank_rows, colWidths=[7.0*cm, 2.0*cm, 2.0*cm, 2.4*cm, 2.4*cm])
    rank_table.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1e293b")), ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("FONTSIZE", (0,0), (-1,-1), 7.5), ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#cbd5e1")), ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f8fafc")])]))
    story.append(rank_table)
    story.append(Paragraph("6. Interpretación y decisiones sugeridas", h2))
    high_risk_school = summary.sort_values("riesgo_alto_pct", ascending=False).iloc[0]
    best_avg_school = summary.sort_values("promedio", ascending=False).iloc[0]
    best_compliance_school = summary.sort_values("cumplimiento_pct", ascending=False).iloc[0]
    story.append(Paragraph(f"La institución con mayor proporción de riesgo alto es {high_risk_school['institucion']}, con {high_risk_school['riesgo_alto_pct']:.1f}%. La institución con mejor promedio académico es {best_avg_school['institucion']}, con {best_avg_school['promedio']:.2f}. La institución con mayor cumplimiento es {best_compliance_school['institucion']}, con {best_compliance_school['cumplimiento_pct']:.1f}%.", body))
    story.append(Paragraph("Se recomienda priorizar seguimiento a estudiantes con riesgo alto, fortalecer tutorías, contactar a familias con baja conectividad, revisar la carga de tareas en la plataforma virtual y replicar buenas prácticas de las instituciones con mejores indicadores.", body))
    story.append(PageBreak())
    story.append(Paragraph("7. Gráficos de soporte", h2))
    chart_imgs = [chart_risk(selected_data), chart_scatter(selected_data), chart_grade(selected_data), chart_events(selected_logs), chart_heat(filtered)]
    labels = ["Distribución de riesgo", "Cumplimiento vs nota", "Promedio por grado", "Eventos de plataforma", "Mapa de calor institucional"]
    for label, im in zip(labels, chart_imgs):
        story.append(Paragraph(label, h2))
        tmp = BytesIO(); im.save(tmp, format="PNG"); tmp.seek(0)
        if label == "Mapa de calor institucional":
            story.append(RLImage(tmp, width=16.4*cm, height=5.1*cm))
        else:
            story.append(RLImage(tmp, width=12.0*cm, height=7.0*cm))
        story.append(Spacer(1, 5))
    story.append(Paragraph("8. Estudiantes priorizados", h2))
    priorizados = selected_data[selected_data["riesgo_academico"] == "Alto"].sort_values(["calificacion_final", "accesos_moodle", "avance_curso_pct"]).head(12)
    if priorizados.empty:
        story.append(Paragraph("No hay estudiantes con riesgo alto para los filtros actuales.", body))
    else:
        p_rows = [["Estudiante", "Grado", "Grupo", "Accesos", "Nota", "% avance"]]
        for _, r in priorizados.iterrows():
            p_rows.append([str(r["id_estudiante"]), str(r["grado"]), str(r["grupo"]), f"{r['accesos_moodle']:.0f}", f"{r['calificacion_final']:.2f}", f"{r['avance_curso_pct']:.1f}%"])
        p_table = Table(p_rows, colWidths=[3.7*cm, 2.0*cm, 2.0*cm, 2.2*cm, 2.2*cm, 2.7*cm])
        p_table.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.HexColor("#991b1b")), ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("FONTSIZE", (0,0), (-1,-1), 7.5), ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#cbd5e1")), ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#fff7ed")])]))
        story.append(p_table)
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# =============================================================================
# APP
# =============================================================================

def extraer_institucion_desde_popup(popup_html: str | None) -> str | None:
    """Intenta leer el nombre de la institución desde el HTML del popup."""
    if not popup_html:
        return None
    texto = re.sub(r"<[^>]+>", " ", str(popup_html))
    texto = re.sub(r"\s+", " ", texto).strip()
    parte = texto.split(" Estudiantes:")[0].strip()
    return clean_text_value(parte) if parte else None

if not DATA_PATH.exists():
    st.error("No se encontro el Excel. Coloca 'data_moodle_monteria_dashboard_mapa.xlsx' en la misma carpeta de este script.")
    st.stop()

df, escuelas, logs = load_data(DATA_PATH, CACHE_VERSION)

st.markdown(f"""
<div class="hero">
    <div class="hero-title">{APP_TITLE} · {APP_SUBTITLE}</div>
    <div class="hero-subtitle">
        <span class="chip">Pandemia</span><span class="chip">10 instituciones urbanas</span><span class="chip">Aprendizaje remoto</span><span class="chip">Mapa funcional con Folium</span>
        Dashboard horizontal para apoyar decisiones pedagógicas con indicadores de acceso, cumplimiento, desempeño y alertas tempranas.
    </div>
    <div class="hero-author"><span>Creado por: {AUTHOR}</span></div>
</div>
""", unsafe_allow_html=True)

with st.expander("Filtros del tablero", expanded=True):
    f1, f2, f3, f4, f5 = st.columns([2.2, 1, 1.3, 1.2, 1.1])
    with f1:
        escuelas_sel = st.multiselect("Instituciones", sorted(df["institucion"].unique()), default=sorted(df["institucion"].unique()))
    with f2:
        grados_sel = st.multiselect("Grados", sorted(df["grado"].unique()), default=sorted(df["grado"].unique()))
    with f3:
        riesgo_sel = st.multiselect("Riesgo", ["Alto", "Medio", "Bajo"], default=["Alto", "Medio", "Bajo"])
    with f4:
        conectividad_sel = st.multiselect("Conectividad", sorted(df["conectividad"].unique()), default=sorted(df["conectividad"].unique()))
    with f5:
        vista_ranking = st.selectbox("Ranking", ["Riesgo alto", "Promedio", "Cumplimiento", "Accesos"])

filtered = df[df["institucion"].isin(escuelas_sel) & df["grado"].isin(grados_sel) & df["riesgo_academico"].isin(riesgo_sel) & df["conectividad"].isin(conectividad_sel)].copy()
if filtered.empty:
    st.warning("No hay datos para los filtros seleccionados. Ajusta los filtros del tablero.")
    st.stop()
summary = build_school_summary(filtered)
schools_in_summary = summary.sort_values("institucion")["institucion"].tolist()
default_school = summary.sort_values("riesgo_alto_pct", ascending=False)["institucion"].iloc[0]

if "selected_school" not in st.session_state or st.session_state["selected_school"] not in schools_in_summary:
    st.session_state["selected_school"] = default_school

selected_manual = st.selectbox("Institución en detalle: selecciona aquí o haz clic en el mapa", schools_in_summary, index=schools_in_summary.index(st.session_state["selected_school"]), key="selector_institucion_detalle")
st.session_state["selected_school"] = selected_manual

# KPI globales
total_est = len(filtered)
prom_final = filtered["calificacion_final"].mean()
tareas_pct = pct(filtered["tareas_entregadas"].sum(), filtered["tareas_asignadas"].sum())
riesgo_alto = int((filtered["riesgo_academico"] == "Alto").sum())
riesgo_alto_pct = pct(riesgo_alto, total_est)
accesos_prom = filtered["accesos_moodle"].mean()
horas_total = filtered["minutos_plataforma"].sum() / 60
avance_prom = filtered["avance_curso_pct"].mean()

k1, k2, k3, k4, k5, k6, k7 = st.columns([1,1,1,1,1,1,1])
with k1: st.markdown(metric("Estudiantes", format_int(total_est), "muestra filtrada"), unsafe_allow_html=True)
with k2: st.markdown(metric("Promedio", f"{prom_final:.2f}", "escala 0 a 5"), unsafe_allow_html=True)
with k3: st.markdown(metric("Cumplimiento", f"{tareas_pct:.1f}%", "tareas entregadas"), unsafe_allow_html=True)
with k4: st.markdown(metric("Riesgo alto", f"{riesgo_alto_pct:.1f}%", f"{format_int(riesgo_alto)} estudiantes"), unsafe_allow_html=True)
with k5: st.markdown(metric("Accesos", f"{accesos_prom:.1f}", "promedio por estudiante"), unsafe_allow_html=True)
with k6: st.markdown(metric("Horas plataforma", format_int(horas_total), "acumuladas"), unsafe_allow_html=True)
with k7: st.markdown(metric("Avance", f"{avance_prom:.1f}%", "curso virtual"), unsafe_allow_html=True)

main_left, main_center, main_right = st.columns([1.45, 1.0, 1.05])

with main_left:
    st.markdown('<div class="panel"><div class="panel-title">Mapa interactivo de instituciones</div><div class="panel-help">Haz clic en un marcador o usa el selector. Azul indica la institución seleccionada.</div>', unsafe_allow_html=True)
    fmap = build_folium_map(summary, st.session_state["selected_school"])
    map_data = st_folium(
        fmap,
        height=432,
        use_container_width=True,
        key="folium_mapa_instituciones_v8",
        returned_objects=["last_object_clicked_tooltip", "last_object_clicked_popup", "last_object_clicked"],
    )
    clicked_school = None
    if isinstance(map_data, dict):
        clicked_school = map_data.get("last_object_clicked_tooltip")
        if not clicked_school:
            clicked_school = extraer_institucion_desde_popup(map_data.get("last_object_clicked_popup"))
    clicked_school = clean_text_value(clicked_school) if clicked_school else None
    if clicked_school in schools_in_summary and clicked_school != st.session_state["selected_school"]:
        st.session_state["selected_school"] = clicked_school
        st.rerun()
    st.caption("Si el clic no actualiza de inmediato, selecciona la institución en el selector superior. El mapa tiene marcadores visibles y popups informativos.")
    st.markdown('</div>', unsafe_allow_html=True)

selected_school = st.session_state["selected_school"]
selected_row = summary[summary["institucion"] == selected_school].iloc[0]
selected_data = filtered[filtered["institucion"] == selected_school].copy()
selected_logs = logs[logs["institucion"] == selected_school].copy()
rank_fig, fig_risk, fig_scatter, fig_grade, fig_events = plotly_figures(summary, selected_data, selected_logs, vista_ranking)

with main_center:
    st.markdown('<div class="panel"><div class="panel-title">Detalle de la institución seleccionada</div>', unsafe_allow_html=True)
    st.markdown(f"""
    <div class="institution-card">
        <div class="institution-name">{selected_row['institucion']}</div>
        <div class="institution-meta">
            <b>Nombre oficial:</b> {selected_row['nombre_oficial']}<br>
            <b>Barrio:</b> {selected_row['barrio']}<br>
            <b>Dirección:</b> {selected_row['direccion']}<br>
            <b>Coordenadas:</b> {selected_row['latitud']:.5f}, {selected_row['longitud']:.5f}
        </div>
    </div>
    """, unsafe_allow_html=True)
    d1, d2 = st.columns(2)
    with d1:
        st.markdown(metric("Estudiantes", format_int(selected_row["estudiantes"]), "en filtros actuales"), unsafe_allow_html=True)
        st.markdown(metric("Promedio", f"{selected_row['promedio']:.2f}", "calificación final"), unsafe_allow_html=True)
    with d2:
        st.markdown(metric("Riesgo alto", f"{selected_row['riesgo_alto_pct']:.1f}%", f"{format_int(selected_row['riesgo_alto'])} casos"), unsafe_allow_html=True)
        st.markdown(metric("Cumplimiento", f"{selected_row['cumplimiento_pct']:.1f}%", "tareas virtuales"), unsafe_allow_html=True)
    if selected_row["riesgo_alto_pct"] >= 25:
        st.markdown(f"<div class='alert-high'>Alerta: esta institución concentra un riesgo alto relevante ({selected_row['riesgo_alto_pct']:.1f}%). Priorizar acompañamiento, contacto familiar y recuperación.</div>", unsafe_allow_html=True)
    else:
        st.markdown(f"<div class='alert-ok'>Condición controlada: el riesgo alto es de {selected_row['riesgo_alto_pct']:.1f}%. Mantener seguimiento semanal.</div>", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

with main_right:
    st.markdown('<div class="panel"><div class="panel-title">Ranking comparativo institucional</div><div class="panel-help">Vista municipal comparativa según el ranking seleccionado.</div>', unsafe_allow_html=True)
    st.plotly_chart(rank_fig, use_container_width=True, key="ranking")
    st.markdown('</div>', unsafe_allow_html=True)

g1, g2, g3, g4 = st.columns([1.05, 1.05, 1.05, .95])
with g1:
    st.markdown('<div class="panel"><div class="panel-title">Riesgo académico</div>', unsafe_allow_html=True)
    st.plotly_chart(fig_risk, use_container_width=True, key="riesgo_detalle")
    st.markdown('</div>', unsafe_allow_html=True)
with g2:
    st.markdown('<div class="panel"><div class="panel-title">Cumplimiento vs nota</div><div class="panel-help">Cada punto es un estudiante de la institución seleccionada.</div>', unsafe_allow_html=True)
    st.plotly_chart(fig_scatter, use_container_width=True, key="scatter_detalle")
    st.markdown('</div>', unsafe_allow_html=True)
with g3:
    st.markdown('<div class="panel"><div class="panel-title">Promedio por grado</div>', unsafe_allow_html=True)
    st.plotly_chart(fig_grade, use_container_width=True, key="grado_detalle")
    st.markdown('</div>', unsafe_allow_html=True)
with g4:
    st.markdown('<div class="panel"><div class="panel-title">Eventos de plataforma</div>', unsafe_allow_html=True)
    st.plotly_chart(fig_events, use_container_width=True, key="eventos_detalle")
    st.markdown('</div>', unsafe_allow_html=True)

bottom_left, bottom_right = st.columns([1.25, 1.0])
with bottom_left:
    st.markdown('<div class="panel"><div class="panel-title">Mapa de calor: calificación por grado e institución</div><div class="panel-help">Vista municipal comparativa.</div>', unsafe_allow_html=True)
    heat = filtered.pivot_table(index="institucion", columns="grado", values="calificacion_final", aggfunc="mean").round(2)
    fig_heat = px.imshow(heat, text_auto=True, aspect="auto", color_continuous_scale="YlGnBu", labels=dict(x="Grado", y="Institución", color="Promedio"), height=330)
    fig_heat.update_layout(margin=dict(l=0, r=0, t=0, b=0), paper_bgcolor="rgba(0,0,0,0)")
    st.plotly_chart(fig_heat, use_container_width=True, key="heat")
    st.markdown('</div>', unsafe_allow_html=True)
with bottom_right:
    st.markdown('<div class="panel"><div class="panel-title">Estudiantes priorizados</div><div class="panel-help">Casos de la institución seleccionada con riesgo alto y baja actividad.</div>', unsafe_allow_html=True)
    priorizados = selected_data[selected_data["riesgo_academico"] == "Alto"].copy()
    priorizados = priorizados.sort_values(["calificacion_final", "accesos_moodle", "avance_curso_pct"]).head(10)
    if priorizados.empty:
        st.success("No hay estudiantes con riesgo alto para los filtros actuales en esta institución.")
    else:
        st.dataframe(priorizados[["id_estudiante", "grado", "grupo", "conectividad", "accesos_moodle", "tareas_entregadas", "calificacion_final", "avance_curso_pct"]].rename(columns={"id_estudiante":"Estudiante","grado":"Grado","grupo":"Grupo","conectividad":"Conectividad","accesos_moodle":"Accesos","tareas_entregadas":"Tareas","calificacion_final":"Nota","avance_curso_pct":"% avance"}), use_container_width=True, hide_index=True, height=330)
    st.markdown('</div>', unsafe_allow_html=True)

# Descargas
st.markdown('<div class="panel"><div class="panel-title">Descargas del dashboard</div><div class="panel-help">Exporta datos, informe ejecutivo o una versión horizontal del tablero en PNG, JPG o PDF. La imagen no depende de captura del navegador.</div>', unsafe_allow_html=True)
filtros_resumen = f"Instituciones: {len(escuelas_sel)} seleccionadas | Grados: {', '.join(map(str, grados_sel))} | Riesgo: {', '.join(riesgo_sel)} | Conectividad: {', '.join(conectividad_sel)}"

try:
    dashboard_png = create_dashboard_export_png(summary, filtered, selected_school, selected_row, selected_data, selected_logs, vista_ranking, total_est, prom_final, tareas_pct, riesgo_alto_pct, accesos_prom, horas_total, avance_prom, filtros_resumen)
    dashboard_jpg = png_to_jpg(dashboard_png)
    dashboard_pdf = png_to_pdf(dashboard_png)
    report_pdf = build_pdf_report(filtered, summary, selected_school, selected_row, selected_data, selected_logs, dashboard_png, total_est, prom_final, tareas_pct, riesgo_alto_pct, accesos_prom, horas_total, avance_prom)
    export_error = None
except Exception as e:
    dashboard_png = dashboard_jpg = dashboard_pdf = report_pdf = None
    export_error = str(e)

r1c1, r1c2, r1c3, r1c4 = st.columns([1, 1, 1, 1.15])
with r1c1:
    st.download_button("Descargar Excel base", data=DATA_PATH.read_bytes(), file_name="data_moodle_monteria_dashboard_mapa.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
with r1c2:
    st.download_button("Descargar datos filtrados CSV", data=dataframe_to_csv_bytes(filtered), file_name="datos_filtrados_moodle_monteria.csv", mime="text/csv", use_container_width=True)
with r1c3:
    st.download_button("Descargar resumen institucional CSV", data=dataframe_to_csv_bytes(summary), file_name="resumen_institucional_moodle_monteria.csv", mime="text/csv", use_container_width=True)
with r1c4:
    if report_pdf:
        st.download_button("Descargar informe PDF completo", data=report_pdf, file_name=f"informe_completo_moodle_monteria_{selected_school.replace(' ', '_')}.pdf", mime="application/pdf", use_container_width=True)
    else:
        st.button("Descargar informe PDF completo", disabled=True, use_container_width=True)

r2c1, r2c2, r2c3 = st.columns([1, 1, 1])
with r2c1:
    if dashboard_png:
        st.download_button("Descargar dashboard PNG", data=dashboard_png, file_name=f"dashboard_moodle_monteria_{selected_school.replace(' ', '_')}.png", mime="image/png", use_container_width=True)
    else:
        st.button("Descargar dashboard PNG", disabled=True, use_container_width=True)
with r2c2:
    if dashboard_jpg:
        st.download_button("Descargar dashboard JPG", data=dashboard_jpg, file_name=f"dashboard_moodle_monteria_{selected_school.replace(' ', '_')}.jpg", mime="image/jpeg", use_container_width=True)
    else:
        st.button("Descargar dashboard JPG", disabled=True, use_container_width=True)
with r2c3:
    if dashboard_pdf:
        st.download_button("Descargar dashboard PDF", data=dashboard_pdf, file_name=f"dashboard_moodle_monteria_{selected_school.replace(' ', '_')}.pdf", mime="application/pdf", use_container_width=True)
    else:
        st.button("Descargar dashboard PDF", disabled=True, use_container_width=True)

if export_error:
    st.warning(f"No se pudo generar la exportación del dashboard: {export_error}")
else:
    st.caption("Las exportaciones PNG/JPG/PDF se generan en formato horizontal y compactan el contenido principal del tablero en un solo lienzo.")
st.markdown('</div>', unsafe_allow_html=True)

st.caption("El mapa ahora usa Folium para capturar el clic en los marcadores. Si el navegador demora en actualizar, usa el selector manual ubicado sobre el tablero; ambos mecanismos alimentan los mismos gráficos de detalle.")
