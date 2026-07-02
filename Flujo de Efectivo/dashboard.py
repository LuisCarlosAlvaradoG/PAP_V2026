import pandas as pd
import streamlit as st
from datetime import date, timedelta
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# ============================================================
#  CONFIGURACIÓN GENERAL
# ============================================================
ARCHIVO_DATOS = "Datos flujo de efectivo.xlsx"
ARCHIVO_SALIDA = "Flujo de efectivo.xlsx"

MESES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
    5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
    9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
}

# Logo oficial de FV Procesados, incrustado en base64 (fondo blanco removido)
# para que la app sea un solo archivo y no dependa de imágenes externas.
LOGO_BASE64 = "https://i.imgur.com/SNggzEh.png"

# ============================================================
#  SISTEMA DE DISEÑO — derivado del logo FV Procesados
# ============================================================
VERDE_OSCURO   = "#3D5A28"
VERDE_MEDIO    = "#5B7F3B"
VERDE_OLIVA    = "#8A9A4E"
VERDE_CLARO    = "#EEF3E3"
ROJO_FRESA     = "#D6453D"
CREMA_FONDO    = "#FBFAF9"
SIDEBAR_BG     = "#24331A"
TINTA          = "#22301A"

# ============================================================
#  FUNCIONES AUXILIARES
# ============================================================
def formato_semana(inicio, fin):
    if inicio.month == fin.month:
        return f"{inicio.day:02d}-{fin.day:02d} {MESES[fin.month]}"
    return f"{inicio.day:02d} {MESES[inicio.month]}-{fin.day:02d} {MESES[fin.month]}"


def formato_titulo(texto):
    return str(texto).strip().title()


def obtener_categorias(df):
    return sorted(df["categoría"].dropna().astype(str).str.strip().unique())


def obtener_conceptos(df):
    return sorted(df["descripción"].dropna().astype(str).str.strip().unique())


def obtener_tipo_por_categoria(categoria):
    if categoria in ["Ventas", "Préstamo"]:
        return "Ingreso"
    return "Egreso"


def buscar_similares(texto, opciones):
    texto = texto.upper().strip()
    palabras = texto.split()
    if texto == "":
        return []
    coincidencias = []
    for opcion in opciones:
        opcion_mayus = str(opcion).upper().strip()
        if texto in opcion_mayus:
            coincidencias.append(opcion)
        elif palabras and all(palabra in opcion_mayus for palabra in palabras):
            coincidencias.append(opcion)
        elif palabras and any(palabra in opcion_mayus for palabra in palabras):
            coincidencias.append(opcion)
    return list(dict.fromkeys(coincidencias))[:8]


def generar_flujo(df):
    df = df.copy()
    df["inicio_semana"] = pd.to_datetime(df["inicio_semana"], dayfirst=True)
    df["tipo"] = df["tipo"].astype(str).str.strip()
    df["categoría"] = df["categoría"].astype(str).str.strip()
    df["descripción"] = df["descripción"].astype(str).str.strip()
    df["importe"] = pd.to_numeric(df["importe"], errors="coerce").fillna(0)
    df["importe_flujo"] = df["importe"]
    df.loc[df["tipo"] == "Egreso", "importe_flujo"] *= -1

    flujo = (
        df.groupby(["categoría", "inicio_semana"])["importe_flujo"]
        .sum()
        .unstack(fill_value=0)
        .sort_index()
    )

    total_ingresos = (
        df[df["tipo"] == "Ingreso"]
        .groupby("inicio_semana")["importe"]
        .sum()
        .reindex(flujo.columns, fill_value=0)
    )
    total_egresos = (
        df[df["tipo"] == "Egreso"]
        .groupby("inicio_semana")["importe"]
        .sum()
        .reindex(flujo.columns, fill_value=0)
    )
    flujo_neto = total_ingresos - total_egresos
    saldo_inicial = 0
    saldo = saldo_inicial + flujo_neto.cumsum()

    flujo.loc["TOTAL INGRESOS"] = total_ingresos
    flujo.loc["TOTAL EGRESOS"] = -total_egresos
    flujo.loc["FLUJO NETO"] = flujo_neto
    flujo.loc["SALDO"] = saldo

    nombres_semanas = (
        df[["inicio_semana", "semana"]]
        .drop_duplicates()
        .set_index("inicio_semana")["semana"]
        .to_dict()
    )
    flujo.columns = [nombres_semanas[col] for col in flujo.columns]
    flujo.index.name = "Categoría"
    flujo = flujo.round(2)

    with pd.ExcelWriter(ARCHIVO_SALIDA, engine="openpyxl") as writer:
        flujo.to_excel(writer, sheet_name="Flujo")
        ws = writer.sheets["Flujo"]

        verde_encabezado = PatternFill(fill_type="solid", fgColor=VERDE_OSCURO.replace("#", ""))
        verde_claro_fill = PatternFill(fill_type="solid", fgColor=VERDE_CLARO.replace("#", ""))
        verde_flujo_fill = PatternFill(fill_type="solid", fgColor=VERDE_MEDIO.replace("#", ""))

        borde = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

        for row in ws.iter_rows():
            for cell in row:
                cell.border = borde
                cell.alignment = Alignment(horizontal="center", vertical="center")

        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = verde_encabezado

        for row in ws.iter_rows():
            if row[0].value in ["TOTAL INGRESOS", "TOTAL EGRESOS", "SALDO"]:
                for cell in row:
                    cell.font = Font(bold=True, color=VERDE_OSCURO.replace("#", ""))
                    cell.fill = verde_claro_fill
            if row[0].value == "FLUJO NETO":
                for cell in row:
                    cell.font = Font(bold=True, size=13, color="FFFFFF")
                    cell.fill = verde_flujo_fill

        for row in ws.iter_rows(min_row=2, min_col=2):
            for cell in row:
                cell.number_format = '#,##0.00'

        for col in ws.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max_length + 2

    return flujo


def actualizar_fin():
    st.session_state.fin_semana = st.session_state.inicio_semana + timedelta(days=5)


def moneda(valor):
    signo = "-" if valor < 0 else ""
    return f"{signo}${abs(valor):,.2f}"


# ============================================================
#  PÁGINA + SISTEMA DE DISEÑO
# ============================================================
st.set_page_config(
    page_title="FV Procesados | Sistema Financiero",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown(f"""
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap" rel="stylesheet">

<style>
    html, body, [class*="css"] {{
        font-family: 'IBM Plex Sans', sans-serif;
        color: {TINTA};
    }}
    .stApp {{
        background-color: {CREMA_FONDO};
    }}
    #MainMenu, header[data-testid="stHeader"] {{
        background-color: transparent;
    }}

    /* -------- Tipografía de títulos -------- */
    h1, h2, h3 {{
        font-family: 'Fraunces', serif;
        color: {TINTA} !important;
        letter-spacing: -0.01em;
    }}

    /* -------- La regla-firma (línea del logo) -------- */
    .regla-marca {{
        height: 2px;
        background-color: {VERDE_OLIVA};
        width: 100%;
        margin: 6px 0 22px 0;
        position: relative;
    }}
    .regla-marca::before {{
        content: "";
        position: absolute;
        left: 0;
        top: -3px;
        width: 8px;
        height: 8px;
        border-radius: 50%;
        background-color: {ROJO_FRESA};
    }}

    /* ================= SIDEBAR ================= */
    section[data-testid="stSidebar"] {{
        background-color: {SIDEBAR_BG};
        border-right: none;
    }}
    section[data-testid="stSidebar"] * {{
        color: {VERDE_CLARO};
    }}
    .marca-wordmark {{
        padding: 26px 24px 0 24px;
        text-align: center;
    }}
    .marca-wordmark .logo-sidebar {{
        width: 100%;
        max-width: 168px;
    }}
    .marca-wordmark .regla-marca {{
        margin: 18px 0 24px 0;
    }}

    .nav-eyebrow {{
        color: {VERDE_OLIVA};
        text-transform: uppercase;
        font-size: 11px;
        letter-spacing: 2px;
        font-weight: 600;
        padding: 4px 20px 8px 20px;
    }}

    section[data-testid="stSidebar"] .stButton > button {{
        background-color: transparent;
        border: none;
        border-left: 3px solid transparent;
        border-radius: 0;
        text-align: left;
        justify-content: flex-start;
        width: 100%;
        padding: 10px 20px;
        font-weight: 500;
        font-size: 15px;
        color: {VERDE_CLARO};
    }}
    section[data-testid="stSidebar"] .stButton > button:hover {{
        background-color: rgba(255,255,255,0.07);
        color: #FFFFFF;
        border-left: 3px solid {VERDE_OLIVA};
    }}
    section[data-testid="stSidebar"] .stButton > button:focus:not(:active) {{
        color: #FFFFFF;
    }}

    .nav-item-activo {{
        border-left: 3px solid {ROJO_FRESA};
        background-color: rgba(255,255,255,0.09);
        padding: 10px 17px;
        font-weight: 700;
        font-size: 15px;
        color: #FFFFFF;
        margin-bottom: 1px;
    }}
    .nav-item-inactivo-sep {{
        height: 10px;
    }}

    .nav-sub-activo {{
        border-left: 3px solid {ROJO_FRESA};
        background-color: rgba(255,255,255,0.06);
        padding: 10px 17px 10px 31px;
        font-weight: 600;
        font-size: 14px;
        color: #FFFFFF;
        margin-top: 1px !important;
        margin-bottom: 6px !important;
        display: block !important;
    }}
    section[data-testid="stSidebar"] .nav-sub .stButton > button {{
        padding-left: 34px;
        font-size: 14px;
        font-weight: 400;
        margin-top: 4px !important;
        margin-bottom: 4px !important;
    }}

    /* ================= CONTENIDO PRINCIPAL ================= */
    .encabezado-modulo {{
        margin-bottom: 4px;
    }}
    .encabezado-modulo .eyebrow {{
        color: {VERDE_OLIVA};
        text-transform: uppercase;
        letter-spacing: 2px;
        font-size: 12px;
        font-weight: 600;
    }}
    .encabezado-modulo h1 {{
        margin: 4px 0 0 0 !important;
        font-size: 30px !important;
    }}

    /* -------- Tarjetas KPI -------- */
    .kpi-tarjeta {{
        background-color: #FFFFFF;
        border: 1px solid rgba(0,0,0,0.06);
        border-left: 4px solid {VERDE_MEDIO};
        border-radius: 6px;
        padding: 16px 18px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }}
    .kpi-tarjeta .kpi-label {{
        text-transform: uppercase;
        font-size: 11px;
        letter-spacing: 1.5px;
        font-weight: 600;
        color: {VERDE_OLIVA};
    }}
    .kpi-tarjeta .kpi-valor {{
        font-family: 'IBM Plex Mono', monospace;
        font-size: 24px;
        font-weight: 600;
        color: {TINTA};
        margin-top: 4px;
    }}
    .kpi-egreso {{ border-left-color: {ROJO_FRESA}; }}
    .kpi-saldo {{ border-left-color: {VERDE_OSCURO}; }}

    /* -------- Contenedores tipo tarjeta (formularios, tabla) -------- */
    /* Se usan con st.container(border=True): así los widgets quedan
       realmente adentro, en vez de un <div> vacío flotando. */
    div[data-testid="stVerticalBlockBorderWrapper"] {{
        background-color: #FFFFFF;
        border-color: rgba(0,0,0,0.08) !important;
        border-radius: 8px !important;
    }}
    div[data-testid="stVerticalBlockBorderWrapper"] > div {{
        border-radius: 8px !important;
    }}

    /* -------- Pantalla de bienvenida: solo el logo, centrado -------- */
    .bienvenida-hero {{
        text-align: center;
        display: flex;
        align-items: center;
        justify-content: center;
        min-height: 70vh;
    }}
    .bienvenida-hero .logo-hero {{
        width: 100%;
        max-width: 420px;
    }}

    /* -------- Botones de la app -------- */
    .stButton > button {{
        border-radius: 6px;
        font-weight: 600;
    }}
    div[data-testid="stAppViewContainer"] .stButton > button {{
        background-color: {VERDE_MEDIO};
        color: white;
        border: none;
    }}
    div[data-testid="stAppViewContainer"] .stButton > button:hover {{
        background-color: {VERDE_OSCURO};
        color: white;
    }}
    .stDownloadButton > button {{
        background-color: {ROJO_FRESA};
        color: white;
        border: none;
        border-radius: 6px;
        font-weight: 700;
    }}
    .stDownloadButton > button:hover {{
        background-color: #B8342D;
        color: white;
    }}

    div[data-testid="stAlert"] {{
        border-radius: 8px;
    }}
</style>
""", unsafe_allow_html=True)

# ============================================================
#  ESTADO DE NAVEGACIÓN
# ============================================================
if "modulo_activo" not in st.session_state:
    st.session_state.modulo_activo = None
if "seccion_flujo" not in st.session_state:
    st.session_state.seccion_flujo = None

if "inicio_semana" not in st.session_state:
    st.session_state.inicio_semana = date.today()
if "fin_semana" not in st.session_state:
    st.session_state.fin_semana = date.today() + timedelta(days=5)
if "concepto_seleccionado" not in st.session_state:
    st.session_state.concepto_seleccionado = None
if "categoria_seleccionada" not in st.session_state:
    st.session_state.categoria_seleccionada = None


def ir_a(modulo, seccion=None):
    st.session_state.modulo_activo = modulo
    if seccion is not None:
        st.session_state.seccion_flujo = seccion
    st.rerun()


# ============================================================
#  SIDEBAR — NAVEGACIÓN
# ============================================================
with st.sidebar:
    st.markdown(f"""
    <div class="marca-wordmark">
        <img src="{LOGO_BASE64}" class="logo-sidebar" alt="FV Procesados">
        <div class="regla-marca"></div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="nav-eyebrow">Módulos</div>', unsafe_allow_html=True)

    # ---- Módulo: Flujo de Efectivo ----
    if st.session_state.modulo_activo == "flujo":
        st.markdown('<div class="nav-item-activo">Flujo de Efectivo</div>', unsafe_allow_html=True)

        st.markdown('<div class="nav-sub">', unsafe_allow_html=True)
        secciones = [
            ("ver", "Ver flujo de efectivo"),
            ("agregar", "Agregar movimiento"),
            ("eliminar", "Eliminar movimiento"),
        ]
        for clave, etiqueta in secciones:
            if st.session_state.seccion_flujo == clave:
                st.markdown(f'<div class="nav-sub-activo">{etiqueta}</div>', unsafe_allow_html=True)
            else:
                if st.button(etiqueta, key=f"nav_{clave}", use_container_width=True):
                    ir_a("flujo", clave)
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        if st.button("Flujo de Efectivo", key="nav_modulo_flujo", use_container_width=True):
            ir_a("flujo", "ver")

    st.markdown('<div class="nav-item-inactivo-sep"></div>', unsafe_allow_html=True)

    # ---- Módulo: Optimización ----
    if st.session_state.modulo_activo == "optimizacion":
        st.markdown('<div class="nav-item-activo">Optimización</div>', unsafe_allow_html=True)
    else:
        if st.button("Optimización", key="nav_modulo_optimizacion", use_container_width=True):
            ir_a("optimizacion")

# ============================================================
#  CARGA DE DATOS
# ============================================================
df = pd.read_excel(ARCHIVO_DATOS)
df["inicio_semana"] = pd.to_datetime(df["inicio_semana"], dayfirst=True)
df["tipo"] = df["tipo"].astype(str).str.strip()
df["categoría"] = df["categoría"].astype(str).str.strip()
df["descripción"] = df["descripción"].astype(str).str.strip()

TITULOS_SECCION = {
    "ver": "Ver flujo de efectivo",
    "agregar": "Agregar movimiento",
    "eliminar": "Eliminar movimiento",
}

# ============================================================
#  MÓDULO 1: FLUJO DE EFECTIVO
# ============================================================
if st.session_state.modulo_activo == "flujo":

    seccion = st.session_state.seccion_flujo

    st.markdown(f"""
    <div class="encabezado-modulo">
        <div class="eyebrow">Flujo de Efectivo</div>
        <h1>{TITULOS_SECCION[seccion]}</h1>
    </div>
    <div class="regla-marca"></div>
    """, unsafe_allow_html=True)

    # ---------------- VER FLUJO DE EFECTIVO ----------------
    if seccion == "ver":
        df_actual = pd.read_excel(ARCHIVO_DATOS)
        df_actual["tipo"] = df_actual["tipo"].astype(str).str.strip()
        df_actual["importe"] = pd.to_numeric(df_actual["importe"], errors="coerce").fillna(0)

        flujo = generar_flujo(df_actual)

        total_ingresos = df_actual.loc[df_actual["tipo"] == "Ingreso", "importe"].sum()
        total_egresos = df_actual.loc[df_actual["tipo"] == "Egreso", "importe"].sum()
        flujo_neto = total_ingresos - total_egresos
        saldo_actual = flujo.loc["SALDO"].iloc[-1] if flujo.shape[1] > 0 else 0

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f"""
            <div class="kpi-tarjeta">
                <div class="kpi-label">Ingresos totales</div>
                <div class="kpi-valor">{moneda(total_ingresos)}</div>
            </div>""", unsafe_allow_html=True)
        with c2:
            st.markdown(f"""
            <div class="kpi-tarjeta kpi-egreso">
                <div class="kpi-label">Egresos totales</div>
                <div class="kpi-valor">{moneda(total_egresos)}</div>
            </div>""", unsafe_allow_html=True)
        with c3:
            st.markdown(f"""
            <div class="kpi-tarjeta">
                <div class="kpi-label">Flujo neto</div>
                <div class="kpi-valor">{moneda(flujo_neto)}</div>
            </div>""", unsafe_allow_html=True)
        with c4:
            st.markdown(f"""
            <div class="kpi-tarjeta kpi-saldo">
                <div class="kpi-label">Saldo actual</div>
                <div class="kpi-valor">{moneda(saldo_actual)}</div>
            </div>""", unsafe_allow_html=True)

        st.write("")
        with st.container(border=True):
            st.dataframe(flujo, use_container_width=True)

        with open(ARCHIVO_SALIDA, "rb") as archivo:
            st.download_button(
                label="Descargar Excel",
                data=archivo,
                file_name=ARCHIVO_SALIDA,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # ---------------- AGREGAR MOVIMIENTO ----------------
    elif seccion == "agregar":
        with st.container(border=True):

            col1, col2 = st.columns(2)
            with col1:
                inicio = st.date_input(
                    "Inicio de semana",
                    key="inicio_semana",
                    on_change=actualizar_fin
                )
            with col2:
                fin = st.date_input(
                    "Fin de semana",
                    key="fin_semana"
                )

            semana = formato_semana(inicio, fin)
            st.info(f"Semana seleccionada: {semana}")

            importe = st.number_input("Importe", min_value=0.0, step=100.0)

            st.markdown('<div class="regla-marca"></div>', unsafe_allow_html=True)

            conceptos_existentes = obtener_conceptos(df)
            categorias_existentes = obtener_categorias(df)

            modo_concepto = st.radio(
                "Concepto",
                ["Usar concepto existente", "Agregar concepto nuevo"],
                horizontal=True
            )

            concepto = ""
            categoria = ""
            tipo = ""

            if modo_concepto == "Usar concepto existente":
                concepto = st.selectbox(
                    "Selecciona el concepto",
                    [""] + conceptos_existentes
                )
                if concepto != "":
                    datos_concepto = df[
                        df["descripción"].astype(str).str.strip() == concepto
                    ].iloc[0]
                    categoria = datos_concepto["categoría"]
                    tipo = datos_concepto["tipo"]
                    st.text_input("Categoría", value=categoria, disabled=True)
                    st.text_input("Tipo", value=tipo, disabled=True)
            else:
                concepto_escrito = st.text_input(
                    "Nuevo concepto",
                    placeholder="Escribe el concepto"
                ).strip()

                if st.session_state.concepto_seleccionado is not None:
                    concepto = st.session_state.concepto_seleccionado
                    datos_concepto = df[
                        df["descripción"].astype(str).str.strip() == concepto
                    ].iloc[0]
                    categoria = datos_concepto["categoría"]
                    tipo = datos_concepto["tipo"]
                    st.text_input("Concepto seleccionado", value=concepto, disabled=True)
                    st.text_input("Categoría", value=categoria, disabled=True)
                    st.text_input("Tipo", value=tipo, disabled=True)
                else:
                    concepto = formato_titulo(concepto_escrito)
                    coincidencias_concepto = buscar_similares(
                        concepto_escrito,
                        conceptos_existentes
                    )
                    if concepto_escrito and coincidencias_concepto:
                        st.caption("Coincidencias encontradas:")
                        columnas = st.columns(4)
                        for i, sugerencia in enumerate(coincidencias_concepto):
                            with columnas[i % 4]:
                                if st.button(sugerencia, key=f"concepto_{i}"):
                                    st.session_state.concepto_seleccionado = sugerencia
                                    st.rerun()

                    modo_categoria = st.radio(
                        "Categoría",
                        ["Usar categoría existente", "Agregar categoría nueva"],
                        horizontal=True
                    )

                    if modo_categoria == "Usar categoría existente":
                        categoria = st.selectbox(
                            "Selecciona la categoría",
                            [""] + categorias_existentes
                        )
                        if categoria != "":
                            tipo = obtener_tipo_por_categoria(categoria)
                            st.text_input("Tipo", value=tipo, disabled=True)
                    else:
                        categoria_escrita = st.text_input(
                            "Nueva categoría",
                            placeholder="Escribe la categoría"
                        ).strip()

                        if st.session_state.categoria_seleccionada is not None:
                            categoria = st.session_state.categoria_seleccionada
                            tipo = obtener_tipo_por_categoria(categoria)
                            st.text_input("Categoría seleccionada", value=categoria, disabled=True)
                            st.text_input("Tipo", value=tipo, disabled=True)
                        else:
                            categoria = formato_titulo(categoria_escrita)
                            coincidencias_categoria = buscar_similares(
                                categoria_escrita,
                                categorias_existentes
                            )
                            if categoria_escrita and coincidencias_categoria:
                                st.caption("Coincidencias encontradas:")
                                columnas_cat = st.columns(4)
                                for i, sugerencia in enumerate(coincidencias_categoria):
                                    with columnas_cat[i % 4]:
                                        if st.button(sugerencia, key=f"categoria_{i}"):
                                            st.session_state.categoria_seleccionada = sugerencia
                                            st.rerun()
                            if categoria_escrita:
                                tipo = st.selectbox("Tipo", ["Ingreso", "Egreso"], index=1)

            if st.button("Guardar movimiento"):
                if concepto == "":
                    st.error("Debes seleccionar o escribir un concepto.")
                elif importe <= 0:
                    st.error("El importe debe ser mayor a 0.")
                elif categoria == "":
                    st.error("Debes seleccionar o escribir una categoría.")
                elif tipo == "":
                    st.error("Debes seleccionar o confirmar el tipo.")
                elif fin < inicio:
                    st.error("La fecha final no puede ser anterior a la fecha inicial.")
                else:
                    nuevo = pd.DataFrame([{
                        "inicio_semana": pd.to_datetime(inicio),
                        "semana": semana,
                        "tipo": tipo,
                        "categoría": categoria,
                        "descripción": concepto,
                        "importe": importe
                    }])
                    df_actualizado = pd.concat([df, nuevo], ignore_index=True)
                    df_actualizado.to_excel(ARCHIVO_DATOS, index=False)
                    generar_flujo(df_actualizado)
                    st.session_state.concepto_seleccionado = None
                    st.session_state.categoria_seleccionada = None
                    st.success("Movimiento agregado correctamente y flujo actualizado.")


    # ---------------- ELIMINAR MOVIMIENTO ----------------
    elif seccion == "eliminar":
        with st.container(border=True):

            df_eliminar = pd.read_excel(ARCHIVO_DATOS)
            df_eliminar["inicio_semana"] = pd.to_datetime(
                df_eliminar["inicio_semana"],
                dayfirst=True
            )
            df_eliminar["importe"] = pd.to_numeric(
                df_eliminar["importe"],
                errors="coerce"
            ).fillna(0)
            df_eliminar["movimiento"] = (
                df_eliminar["semana"].astype(str)
                + " | "
                + df_eliminar["tipo"].astype(str)
                + " | "
                + df_eliminar["categoría"].astype(str)
                + " | "
                + df_eliminar["descripción"].astype(str)
                + " | $"
                + df_eliminar["importe"].map("{:,.2f}".format)
            )

            movimiento = st.selectbox(
                "Selecciona el movimiento que deseas eliminar",
                [""] + df_eliminar["movimiento"].tolist()
            )
            confirmar = st.checkbox("Confirmo que quiero eliminar este movimiento")

            if st.button("Eliminar movimiento"):
                if movimiento == "":
                    st.error("Debes seleccionar un movimiento.")
                elif not confirmar:
                    st.error("Debes confirmar antes de eliminar.")
                else:
                    indice = df_eliminar[
                        df_eliminar["movimiento"] == movimiento
                    ].index[0]
                    df_eliminar = df_eliminar.drop(index=indice)
                    df_eliminar = df_eliminar.drop(columns=["movimiento"])
                    df_eliminar.to_excel(ARCHIVO_DATOS, index=False)
                    generar_flujo(df_eliminar)
                    st.success("Movimiento eliminado correctamente y flujo actualizado.")


# ============================================================
#  MÓDULO 2: OPTIMIZACIÓN
# ============================================================
elif st.session_state.modulo_activo == "optimizacion":
    st.markdown(f"""
    <div class="encabezado-modulo">
        <div class="eyebrow">Módulo</div>
        <h1>Optimización</h1>
    </div>
    <div class="regla-marca"></div>
    """, unsafe_allow_html=True)

    # A partir de aquí puede agregarse el código de optimización.
    # Para mantener el mismo lenguaje visual, envuelve cada bloque en:
    #
    # with st.container(border=True):
    #     ... contenido ...

# ============================================================
#  PANTALLA DE BIENVENIDA (cuando no hay ningún módulo abierto)
# ============================================================
else:
    st.markdown(f"""
    <div class="bienvenida-hero">
        <img src="{LOGO_BASE64}" class="logo-hero" alt="FV Procesados">
    </div>
    """, unsafe_allow_html=True)