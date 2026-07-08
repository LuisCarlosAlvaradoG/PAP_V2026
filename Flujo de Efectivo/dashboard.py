from datetime import date, timedelta

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from scipy.optimize import linprog, minimize

# ============================================================
#  CONFIGURACIÓN GENERAL
# ============================================================
ARCHIVO_DATOS = "Datos flujo de efectivo.xlsx"
ARCHIVO_SALIDA = "Flujo de efectivo.xlsx"
ARCHIVO_PAGOS = "Pagos pendientes.xlsx"
ARCHIVO_OPTIMIZACION = "Optimizacion de pagos.xlsx"

MESES = {
    1: "enero",
    2: "febrero",
    3: "marzo",
    4: "abril",
    5: "mayo",
    6: "junio",
    7: "julio",
    8: "agosto",
    9: "septiembre",
    10: "octubre",
    11: "noviembre",
    12: "diciembre",
}

# Logo oficial de FV Procesados, incrustado en base64 (fondo blanco removido)
# para que la app sea un solo archivo y no dependa de imágenes externas.
LOGO_BASE64 = "https://i.imgur.com/SNggzEh.png"

# ============================================================
#  SISTEMA DE DISEÑO — derivado del logo FV Procesados
# ============================================================
VERDE_OSCURO = "#3D5A28"
VERDE_MEDIO = "#5B7F3B"
VERDE_OLIVA = "#8A9A4E"
VERDE_CLARO = "#EEF3E3"
ROJO_FRESA = "#D6453D"
CREMA_FONDO = "#FBFAF9"
SIDEBAR_BG = "#24331A"
TINTA = "#22301A"


# ============================================================
#  FUNCIONES AUXILIARES
# ============================================================
def formato_semana(inicio, fin):
    if inicio.month == fin.month:
        return f"{inicio.day:02d}-{fin.day:02d} {MESES[fin.month]} {fin.year}"
    return f"{inicio.day:02d} {MESES[inicio.month]}-{fin.day:02d} {MESES[fin.month]} {fin.year}"


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


def moneda(valor):
    signo = "-" if valor < 0 else ""
    return f"{signo}${abs(valor):,.2f}"


# ------------------------------------------------------------
#  Formato compartido para las hojas de Excel generadas
# ------------------------------------------------------------
def aplicar_formato_tabla(ws, filas_resaltadas=None, fila_flujo_neto=None):
    """Aplica el mismo lenguaje visual (bordes, encabezado verde,
    números con separador de miles) a cualquier hoja generada por la app."""
    filas_resaltadas = filas_resaltadas or []

    verde_encabezado = PatternFill(
        fill_type="solid", fgColor=VERDE_OSCURO.replace("#", "")
    )
    verde_claro_fill = PatternFill(
        fill_type="solid", fgColor=VERDE_CLARO.replace("#", "")
    )
    verde_flujo_fill = PatternFill(
        fill_type="solid", fgColor=VERDE_MEDIO.replace("#", "")
    )

    borde = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    for row in ws.iter_rows():
        for cell in row:
            cell.border = borde
            cell.alignment = Alignment(horizontal="center", vertical="center")

    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = verde_encabezado

    for row in ws.iter_rows():
        if row[0].value in filas_resaltadas:
            for cell in row:
                cell.font = Font(bold=True, color=VERDE_OSCURO.replace("#", ""))
                cell.fill = verde_claro_fill
        if fila_flujo_neto and row[0].value == fila_flujo_neto:
            for cell in row:
                cell.font = Font(bold=True, size=13, color="FFFFFF")
                cell.fill = verde_flujo_fill

    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 2


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
    saldo_inicial = 500000
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
        aplicar_formato_tabla(
            ws,
            filas_resaltadas=["TOTAL INGRESOS", "TOTAL EGRESOS", "SALDO"],
            fila_flujo_neto="FLUJO NETO",
        )
        for row in ws.iter_rows(min_row=2, min_col=2):
            for cell in row:
                cell.number_format = "#,##0.00"

    return flujo


def actualizar_fin():
    st.session_state.fin_semana = st.session_state.inicio_semana + timedelta(days=5)


# ------------------------------------------------------------
#  Pagos pendientes: carga y guardado
# ------------------------------------------------------------
COLUMNAS_PAGOS = [
    "concepto",
    "categoría",
    "tipo",
    "importe",
    "semana",
    "inicio_semana",
    "fecha_registro",
    "estatus",
    "fecha_pago",
]


def cargar_pagos():
    try:
        dfp = pd.read_excel(ARCHIVO_PAGOS)
        for col in COLUMNAS_PAGOS:
            if col not in dfp.columns:
                dfp[col] = ""
        dfp = dfp[COLUMNAS_PAGOS]
    except FileNotFoundError:
        dfp = pd.DataFrame(columns=COLUMNAS_PAGOS)
        guardar_pagos(dfp)
    return dfp


def guardar_pagos(dfp):
    with pd.ExcelWriter(ARCHIVO_PAGOS, engine="openpyxl") as writer:
        dfp.to_excel(writer, sheet_name="Pagos", index=False)
        ws = writer.sheets["Pagos"]
        aplicar_formato_tabla(ws, filas_resaltadas=[])
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0.00"


# ============================================================
#  OPTIMIZACIÓN DE PAGOS
#  Implementación de los modelos descritos en el documento
#  "Modelo de Optimización F.E. FV-Procesados"
# ============================================================
def obtener_proyeccion_base(df, horizonte):
    """Proyecta ingresos y egresos operativos futuros usando el promedio
    de las últimas semanas registradas en el flujo de efectivo."""
    df2 = df.copy()
    df2["importe"] = pd.to_numeric(df2["importe"], errors="coerce").fillna(0)
    resumen = (
        df2.groupby("inicio_semana")
        .apply(
            lambda g: pd.Series(
                {
                    "ingreso": g.loc[g["tipo"] == "Ingreso", "importe"].sum(),
                    "egreso": g.loc[g["tipo"] == "Egreso", "importe"].sum(),
                }
            )
        )
        .sort_index()
    )
    ultimas = resumen.tail(6)
    ingreso_prom = float(ultimas["ingreso"].mean()) if not ultimas.empty else 0.0
    egreso_prom = float(ultimas["egreso"].mean()) if not ultimas.empty else 0.0
    return [ingreso_prom] * horizonte, [egreso_prom] * horizonte


def obtener_deudas_pendientes(dfp):
    """Agrupa la deuda pendiente por concepto/proveedor (D_i)."""
    pend = dfp[dfp["estatus"] == "Pendiente"].copy()
    pend["importe"] = pd.to_numeric(pend["importe"], errors="coerce").fillna(0)
    resumen = pend.groupby("concepto")["importe"].sum()
    return resumen[resumen > 0].sort_values(ascending=False)


def calcular_saldo_actual(df):
    df2 = df.copy()
    df2["importe"] = pd.to_numeric(df2["importe"], errors="coerce").fillna(0)
    ingresos = df2.loc[df2["tipo"] == "Ingreso", "importe"].sum()
    egresos = df2.loc[df2["tipo"] == "Egreso", "importe"].sum()
    return float(ingresos - egresos)


def optimizar_iter1(D0, pesos, I, E, C0, L_min, alpha, beta, T):
    """Eq. 1: min sum_t [ sum_i x_i D_i,t - alpha U_t + beta max(0, L_min-C_t) ]
    Se resuelve como programa lineal, linealizando el término de penalización
    con una variable de holgura s_t >= 0."""
    n = len(D0)
    IE = np.array(I, dtype=float) - np.array(E, dtype=float)
    cum_IE = np.cumsum(IE)

    nvars = n * T + T  # p[i,t] y s[t]
    c = np.zeros(nvars)
    for i in range(n):
        for tau in range(T):
            restante = T - tau  # semanas en las que este pago reduce D_i,t
            c[i * T + tau] = -pesos[i] * restante + alpha
    for t in range(T):
        c[n * T + t] = beta

    A_ub, b_ub = [], []

    # No se puede pagar más de la deuda pendiente
    for i in range(n):
        row = np.zeros(nvars)
        row[i * T : (i + 1) * T] = 1
        A_ub.append(row)
        b_ub.append(D0[i])

    # Caja no negativa
    for t in range(T):
        row = np.zeros(nvars)
        for i in range(n):
            row[i * T : i * T + t + 1] = 1
        A_ub.append(row)
        b_ub.append(C0 + cum_IE[t])

    # Holgura de liquidez: s_t >= L_min - C_t
    for t in range(T):
        row = np.zeros(nvars)
        for i in range(n):
            row[i * T : i * T + t + 1] = 1
        row[n * T + t] = -1
        A_ub.append(row)
        b_ub.append(C0 + cum_IE[t] - L_min)

    bounds = [(0, None)] * nvars
    res = linprog(
        c, A_ub=np.array(A_ub), b_ub=np.array(b_ub), bounds=bounds, method="highs"
    )
    P = res.x[: n * T].reshape(n, T) if res.success else np.zeros((n, T))
    return P, res.success


def optimizar_iter2(D0, I, E, C0, C_bar, alpha, beta, T):
    """Eq. 2: min alpha*sum (C_t - C_bar)^2 + beta*sum D_i (deuda final).
    Programa cuadrático convexo, resuelto con SLSQP."""
    n = len(D0)
    IE = np.array(I, dtype=float) - np.array(E, dtype=float)
    cum_IE = np.cumsum(IE)

    def caja(x):
        P = x.reshape(n, T)
        pagos_sem = P.sum(axis=0)
        return C0 + cum_IE - np.cumsum(pagos_sem)

    def objetivo(x):
        P = x.reshape(n, T)
        C = caja(x)
        deuda_final = np.array(D0) - P.sum(axis=1)
        valor = alpha * np.sum((C - C_bar) ** 2) + beta * np.sum(deuda_final)
        return valor / 1e6  # normaliza la escala para estabilidad numérica del solver

    def restr_caja(x):
        return caja(x)

    def restr_deuda(x):
        P = x.reshape(n, T)
        return np.array(D0) - P.sum(axis=1)

    x0 = np.zeros(n * T)
    bounds = [(0, None)] * (n * T)
    res = minimize(
        objetivo,
        x0,
        method="SLSQP",
        bounds=bounds,
        constraints=[
            {"type": "ineq", "fun": restr_caja},
            {"type": "ineq", "fun": restr_deuda},
        ],
        options={"maxiter": 300, "ftol": 1e-7},
    )
    P = res.x.reshape(n, T) if res.success else np.zeros((n, T))
    return P, res.success


def optimizar_max_resiliencia(D0, I, E, C0, T):
    """max min_t C_t, programa lineal vía epígrafe:
    max m  s.a.  C_t >= m  para toda t."""
    n = len(D0)
    IE = np.array(I, dtype=float) - np.array(E, dtype=float)
    cum_IE = np.cumsum(IE)

    nvars = n * T + 1  # p[i,t] y m
    c = np.zeros(nvars)
    c[-1] = -1  # minimizar -m equivale a maximizar m

    A_ub, b_ub = [], []
    for i in range(n):
        row = np.zeros(nvars)
        row[i * T : (i + 1) * T] = 1
        A_ub.append(row)
        b_ub.append(D0[i])

    for t in range(T):
        base = np.zeros(nvars)
        for i in range(n):
            base[i * T : i * T + t + 1] = 1
        A_ub.append(base.copy())
        b_ub.append(C0 + cum_IE[t])  # caja no negativa

        con_m = base.copy()
        con_m[-1] = 1
        A_ub.append(con_m)
        b_ub.append(C0 + cum_IE[t])  # C_t >= m

    bounds = [(0, None)] * (n * T) + [(None, None)]
    res = linprog(
        c, A_ub=np.array(A_ub), b_ub=np.array(b_ub), bounds=bounds, method="highs"
    )
    P = res.x[: n * T].reshape(n, T) if res.success else np.zeros((n, T))
    return P, res.success


def optimizar_max_crecimiento(D0, I, E, C0, alpha, T):
    """max sum U_t - alpha*Var(C_t), programa cuadrático resuelto con SLSQP."""
    n = len(D0)
    IE = np.array(I, dtype=float) - np.array(E, dtype=float)
    cum_IE = np.cumsum(IE)

    def caja(x):
        P = x.reshape(n, T)
        return C0 + cum_IE - np.cumsum(P.sum(axis=0))

    def objetivo(x):
        P = x.reshape(n, T)
        U = IE - P.sum(axis=0)
        C = caja(x)
        valor = -(np.sum(U) - alpha * np.var(C))
        return valor / 1e6  # normaliza la escala para estabilidad numérica del solver

    def restr_caja(x):
        return caja(x)

    def restr_deuda(x):
        P = x.reshape(n, T)
        return np.array(D0) - P.sum(axis=1)

    x0 = np.zeros(n * T)
    bounds = [(0, None)] * (n * T)
    res = minimize(
        objetivo,
        x0,
        method="SLSQP",
        bounds=bounds,
        constraints=[
            {"type": "ineq", "fun": restr_caja},
            {"type": "ineq", "fun": restr_deuda},
        ],
        options={"maxiter": 300, "ftol": 1e-7},
    )
    P = res.x.reshape(n, T) if res.success else np.zeros((n, T))
    return P, res.success


def optimizar_ratio_deuda_caja(D0, I, E, C0, L_min, T, agresividad):
    """min sum D_t/C_t. Es un ratio no convexo; se resuelve con una
    heurística secuencial que cada semana destina una fracción del
    excedente de caja (por encima de L_min) a las deudas más grandes."""
    n = len(D0)
    IE = np.array(I, dtype=float) - np.array(E, dtype=float)
    D_restante = np.array(D0, dtype=float)
    C = C0
    P = np.zeros((n, T))

    for t in range(T):
        C_disponible = C + IE[t]
        excedente = max(0.0, C_disponible - L_min)
        deuda_total = D_restante.sum()
        pago_total = min(excedente * agresividad, deuda_total)
        if pago_total > 0 and deuda_total > 0:
            proporciones = D_restante / deuda_total
            pagos_t = np.minimum(proporciones * pago_total, D_restante)
            P[:, t] = pagos_t
            D_restante = D_restante - pagos_t
        C = C_disponible - P[:, t].sum()

    return P, True


def construir_tabla_recomendacion(P, conceptos, I, E, C0, L_min):
    """Arma la tabla final de recomendaciones semana a semana, con el
    mismo criterio de semáforo (Operable / Ajustado / Sin margen)
    usado en el documento de referencia."""
    n, T = P.shape
    IE = np.array(I, dtype=float) - np.array(E, dtype=float)
    pagos_sem = P.sum(axis=0)
    C = C0 + np.cumsum(IE) - np.cumsum(pagos_sem)
    U = IE - pagos_sem

    filas = []
    for t in range(T):
        fila = {"Semana": t + 1}
        for i, concepto in enumerate(conceptos):
            fila[f"Pagar {concepto} ($)"] = round(float(P[i, t]), 2)
        fila["Utilidad/Retenida ($)"] = round(float(U[t]), 2)
        fila["Caja Proyectada ($)"] = round(float(C[t]), 2)
        if C[t] <= L_min:
            estado = "Sin margen"
        elif C[t] <= L_min * 1.5:
            estado = "Ajustado"
        else:
            estado = "Operable"
        fila["Estado"] = estado
        filas.append(fila)

    return pd.DataFrame(filas)


# ============================================================
#  PÁGINA + SISTEMA DE DISEÑO
# ============================================================
st.set_page_config(
    page_title="FV Procesados | Sistema Financiero",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    f"""
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
""",
    unsafe_allow_html=True,
)

# ============================================================
#  ESTADO DE NAVEGACIÓN
# ============================================================
if "modulo_activo" not in st.session_state:
    st.session_state.modulo_activo = None
if "seccion_flujo" not in st.session_state:
    st.session_state.seccion_flujo = None
if "seccion_pagos" not in st.session_state:
    st.session_state.seccion_pagos = None

if "inicio_semana" not in st.session_state:
    st.session_state.inicio_semana = date.today()
if "fin_semana" not in st.session_state:
    st.session_state.fin_semana = date.today() + timedelta(days=5)
if "concepto_seleccionado" not in st.session_state:
    st.session_state.concepto_seleccionado = None
if "categoria_seleccionada" not in st.session_state:
    st.session_state.categoria_seleccionada = None

# Estado para el sub-formulario "Agregar deuda"
if "deuda_inicio_semana" not in st.session_state:
    st.session_state.deuda_inicio_semana = date.today()
if "deuda_fin_semana" not in st.session_state:
    st.session_state.deuda_fin_semana = date.today() + timedelta(days=5)


def ir_a(modulo, seccion=None):
    st.session_state.modulo_activo = modulo
    if modulo == "flujo":
        st.session_state.seccion_flujo = seccion
    elif modulo == "pagos":
        st.session_state.seccion_pagos = seccion
    st.rerun()


def actualizar_fin_deuda():
    st.session_state.deuda_fin_semana = (
        st.session_state.deuda_inicio_semana + timedelta(days=5)
    )


# ============================================================
#  SIDEBAR — NAVEGACIÓN
# ============================================================
with st.sidebar:
    st.markdown(
        f"""
    <div class="marca-wordmark">
        <img src="{LOGO_BASE64}" class="logo-sidebar" alt="FV Procesados">
        <div class="regla-marca"></div>
    </div>
    """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="nav-eyebrow">Módulos</div>', unsafe_allow_html=True)

    # ---- Módulo: Flujo de Efectivo ----
    if st.session_state.modulo_activo == "flujo":
        st.markdown(
            '<div class="nav-item-activo">Flujo de Efectivo</div>',
            unsafe_allow_html=True,
        )

        st.markdown('<div class="nav-sub">', unsafe_allow_html=True)
        secciones = [
            ("ver", "Ver flujo de efectivo"),
            ("agregar", "Agregar movimiento"),
            ("eliminar", "Eliminar movimiento"),
        ]
        for clave, etiqueta in secciones:
            if st.session_state.seccion_flujo == clave:
                st.markdown(
                    f'<div class="nav-sub-activo">{etiqueta}</div>',
                    unsafe_allow_html=True,
                )
            else:
                if st.button(etiqueta, key=f"nav_{clave}", use_container_width=True):
                    ir_a("flujo", clave)
        st.markdown("</div>", unsafe_allow_html=True)
    else:
        if st.button(
            "Flujo de Efectivo", key="nav_modulo_flujo", use_container_width=True
        ):
            ir_a("flujo", "ver")

    st.markdown('<div class="nav-item-inactivo-sep"></div>', unsafe_allow_html=True)

    # ---- Módulo: Pagos Pendientes ----
    if st.session_state.modulo_activo == "pagos":
        st.markdown(
            '<div class="nav-item-activo">Pagos Pendientes</div>',
            unsafe_allow_html=True,
        )

        st.markdown('<div class="nav-sub">', unsafe_allow_html=True)
        secciones_pagos = [
            ("agregar", "Agregar deuda"),
            ("pagar", "Pagar un pago pendiente"),
            ("optimizar", "Optimización de pagos"),
        ]
        for clave, etiqueta in secciones_pagos:
            if st.session_state.seccion_pagos == clave:
                st.markdown(
                    f'<div class="nav-sub-activo">{etiqueta}</div>',
                    unsafe_allow_html=True,
                )
            else:
                if st.button(
                    etiqueta, key=f"nav_pagos_{clave}", use_container_width=True
                ):
                    ir_a("pagos", clave)
        st.markdown("</div>", unsafe_allow_html=True)
    else:
        if st.button(
            "Pagos Pendientes", key="nav_modulo_pagos", use_container_width=True
        ):
            ir_a("pagos", "agregar")

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

TITULOS_SECCION_PAGOS = {
    "agregar": "Agregar deuda",
    "pagar": "Pagar un pago pendiente",
    "optimizar": "Optimización de pagos",
}

# ============================================================
#  MÓDULO 1: FLUJO DE EFECTIVO
# ============================================================
if st.session_state.modulo_activo == "flujo":
    seccion = st.session_state.seccion_flujo

    st.markdown(
        f"""
    <div class="encabezado-modulo">
        <div class="eyebrow">Flujo de Efectivo</div>
        <h1>{TITULOS_SECCION[seccion]}</h1>
    </div>
    <div class="regla-marca"></div>
    """,
        unsafe_allow_html=True,
    )

    # ---------------- VER FLUJO DE EFECTIVO ----------------
    if seccion == "ver":
        df_actual = pd.read_excel(ARCHIVO_DATOS)
        df_actual["tipo"] = df_actual["tipo"].astype(str).str.strip()
        df_actual["importe"] = pd.to_numeric(
            df_actual["importe"], errors="coerce"
        ).fillna(0)

        flujo = generar_flujo(df_actual)

        total_ingresos = df_actual.loc[df_actual["tipo"] == "Ingreso", "importe"].sum()
        total_egresos = df_actual.loc[df_actual["tipo"] == "Egreso", "importe"].sum()
        flujo_neto = total_ingresos - total_egresos
        saldo_actual = flujo.loc["SALDO"].iloc[-1] if flujo.shape[1] > 0 else 0

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(
                f"""
            <div class="kpi-tarjeta">
                <div class="kpi-label">Ingresos totales</div>
                <div class="kpi-valor">{moneda(total_ingresos)}</div>
            </div>""",
                unsafe_allow_html=True,
            )
        with c2:
            st.markdown(
                f"""
            <div class="kpi-tarjeta kpi-egreso">
                <div class="kpi-label">Egresos totales</div>
                <div class="kpi-valor">{moneda(total_egresos)}</div>
            </div>""",
                unsafe_allow_html=True,
            )
        with c3:
            st.markdown(
                f"""
            <div class="kpi-tarjeta">
                <div class="kpi-label">Flujo neto</div>
                <div class="kpi-valor">{moneda(flujo_neto)}</div>
            </div>""",
                unsafe_allow_html=True,
            )
        with c4:
            st.markdown(
                f"""
            <div class="kpi-tarjeta kpi-saldo">
                <div class="kpi-label">Saldo actual</div>
                <div class="kpi-valor">{moneda(saldo_actual)}</div>
            </div>""",
                unsafe_allow_html=True,
            )

        st.write("")
        with st.container(border=True):
            st.dataframe(flujo, use_container_width=True)

        with open(ARCHIVO_SALIDA, "rb") as archivo:
            st.download_button(
                label="Descargar Excel",
                data=archivo,
                file_name=ARCHIVO_SALIDA,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # ---------------- AGREGAR MOVIMIENTO ----------------
    elif seccion == "agregar":
        with st.container(border=True):
            col1, col2 = st.columns(2)
            with col1:
                inicio = st.date_input(
                    "Inicio de semana", key="inicio_semana", on_change=actualizar_fin
                )
            with col2:
                fin = st.date_input("Fin de semana", key="fin_semana")

            semana = formato_semana(inicio, fin)
            st.info(f"Semana seleccionada: {semana}")

            importe = st.number_input("Importe", min_value=0.0, step=100.0)

            st.markdown('<div class="regla-marca"></div>', unsafe_allow_html=True)

            conceptos_existentes = obtener_conceptos(df)
            categorias_existentes = obtener_categorias(df)

            modo_concepto = st.radio(
                "Concepto",
                ["Usar concepto existente", "Agregar concepto nuevo"],
                horizontal=True,
            )

            concepto = ""
            categoria = ""
            tipo = ""

            if modo_concepto == "Usar concepto existente":
                concepto = st.selectbox(
                    "Selecciona el concepto", [""] + conceptos_existentes
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
                    "Nuevo concepto", placeholder="Escribe el concepto"
                ).strip()

                if st.session_state.concepto_seleccionado is not None:
                    concepto = st.session_state.concepto_seleccionado
                    datos_concepto = df[
                        df["descripción"].astype(str).str.strip() == concepto
                    ].iloc[0]
                    categoria = datos_concepto["categoría"]
                    tipo = datos_concepto["tipo"]
                    st.text_input(
                        "Concepto seleccionado", value=concepto, disabled=True
                    )
                    st.text_input("Categoría", value=categoria, disabled=True)
                    st.text_input("Tipo", value=tipo, disabled=True)
                else:
                    concepto = formato_titulo(concepto_escrito)
                    coincidencias_concepto = buscar_similares(
                        concepto_escrito, conceptos_existentes
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
                        horizontal=True,
                    )

                    if modo_categoria == "Usar categoría existente":
                        categoria = st.selectbox(
                            "Selecciona la categoría", [""] + categorias_existentes
                        )
                        if categoria != "":
                            tipo = obtener_tipo_por_categoria(categoria)
                            st.text_input("Tipo", value=tipo, disabled=True)
                    else:
                        categoria_escrita = st.text_input(
                            "Nueva categoría", placeholder="Escribe la categoría"
                        ).strip()

                        if st.session_state.categoria_seleccionada is not None:
                            categoria = st.session_state.categoria_seleccionada
                            tipo = obtener_tipo_por_categoria(categoria)
                            st.text_input(
                                "Categoría seleccionada", value=categoria, disabled=True
                            )
                            st.text_input("Tipo", value=tipo, disabled=True)
                        else:
                            categoria = formato_titulo(categoria_escrita)
                            coincidencias_categoria = buscar_similares(
                                categoria_escrita, categorias_existentes
                            )
                            if categoria_escrita and coincidencias_categoria:
                                st.caption("Coincidencias encontradas:")
                                columnas_cat = st.columns(4)
                                for i, sugerencia in enumerate(coincidencias_categoria):
                                    with columnas_cat[i % 4]:
                                        if st.button(sugerencia, key=f"categoria_{i}"):
                                            st.session_state.categoria_seleccionada = (
                                                sugerencia
                                            )
                                            st.rerun()
                            tipo = st.selectbox(
                                "Tipo", ["Ingreso", "Egreso"], index=1
                            )

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
                    nuevo = pd.DataFrame(
                        [
                            {
                                "inicio_semana": pd.to_datetime(inicio),
                                "semana": semana,
                                "tipo": tipo,
                                "categoría": categoria,
                                "descripción": concepto,
                                "importe": importe,
                            }
                        ]
                    )
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
                df_eliminar["inicio_semana"], dayfirst=True
            )
            df_eliminar["importe"] = pd.to_numeric(
                df_eliminar["importe"], errors="coerce"
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
                [""] + df_eliminar["movimiento"].tolist(),
            )
            confirmar = st.checkbox("Confirmo que quiero eliminar este movimiento")

            if st.button("Eliminar movimiento"):
                if movimiento == "":
                    st.error("Debes seleccionar un movimiento.")
                elif not confirmar:
                    st.error("Debes confirmar antes de eliminar.")
                else:
                    indice = df_eliminar[df_eliminar["movimiento"] == movimiento].index[
                        0
                    ]
                    df_eliminar = df_eliminar.drop(index=indice)
                    df_eliminar = df_eliminar.drop(columns=["movimiento"])
                    df_eliminar.to_excel(ARCHIVO_DATOS, index=False)
                    generar_flujo(df_eliminar)
                    st.success(
                        "Movimiento eliminado correctamente y flujo actualizado."
                    )


# ============================================================
#  MÓDULO 2: PAGOS PENDIENTES
# ============================================================
elif st.session_state.modulo_activo == "pagos":
    seccion_p = st.session_state.seccion_pagos

    st.markdown(
        f"""
    <div class="encabezado-modulo">
        <div class="eyebrow">Pagos Pendientes</div>
        <h1>{TITULOS_SECCION_PAGOS[seccion_p]}</h1>
    </div>
    <div class="regla-marca"></div>
    """,
        unsafe_allow_html=True,
    )

    dfp = cargar_pagos()

    # ---------------- AGREGAR DEUDA ----------------
    if seccion_p == "agregar":
        with st.container(border=True):
            conceptos_existentes = obtener_conceptos(df)

            concepto_deuda = st.selectbox(
                "Selecciona el concepto ya registrado (proveedor / concepto)",
                [""] + conceptos_existentes,
            )

            categoria_deuda = ""
            tipo_deuda = ""
            if concepto_deuda != "":
                datos_concepto = df[
                    df["descripción"].astype(str).str.strip() == concepto_deuda
                ].iloc[0]
                categoria_deuda = datos_concepto["categoría"]
                tipo_deuda = datos_concepto["tipo"]
                st.text_input("Categoría", value=categoria_deuda, disabled=True)
                st.text_input("Tipo", value=tipo_deuda, disabled=True)

            col1, col2 = st.columns(2)
            with col1:
                inicio_deuda = st.date_input(
                    "Inicio de semana a la que corresponde la deuda",
                    key="deuda_inicio_semana",
                    on_change=actualizar_fin_deuda,
                )
            with col2:
                fin_deuda = st.date_input(
                    "Fin de semana a la que corresponde la deuda",
                    key="deuda_fin_semana",
                )

            semana_deuda = formato_semana(inicio_deuda, fin_deuda)
            st.info(f"Semana de la deuda: {semana_deuda}")

            importe_deuda = st.number_input(
                "Importe pendiente", min_value=0.0, step=100.0, key="importe_deuda"
            )

            if st.button("Registrar deuda"):
                if concepto_deuda == "":
                    st.error("Debes seleccionar un concepto.")
                elif importe_deuda <= 0:
                    st.error("El importe debe ser mayor a 0.")
                elif fin_deuda < inicio_deuda:
                    st.error("La fecha final no puede ser anterior a la fecha inicial.")
                else:
                    nueva_deuda = pd.DataFrame(
                        [
                            {
                                "concepto": concepto_deuda,
                                "categoría": categoria_deuda,
                                "tipo": tipo_deuda,
                                "importe": importe_deuda,
                                "semana": semana_deuda,
                                "inicio_semana": pd.to_datetime(inicio_deuda),
                                "fecha_registro": pd.to_datetime(date.today()),
                                "estatus": "Pendiente",
                                "fecha_pago": "",
                            }
                        ]
                    )
                    dfp_actualizado = pd.concat([dfp, nueva_deuda], ignore_index=True)
                    guardar_pagos(dfp_actualizado)
                    st.success("Deuda registrada correctamente como pago pendiente.")

        st.write("")
        pendientes_actuales = dfp[dfp["estatus"] == "Pendiente"]
        if not pendientes_actuales.empty:
            with st.container(border=True):
                st.caption("Pagos pendientes actuales (por semana)")
                st.dataframe(
                    pendientes_actuales[
                        [
                            "concepto",
                            "categoría",
                            "tipo",
                            "importe",
                            "semana",
                            "fecha_registro",
                        ]
                    ],
                    use_container_width=True,
                )

    # ---------------- PAGAR UN PAGO PENDIENTE ----------------
    elif seccion_p == "pagar":
        with st.container(border=True):
            pendientes = dfp[dfp["estatus"] == "Pendiente"].copy()
            pendientes["importe"] = pd.to_numeric(
                pendientes["importe"], errors="coerce"
            ).fillna(0)

            if pendientes.empty:
                st.info("No hay pagos pendientes registrados.")
            else:
                # Se agrupa por concepto: el proveedor puede tener deuda
                # acumulada de varias semanas distintas.
                totales_por_concepto = (
                    pendientes.groupby("concepto")["importe"].sum().sort_index()
                )
                opciones_concepto = totales_por_concepto.index.tolist()

                concepto_pago = st.selectbox(
                    "Selecciona el concepto a pagar", [""] + opciones_concepto
                )

                total_pendiente = 0.0
                if concepto_pago != "":
                    total_pendiente = float(totales_por_concepto[concepto_pago])
                    st.info(
                        f"Deuda pendiente total de {concepto_pago}: {moneda(total_pendiente)}"
                    )

                    detalle = pendientes[pendientes["concepto"] == concepto_pago][
                        ["semana", "importe", "fecha_registro"]
                    ].sort_values("fecha_registro")
                    st.caption("Detalle por semana:")
                    st.dataframe(detalle, use_container_width=True, hide_index=True)

                monto_pagar = st.number_input(
                    "Monto que se va a pagar",
                    min_value=0.0,
                    max_value=total_pendiente if total_pendiente > 0 else 0.0,
                    step=100.0,
                    key="monto_pagar",
                )

                confirmar_pago = st.checkbox("Confirmo que este monto fue liquidado")

                if st.button("Registrar pago"):
                    if concepto_pago == "":
                        st.error("Debes seleccionar un concepto.")
                    elif monto_pagar <= 0:
                        st.error("El monto a pagar debe ser mayor a 0.")
                    elif monto_pagar > total_pendiente:
                        st.error(
                            "El monto a pagar no puede ser mayor a la deuda pendiente."
                        )
                    elif not confirmar_pago:
                        st.error("Debes confirmar antes de registrar el pago.")
                    else:
                        # Se descuenta el monto de las deudas más antiguas
                        # primero, hasta agotar el pago realizado.
                        indices_concepto = (
                            pendientes[pendientes["concepto"] == concepto_pago]
                            .sort_values("fecha_registro")
                            .index.tolist()
                        )

                        datos_concepto_pago = pendientes[
                            pendientes["concepto"] == concepto_pago
                        ].iloc[0]

                        monto_restante = monto_pagar
                        for idx in indices_concepto:
                            if monto_restante <= 0:
                                break
                            saldo_fila = float(dfp.loc[idx, "importe"])
                            abono = min(saldo_fila, monto_restante)
                            nuevo_saldo = round(saldo_fila - abono, 2)
                            dfp.loc[idx, "importe"] = nuevo_saldo
                            monto_restante = round(monto_restante - abono, 2)
                            if nuevo_saldo <= 0:
                                dfp.loc[idx, "importe"] = 0
                                dfp.loc[idx, "estatus"] = "Pagado"
                                dfp.loc[idx, "fecha_pago"] = pd.to_datetime(
                                    date.today()
                                )

                        guardar_pagos(dfp)

                        # Se refleja el pago como egreso real en el flujo de
                        # efectivo, usando la semana actual (la del pago).
                        inicio_pago = date.today()
                        fin_pago = inicio_pago + timedelta(days=5)
                        semana_pago = formato_semana(inicio_pago, fin_pago)

                        nuevo_movimiento = pd.DataFrame(
                            [
                                {
                                    "inicio_semana": pd.to_datetime(inicio_pago),
                                    "semana": semana_pago,
                                    "tipo": datos_concepto_pago["tipo"],
                                    "categoría": datos_concepto_pago["categoría"],
                                    "descripción": concepto_pago,
                                    "importe": monto_pagar,
                                }
                            ]
                        )
                        df_actualizado = pd.concat(
                            [df, nuevo_movimiento], ignore_index=True
                        )
                        df_actualizado.to_excel(ARCHIVO_DATOS, index=False)
                        generar_flujo(df_actualizado)

                        restante = round(total_pendiente - monto_pagar, 2)
                        if restante <= 0:
                            st.success(
                                f"Pago registrado. La deuda de {concepto_pago} quedó liquidada por completo."
                            )
                        else:
                            st.success(
                                f"Pago registrado. Queda un saldo pendiente de {moneda(restante)} en {concepto_pago}."
                            )

        st.write("")
        pagados = dfp[dfp["estatus"] == "Pagado"]
        if not pagados.empty:
            with st.container(border=True):
                st.caption("Historial de pagos liquidados")
                st.dataframe(
                    pagados[
                        [
                            "concepto",
                            "categoría",
                            "importe",
                            "semana",
                            "fecha_registro",
                            "fecha_pago",
                        ]
                    ],
                    use_container_width=True,
                )

    # ---------------- OPTIMIZACIÓN DE PAGOS ----------------
    elif seccion_p == "optimizar":
        deudas = obtener_deudas_pendientes(dfp)

        if deudas.empty:
            st.info(
                "No hay pagos pendientes registrados para optimizar. "
                "Registra deudas en 'Agregar deuda' primero."
            )
        else:
            with st.container(border=True):
                st.caption("Deudas pendientes consideradas (D_i), sumadas por concepto")
                st.dataframe(
                    deudas.rename("Importe pendiente ($)"), use_container_width=True
                )

                saldo_actual = calcular_saldo_actual(df)

                col1, col2 = st.columns(2)
                with col1:
                    horizonte = st.number_input(
                        "Horizonte de planeación (semanas, T)",
                        min_value=2,
                        max_value=26,
                        value=8,
                        step=1,
                    )
                with col2:
                    L_min = st.number_input(
                        "Caja mínima deseada (L_min)",
                        min_value=0.0,
                        value=round(max(saldo_actual * 0.1, 0.0), 2),
                        step=500.0,
                    )

                st.markdown('<div class="regla-marca"></div>', unsafe_allow_html=True)

                modelo = st.selectbox(
                    "Modelo de optimización",
                    [
                        "Prioridad en liquidación de deuda",
                        "Estabilidad de caja",
                        "Máxima resiliencia financiera (max-min caja)",
                        "Máximo crecimiento sostenible",
                        "Mínimo ratio deuda/caja",
                    ],
                )

                alpha = beta = C_bar = agresividad = None
                if modelo.startswith("Prioridad en liquidación"):
                    st.caption("min Σ [ Σ x_i·D_i,t − α·U_t + β·max(0, L_min−C_t) ]")
                    alpha = st.slider(
                        "α — peso a la utilidad retenida", 0.0, 1.0, 0.3, 0.05
                    )
                    beta = st.slider(
                        "β — penalización por caja debajo de L_min", 0.0, 5.0, 1.0, 0.1
                    )
                elif modelo.startswith("Estabilidad"):
                    st.caption("min α·Σ(C_t − C̄)² + β·Σ D_i")
                    C_bar = st.number_input(
                        "Caja objetivo (C̄)",
                        min_value=0.0,
                        value=round(max(saldo_actual, L_min * 1.5), 2),
                        step=500.0,
                    )
                    alpha = st.slider(
                        "α — peso a la estabilidad de caja", 0.01, 5.0, 1.0, 0.1
                    )
                    beta = st.slider("β — peso a liquidar deuda", 0.0, 5.0, 0.5, 0.1)
                elif modelo.startswith("Máxima resiliencia"):
                    st.caption("max mín_t C_t")
                    st.caption(
                        "Nota: este modelo solo protege el piso de caja; si la caja ya es "
                        "holgada frente a L_min, puede recomendar no pagar nada."
                    )
                elif modelo.startswith("Máximo crecimiento"):
                    st.caption("max Σ U_t − α·Var(C_t)")
                    alpha = st.slider(
                        "α — penalización a la varianza de caja", 0.0, 5.0, 1.0, 0.1
                    )
                else:
                    st.caption(
                        "min Σ D_t / C_t  (heurística: reparte el excedente semanal entre las deudas)"
                    )
                    agresividad = st.slider(
                        "Agresividad del pago semanal", 0.0, 1.0, 0.5, 0.05
                    )

                calcular = st.button("Calcular recomendaciones")

            if calcular:
                I, E = obtener_proyeccion_base(df, horizonte)
                C0 = calcular_saldo_actual(df)
                conceptos = deudas.index.tolist()
                D0 = deudas.values.tolist()
                pesos = (deudas / deudas.sum()).values.tolist()

                if modelo.startswith("Prioridad en liquidación"):
                    P, exito = optimizar_iter1(
                        D0, pesos, I, E, C0, L_min, alpha, beta, horizonte
                    )
                elif modelo.startswith("Estabilidad"):
                    P, exito = optimizar_iter2(
                        D0, I, E, C0, C_bar, alpha, beta, horizonte
                    )
                elif modelo.startswith("Máxima resiliencia"):
                    P, exito = optimizar_max_resiliencia(D0, I, E, C0, horizonte)
                elif modelo.startswith("Máximo crecimiento"):
                    P, exito = optimizar_max_crecimiento(D0, I, E, C0, alpha, horizonte)
                else:
                    P, exito = optimizar_ratio_deuda_caja(
                        D0, I, E, C0, L_min, horizonte, agresividad
                    )

                if not exito:
                    st.error(
                        "El modelo no encontró una solución factible con estos parámetros. "
                        "Intenta reducir L_min, ampliar el horizonte, o revisar los ingresos/egresos proyectados."
                    )
                    st.session_state.pop("tabla_optimizacion", None)
                else:
                    st.session_state["tabla_optimizacion"] = (
                        construir_tabla_recomendacion(P, conceptos, I, E, C0, L_min)
                    )
                    st.session_state["modelo_optimizacion"] = modelo
                    st.session_state["proyeccion_optimizacion"] = (
                        f"Proyección basada en el promedio de las últimas semanas registradas: "
                        f"ingreso ≈ {moneda(I[0])}/semana, egreso operativo ≈ {moneda(E[0])}/semana, "
                        f"caja inicial {moneda(C0)}."
                    )

            if "tabla_optimizacion" in st.session_state:
                st.write("")
                st.caption(
                    f"Recomendación generada — {st.session_state['modelo_optimizacion']}"
                )
                st.caption(st.session_state["proyeccion_optimizacion"])
                with st.container(border=True):
                    st.dataframe(
                        st.session_state["tabla_optimizacion"],
                        use_container_width=True,
                        hide_index=True,
                    )

                with pd.ExcelWriter(ARCHIVO_OPTIMIZACION, engine="openpyxl") as writer:
                    st.session_state["tabla_optimizacion"].to_excel(
                        writer, sheet_name="Optimización", index=False
                    )
                    ws = writer.sheets["Optimización"]
                    aplicar_formato_tabla(ws)
                    for row in ws.iter_rows(min_row=2):
                        for cell in row:
                            if isinstance(cell.value, (int, float)):
                                cell.number_format = "#,##0.00"

                with open(ARCHIVO_OPTIMIZACION, "rb") as archivo:
                    st.download_button(
                        "Descargar recomendaciones en Excel",
                        data=archivo,
                        file_name=ARCHIVO_OPTIMIZACION,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="descarga_optimizacion",
                    )

# ============================================================
#  PANTALLA DE BIENVENIDA (cuando no hay ningún módulo abierto)
# ============================================================
else:
    st.markdown(
        f"""
    <div class="bienvenida-hero">
        <img src="{LOGO_BASE64}" class="logo-hero" alt="FV Procesados">
    </div>
    """,
        unsafe_allow_html=True,
    )
