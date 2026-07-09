from datetime import date, timedelta

import cvxpy as cp
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from scipy.optimize import linprog

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
def filtrar_semanas_representativas(df, umbral_min_transacciones=None):
    """Excluye del histórico, SOLO para efectos de proyección, cualquier semana que:
        (a) todavía esté en curso (inicio_semana + 6 días >= hoy), o
        (b) tenga muy pocas transacciones comparada con una semana normal --
            es decir, no sea representativa de una semana típica de operación.

    CORRECCIÓN (caso real): al probar el sistema, se inyectó un ingreso de
    prueba de $500,000 como único movimiento de una semana. Esa semana ya
    había "terminado" en el calendario (pasaba el filtro (a)), pero seguía sin
    ser una semana real de operación -- un solo movimiento no son las 15-30
    transacciones típicas de una semana normal. Si cae dentro de la ventana de
    8 semanas que usa la media móvil, el filtro (a) por sí solo no la detecta,
    e infla artificialmente la proyección futura. El filtro (b) generaliza la
    corrección para que esto no dependa de en qué semana caiga el movimiento.

    NOTA IMPORTANTE: este filtro solo debe aplicarse a la PROYECCIÓN futura
    (`obtener_proyeccion`), nunca al saldo actual (`calcular_saldo_actual`) --
    el saldo actual debe reflejar TODO lo registrado hasta hoy, representativo
    o no (si de verdad entró dinero, ya es saldo real, sin importar si esa
    semana "se ve típica" o no).
    """
    df = df.copy()
    hoy = pd.Timestamp.today().normalize()
    fin_de_semana = df["inicio_semana"] + pd.Timedelta(days=6)
    df = df.loc[fin_de_semana < hoy].copy()

    if df.empty:
        return df

    conteo = df.groupby("inicio_semana").size()
    mediana = conteo.median()
    if umbral_min_transacciones is None:
        umbral_min_transacciones = max(5, mediana * 0.4)
    semanas_validas = conteo[conteo >= umbral_min_transacciones].index
    return df[df["inicio_semana"].isin(semanas_validas)].copy()


def obtener_proyeccion(df, horizonte):
    """
    Proyección adaptativa de ingresos y egresos semanales.

    Estrategia automática según la cantidad de semanas históricas:
        < 8 semanas       : promedio simple de todo el historial.
        8 a 51 semanas    : media móvil ponderada de las últimas 8 semanas.
        >= 52 semanas     : media móvil ponderada + factor estacional por
                            semana del año (ISO week).

    Usa `filtrar_semanas_representativas` antes de construir el histórico
    (ver esa función para el detalle de por qué es necesario).

    Siempre devuelve dos listas de longitud `horizonte`:
        I : ingresos proyectados por semana
        E : egresos proyectados por semana
    """
    df = filtrar_semanas_representativas(df)
    df["importe"] = pd.to_numeric(df["importe"], errors="coerce").fillna(0)

    resumen = df.pivot_table(
        index="inicio_semana",
        columns="tipo",
        values="importe",
        aggfunc="sum",
        fill_value=0,
    ).sort_index()

    for col in ["Ingreso", "Egreso"]:
        if col not in resumen.columns:
            resumen[col] = 0.0

    ingreso_hist = resumen["Ingreso"].values.astype(float)
    egreso_hist = resumen["Egreso"].values.astype(float)
    n_semanas = len(ingreso_hist)

    pesos = np.array([1, 2, 3, 4, 5, 6, 7, 8], dtype=float)
    pesos = pesos / pesos.sum()

    def wma(serie):
        ultimas = serie[-8:]
        n = len(ultimas)
        pesos_ajustados = pesos[-n:]
        pesos_ajustados = pesos_ajustados / pesos_ajustados.sum()
        return np.dot(ultimas, pesos_ajustados)

    factores_ingreso = None
    factores_egreso = None
    if n_semanas >= 52:
        semanas_iso = resumen.index.isocalendar().week.values

        def calcular_factores(serie):
            df_temp = pd.DataFrame({"semana": semanas_iso, "valor": serie})
            promedio_semana = df_temp.groupby("semana")["valor"].mean()
            global_avg = serie.mean()
            if global_avg == 0:
                return {s: 1.0 for s in range(1, 54)}
            return (promedio_semana / global_avg).to_dict()

        factores_ingreso = calcular_factores(ingreso_hist)
        factores_egreso = calcular_factores(egreso_hist)

    if n_semanas == 0:
        return [0.0] * horizonte, [0.0] * horizonte

    ultima_fecha = resumen.index[-1]
    fechas_futuras = pd.date_range(
        start=ultima_fecha + pd.Timedelta(weeks=1), periods=horizonte, freq="W"
    )
    semanas_futuras = fechas_futuras.isocalendar().week

    def proyectar(serie_hist, factores):
        if n_semanas < 8:
            base = serie_hist.mean()
            return np.full(horizonte, base)
        base = wma(serie_hist)
        if n_semanas < 52 or factores is None:
            return np.full(horizonte, base)
        proy = np.array([base * factores.get(s, 1.0) for s in semanas_futuras])
        return proy

    I = proyectar(ingreso_hist, factores_ingreso).tolist()
    E = proyectar(egreso_hist, factores_egreso).tolist()
    return I, E


def obtener_deudas_pendientes(dfp):
    """Agrupa la deuda pendiente por concepto/proveedor (D_i).

    Ordenado de MENOR a MAYOR a propósito: como el snowball paga las deudas
    más chicas primero, este orden deja la actividad de pago real visible
    desde el principio de las tablas de recomendación, en vez de escondida
    hasta el final.
    """
    pend = dfp[dfp["estatus"] == "Pendiente"].copy()
    pend["importe"] = pd.to_numeric(pend["importe"], errors="coerce").fillna(0)
    resumen = pend.groupby("concepto")["importe"].sum()
    return resumen[resumen > 0].sort_values(ascending=True)


def calcular_saldo_actual(df):
    """Saldo actual (C0) = ingresos - egresos de TODO el histórico registrado,
    sin filtrar nada -- a diferencia de la proyección, aquí sí debe contar
    cualquier movimiento ya registrado (incluida la semana en curso, o una
    inyección de capital de una sola vez), porque ya es saldo real, sin
    importar qué tan "representativo" sea de una semana típica.
    """
    df2 = df.copy()
    df2["importe"] = pd.to_numeric(df2["importe"], errors="coerce").fillna(0)
    ingresos = df2.loc[df2["tipo"] == "Ingreso", "importe"].sum()
    egresos = df2.loc[df2["tipo"] == "Egreso", "importe"].sum()
    return float(ingresos - egresos)


def modelo_a_optimizacion_global(
    D0, I, E, C0, L_min, L_survival, T, pago_max_frac_ingreso=0.5, eps=1e-6
):
    """
    Modelo A - Optimización global de caja (dos etapas), DECISIÓN LIBRE.

    HISTORIAL DEL DISEÑO: en versiones anteriores, este modelo recibía un
    presupuesto de pago FIJO (p. ej. "paga el 100% de la deuda"), porque con
    una caja estructuralmente negativa, cualquier piso positivo (L_survival)
    era matemáticamente imposible de alcanzar como restricción dura -- el
    problema se volvía infactible sin importar el pago, así que había que
    decírselo de antemano.

    Con una caja inicial ya positiva (por ejemplo, tras una inyección de
    capital), L_survival SÍ es alcanzable como restricción dura real. Esto
    permite regresar al diseño original: el modelo YA NO recibe un presupuesto
    fijo -- decide por sí mismo cuánta deuda pagar, maximizando el total
    pagado sin que la caja baje nunca del piso de emergencia (L_survival).
    L_min (el colchón cómodo) ya no es obligatorio, pero se usa como
    desempate suave para escoger, entre varios calendarios de pago igual de
    buenos, el que más proteja ese colchón.

    Etapa A (QP) - decide CUÁNTO pagar en total cada semana, maximizando la
    deuda total pagada sujeta a C[t] >= L_survival (restricción dura) y a un
    tope semanal realista.

    Etapa B (snowball estricto): reparte el calendario de pagos ya decidido
    entre los conceptos de deuda, de la más chica a la más grande, saldando
    cada una por completo antes de tocar la siguiente.

    Si L_survival no es alcanzable (p. ej. la caja sigue siendo negativa),
    el problema será infactible y se reporta `exito=False` -- en ese caso, la
    caja inicial no alcanza ni para sobrevivir sin pagar nada, y hace falta
    revisar los datos o el saldo inicial antes de pedirle al modelo pagar deuda.
    """
    D0 = np.asarray(D0, dtype=float)
    n = len(D0)
    D_total0 = D0.sum()
    I_arr = np.asarray(I, dtype=float)
    tope_semanal = pago_max_frac_ingreso * I_arr

    alpha_tiebreak = (
        1.0 / (L_min**2) if L_min > 0 else 1.0 / (max(abs(C0), D_total0, 1.0) ** 2)
    )

    Pago = cp.Variable(T, nonneg=True)
    C = cp.Variable(T + 1)
    D_total = cp.Variable(T + 1)

    constraints = [C[0] == C0, D_total[0] == D_total0]
    for t in range(1, T + 1):
        idx = t - 1
        constraints.append(C[t] == C[t - 1] + I[idx] - E[idx] - Pago[idx])
        constraints.append(D_total[t] == D_total[t - 1] - Pago[idx])
        constraints.append(D_total[t] >= 0)
        constraints.append(Pago[idx] <= D_total[t - 1])
        constraints.append(Pago[idx] <= tope_semanal[idx])
        constraints.append(C[t] >= L_survival)  # restricción DURA real (ya alcanzable)

    deficit_min = cp.pos(L_min - C[1:])  # desempate suave, no obligatorio
    objetivo = cp.Maximize(cp.sum(Pago) - alpha_tiebreak * cp.sum_squares(deficit_min))

    prob = cp.Problem(objetivo, constraints)
    prob.solve(solver=cp.OSQP, max_iter=50000, eps_abs=1e-6, eps_rel=1e-6)

    if (
        prob.status not in ["optimal", "optimal_inaccurate", "user_limit"]
        or Pago.value is None
    ):
        return np.zeros((n, T)), [0.0] * T, False

    pago_sem = np.clip(Pago.value, 0, None)
    C_opt = C.value

    D_restante = D0.copy()
    orden = np.argsort(D_restante)
    P = np.zeros((n, T))
    for t in range(T):
        restante_a_pagar = pago_sem[t]
        for i in orden:
            if restante_a_pagar <= 0:
                break
            abono = min(D_restante[i], restante_a_pagar)
            P[i, t] = abono
            D_restante[i] -= abono
            restante_a_pagar -= abono

    D_track = D0.copy()
    ratios = []
    for t in range(T):
        D_track = D_track - P[:, t]
        denom = C_opt[t + 1] + L_min
        denom = denom if abs(denom) > eps else eps
        ratios.append(float(D_track.sum() / denom))

    return P, ratios, True


def modelo2_resiliencia(I, E, C0, T, R, F, s, E_max=None, L_min=None, D0=None):
    """
    Modelo 2 - Máxima resiliencia financiera.

    NOTA: se conserva SOLO como referencia académica -- ya no aparece en el
    selector de la interfaz ni en la comparación interactiva de modelos. Por
    diseño, nunca paga ni un peso de deuda (solo reprograma egresos), así que
    no aporta al objetivo de "ver al modelo decidir cuánto pagar" que motivó
    esta versión del módulo. Queda documentada aquí por si se necesita
    referenciarla o retomarla más adelante.

    Extremo CONSERVADOR del espectro: por diseño nunca toca D0 ni paga un peso de
    deuda (solo reprograma egresos flexibles dentro de una ventana +-s semanas,
    respetando los egresos rígidos). Maximiza el piso de caja alcanzable.

    L_min/L_survival ya NO son restricción dura (solo referencia informativa): con
    flujos donde egresos > ingresos, el piso máximo alcanzable suele ser inferior a
    cualquier L_min positivo, y exigirlo como obligatorio vuelve el problema
    infeasible sin importar cómo se reacomoden los pagos.

    Se reporta también la "resiliencia ganada" (m_opt - m_base): cuánto mejora el
    piso gracias a la reprogramación vs. no reprogramar nada. Matemáticamente esta
    ganancia es $0 si la semana crítica cae en la ÚLTIMA semana del horizonte (la
    caja acumulada al final del horizonte es invariante a cómo se reordenen los
    pagos) -- no es un error, es una propiedad esperada del modelo.
    """
    I = np.asarray(I, dtype=float)
    E = np.asarray(E, dtype=float)
    R = np.asarray(R, dtype=float)
    F = np.asarray(F, dtype=float)
    if not (len(I) == len(E) == len(R) == len(F) == T):
        raise ValueError("I, E, R, F deben tener longitud T")
    if np.any(F < 0) or np.any(R < 0):
        raise ValueError("R y F deben ser no negativos")

    C_base = np.zeros(T + 1)
    C_base[0] = C0
    for t in range(1, T + 1):
        C_base[t] = C_base[t - 1] + I[t - 1] - E[t - 1]
    m_base = float(np.min(C_base[1:]))

    if np.all(F == 0):
        return dict(
            m_opt=m_base,
            m_base=m_base,
            ganancia=0.0,
            C_opt=C_base.tolist(),
            E_opt=E.tolist(),
            x_opt={},
            exito=True,
            t_crit=int(np.argmin(C_base[1:])) + 1,
            K_star=(np.max(E) / np.sum(E)) if np.sum(E) > 0 else np.nan,
        )

    pairs, idx_map, n_x = [], {}, 0
    for tau in range(T):
        t_start, t_end = max(0, tau - s), min(T, tau + s + 1)
        for t in range(t_start, t_end):
            pairs.append((tau, t))
            idx_map[(tau, t)] = n_x
            n_x += 1

    n_vars = n_x + 1
    idx_m = n_x
    c = np.zeros(n_vars)
    c[idx_m] = -1.0

    A_eq, b_eq = [], []
    for tau in range(T):
        row = np.zeros(n_vars)
        t_start, t_end = max(0, tau - s), min(T, tau + s + 1)
        for t in range(t_start, t_end):
            row[idx_map[(tau, t)]] = 1.0
        A_eq.append(row)
        b_eq.append(F[tau])
    A_eq, b_eq = np.array(A_eq), np.array(b_eq)

    A_ub, b_ub = [], []
    cum_I, cum_R = np.cumsum(I), np.cumsum(R)
    for t_idx in range(1, T + 1):
        row = np.zeros(n_vars)
        for k in range(t_idx):
            tau_start, tau_end = max(0, k - s), min(T, k + s + 1)
            for tau in range(tau_start, tau_end):
                if (tau, k) in idx_map:
                    row[idx_map[(tau, k)]] += 1.0
        row[idx_m] = 1.0
        A_ub.append(row)
        b_ub.append(C0 + cum_I[t_idx - 1] - cum_R[t_idx - 1])

    if E_max is not None:
        E_max = np.asarray(E_max, dtype=float)
        for t in range(T):
            if np.isfinite(E_max[t]):
                row = np.zeros(n_vars)
                tau_start, tau_end = max(0, t - s), min(T, t + s + 1)
                for tau in range(tau_start, tau_end):
                    row[idx_map[(tau, t)]] = 1.0
                A_ub.append(row)
                b_ub.append(E_max[t] - R[t])

    A_ub = np.array(A_ub) if A_ub else None
    b_ub = np.array(b_ub) if b_ub else None
    bounds = [(0, None)] * n_x + [(None, None)]

    res = linprog(
        c, A_ub=A_ub, b_ub=b_ub, A_eq=A_eq, b_eq=b_eq, bounds=bounds, method="highs"
    )

    if not res.success:
        return dict(
            m_opt=np.nan,
            m_base=m_base,
            ganancia=np.nan,
            C_opt=[],
            E_opt=[],
            x_opt={},
            exito=False,
            t_crit=None,
            K_star=np.nan,
        )

    x_sol = res.x[:n_x]
    m_opt = res.x[idx_m]
    x_opt = {pair: x_sol[idx_map[pair]] for pair in pairs}

    E_opt = np.copy(R)
    for (tau, t), val in x_opt.items():
        E_opt[t] += val

    C_opt = np.zeros(T + 1)
    C_opt[0] = C0
    for t in range(1, T + 1):
        C_opt[t] = C_opt[t - 1] + I[t - 1] - E_opt[t - 1]

    t_crit = int(np.argmin(C_opt[1:])) + 1
    sum_E_opt = np.sum(E_opt)
    K_star = np.max(E_opt) / sum_E_opt if sum_E_opt > 0 else np.nan

    return dict(
        m_opt=m_opt,
        m_base=m_base,
        ganancia=float(m_opt - m_base),
        C_opt=C_opt.tolist(),
        E_opt=E_opt.tolist(),
        x_opt=x_opt,
        exito=True,
        t_crit=t_crit,
        K_star=K_star,
    )


def modelo_b_ratio_pago(
    D0,
    I,
    E,
    C0,
    T,
    L_min=200_000.0,
    L_survival=10_000.0,
    pago_max_frac_ingreso=0.5,
    D_crit=5.0,
    piso_agres=0.05,
    epsilon=1e-6,
):
    """
    Modelo B - Ratio de Pago (heurística snowball con estrés dinámico), DECISIÓN LIBRE.

    HISTORIAL DEL DISEÑO: en versiones anteriores este modelo recibía un
    presupuesto de pago FIJO (p. ej. 25% de la deuda), porque con una caja
    estructuralmente negativa el mecanismo original -- "paga solo si hay
    excedente real por encima de un piso dinámico" -- nunca encontraba
    excedente (la caja jamás superaba ningún piso positivo), así que nunca
    pagaba nada.

    Con una caja inicial ya positiva, sí puede existir excedente real. Se
    regresa entonces al diseño original: el modelo YA NO recibe presupuesto
    fijo -- cada semana calcula cuánta caja tiene disponible por encima de un
    piso dinámico (entre L_survival y L_min, según qué tan estresada esté la
    situación), y paga una fracción de ese excedente real, modulada por la
    misma agresividad basada en estrés. Si no hay excedente, no paga -- puede
    pagar $0 varias semanas seguidas si la caja está justo en el piso.

    Snowball: el pago de cada semana se reparte entre las deudas individuales,
    de la más chica a la más grande, saldando cada una por completo antes de
    tocar la siguiente.
    """
    n = len(D0)
    D_restante = np.array(D0, dtype=float)
    C = C0
    P = np.zeros((n, T))
    ratios = []

    for t in range(T):
        C_disp = C + (I[t] - E[t])

        liq_ratio = (C_disp + L_min) / (L_min + epsilon)
        stress_liq = float(np.clip(1.0 - liq_ratio, 0.0, 1.0))
        denom = max(abs(C_disp) + L_min, epsilon)
        stress_debt = min(D_restante.sum() / denom / D_crit, 1.0)
        s_stress = max(stress_liq, stress_debt)

        agresividad = max(1.0 - s_stress, piso_agres)
        L_dyn = L_survival + (L_min - L_survival) * (1.0 - s_stress)
        excedente = max(0.0, C_disp - L_dyn)

        tope_semana = pago_max_frac_ingreso * I[t]
        pago_total = max(
            0.0, min(excedente * agresividad, tope_semana, D_restante.sum())
        )

        if pago_total > 0:
            orden = np.argsort(D_restante)
            restante = pago_total
            for i in orden:
                if restante <= 0:
                    break
                abono = min(D_restante[i], restante)
                P[i, t] = abono
                D_restante[i] -= abono
                restante -= abono

        pagado_esta_semana = P[:, t].sum()
        C = C_disp - pagado_esta_semana
        ratio_t = D_restante.sum() / (C + L_min + epsilon)
        ratios.append(ratio_t)

    return P, ratios, True


def construir_tabla_modelo(P, ratios, conceptos, I, E, C0, L_min):
    """Tabla de recomendación semana a semana, compartida por Modelo A y Modelo B
    (ambos reparten pagos por concepto de deuda)."""
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
        fila["Ratio D/C"] = round(float(ratios[t]), 4)
        if C[t] <= L_min:
            estado = "Sin margen"
        elif C[t] <= L_min * 1.5:
            estado = "Ajustado"
        else:
            estado = "Operable"
        fila["Estado"] = estado
        filas.append(fila)

    return pd.DataFrame(filas)


def construir_tabla_modelo2_resiliencia(res, L_min=None, L_survival=None):
    """Tabla de recomendación para Modelo 2 (Resiliencia): no reparte pagos por
    concepto, solo muestra la caja y el egreso reprogramado semana a semana, más
    las métricas de piso/resiliencia ganada.

    La columna "Semana" se guarda como texto (combina números de semana con
    etiquetas de métricas), y las columnas numéricas usan NaN (no "") para las
    celdas en blanco -- una columna de tipo mixto (número + string) rompe la
    serialización Arrow que usa Streamlit para mostrar tablas.
    """
    if not res["exito"]:
        return pd.DataFrame({"Error": ["Optimización fallida"]})
    T_mas_1 = len(res["C_opt"])
    df_tabla = pd.DataFrame(
        {
            "Semana": [str(t) for t in range(T_mas_1)],
            "Caja": res["C_opt"],
            "Egreso_reprogramado": [
                res["E_opt"][t - 1] if t > 0 else np.nan for t in range(T_mas_1)
            ],
        }
    )
    filas_metricas = [
        ("---", np.nan),
        ("Piso SIN reprogramar (m_base)", res["m_base"]),
        ("Piso CON reprogramación (m*)", res["m_opt"]),
        ("Resiliencia ganada (m* - m_base)", res["ganancia"]),
        ("Semana crítica", res["t_crit"]),
        ("Índice K*", res["K_star"]),
    ]
    if L_min is not None:
        filas_metricas.append(("L_min (colchón cómodo, referencia)", L_min))
    if L_survival is not None:
        filas_metricas.append(
            ("L_survival (piso de emergencia, referencia)", L_survival)
        )
    metricas = pd.DataFrame(
        {
            "Semana": [f[0] for f in filas_metricas],
            "Caja": [f[1] for f in filas_metricas],
            "Egreso_reprogramado": [np.nan] * len(filas_metricas),
        }
    )
    return pd.concat([df_tabla, metricas], ignore_index=True)


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
                            if categoria_escrita:
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
                st.caption(
                    "Ordenadas de menor a mayor: como el snowball paga primero las "
                    "deudas más chicas, así se ve de inmediato cuáles se liquidan antes."
                )
                st.dataframe(
                    deudas.rename("Importe pendiente ($)"), use_container_width=True
                )

                saldo_actual = calcular_saldo_actual(df)
                deuda_total = float(deudas.sum())

                col1, col2, col3 = st.columns(3)
                with col1:
                    horizonte = st.number_input(
                        "Horizonte de planeación (semanas, T)",
                        min_value=2,
                        max_value=52,
                        value=8,
                        step=1,
                    )
                with col2:
                    L_min = st.number_input(
                        "Colchón cómodo (L_min)",
                        min_value=0.0,
                        value=200_000.0,
                        step=500.0,
                        help="Nivel de caja con el que la operación está tranquila. Se usa "
                        "como referencia/desempate suave, no es obligatorio.",
                    )
                with col3:
                    L_survival = st.number_input(
                        "Piso de emergencia (L_survival)",
                        min_value=0.0,
                        value=10_000.0,
                        step=500.0,
                        help="Mínimo tolerable, nunca se cruza. Si la caja proyectada no "
                        "alcanza para mantenerse por encima de este piso ni pagando $0 de "
                        "deuda, el Modelo A no encontrará solución factible -- eso es una "
                        "señal real de que hace falta revisar el saldo inicial o la "
                        "proyección antes de pedirle al modelo pagar deuda.",
                    )

                st.markdown('<div class="regla-marca"></div>', unsafe_allow_html=True)

                modelo = st.selectbox(
                    "Modelo de optimización",
                    [
                        "Modelo A — Optimización Global (decide cuánta deuda pagar maximizando el total, sin cruzar el piso de emergencia)",
                        "Modelo B — Ratio de Pago (paga solo cuando hay excedente real de caja, a un ritmo que depende del estrés)",
                    ],
                )

                if modelo.startswith("Modelo A"):
                    st.caption(
                        "Ya no recibe un presupuesto fijo: decide por sí mismo cuánta deuda "
                        "pagar, maximizando el total pagado sin que la caja cruce nunca el "
                        "piso de emergencia. Reparte el pago con snowball estricto."
                    )
                else:
                    st.caption(
                        "Cada semana calcula si hay excedente real de caja por encima de un "
                        "piso dinámico (entre el piso de emergencia y el colchón cómodo, "
                        "según el estrés), y paga una fracción de ese excedente -- puede "
                        "pagar $0 varias semanas si no hay margen real. Snowball estricto."
                    )

                calcular = st.button("Calcular recomendaciones")

            if calcular:
                I, E = obtener_proyeccion(df, horizonte)
                C0 = saldo_actual
                conceptos = deudas.index.tolist()
                D0 = deudas.values.tolist()

                if modelo.startswith("Modelo A"):
                    P, ratios, exito = modelo_a_optimizacion_global(
                        D0, I, E, C0, L_min, L_survival, horizonte
                    )
                else:
                    P, ratios, exito = modelo_b_ratio_pago(
                        D0, I, E, C0, horizonte, L_min=L_min, L_survival=L_survival
                    )

                tabla = (
                    construir_tabla_modelo(P, ratios, conceptos, I, E, C0, L_min)
                    if exito
                    else None
                )

                if not exito or tabla is None:
                    st.error(
                        "El modelo no encontró una solución factible: dado el saldo inicial "
                        "y la proyección actuales, no existe forma de mantener la caja por "
                        "encima del piso de emergencia ni pagando $0 de deuda. Revisa el "
                        "saldo inicial (o usa el ajuste manual de arriba), el horizonte, o "
                        "el piso de emergencia."
                    )
                    st.session_state.pop("tabla_optimizacion", None)
                else:
                    st.session_state["tabla_optimizacion"] = tabla
                    st.session_state["modelo_optimizacion"] = modelo
                    st.session_state["proyeccion_optimizacion"] = (
                        f"Proyección adaptativa (excluye semanas en curso o no "
                        f"representativas): ingreso ≈ {moneda(I[0])}/semana, "
                        f"egreso ≈ {moneda(E[0])}/semana, caja inicial {moneda(C0)}."
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

            st.write("")
            st.markdown('<div class="regla-marca"></div>', unsafe_allow_html=True)
            with st.expander(
                "Comparar Modelo A vs. Modelo B (trade-off caja vs. deuda)"
            ):
                st.caption(
                    "Corre los dos modelos con el mismo horizonte, colchones, y saldo "
                    "inicial (incluyendo el ajuste manual si lo usaste arriba) -- ambos "
                    "deciden libremente cuánta deuda pagar, sin presupuesto forzado."
                )
                comparar = st.button("Generar comparación")

                if comparar:
                    I_c, E_c = obtener_proyeccion(df, horizonte)
                    C0_c = saldo_actual
                    D0_c = deudas.values.tolist()
                    D0_total_c = float(sum(D0_c))

                    P_a, ratios_a, exito_a = modelo_a_optimizacion_global(
                        D0_c, I_c, E_c, C0_c, L_min, L_survival, horizonte
                    )
                    P_b, ratios_b, exito_b = modelo_b_ratio_pago(
                        D0_c,
                        I_c,
                        E_c,
                        C0_c,
                        horizonte,
                        L_min=L_min,
                        L_survival=L_survival,
                    )

                    if not (exito_a and exito_b):
                        st.error(
                            "Alguno de los dos modelos no encontró solución factible con "
                            "estos parámetros; ajusta el saldo inicial, el horizonte, o los "
                            "colchones e intenta de nuevo."
                        )
                    else:

                        def caja_proyectada(P, I, E, C0):
                            IE = np.array(I, dtype=float) - np.array(E, dtype=float)
                            pagos_sem = P.sum(axis=0)
                            return C0 + np.cumsum(IE) - np.cumsum(pagos_sem)

                        def deuda_restante_proyectada(P, D0_total):
                            pagos_sem = P.sum(axis=0)
                            return D0_total - np.cumsum(pagos_sem)

                        caja_a = caja_proyectada(P_a, I_c, E_c, C0_c)
                        caja_b = caja_proyectada(P_b, I_c, E_c, C0_c)
                        deuda_a = deuda_restante_proyectada(P_a, D0_total_c)
                        deuda_b = deuda_restante_proyectada(P_b, D0_total_c)

                        resumen = pd.DataFrame(
                            {
                                "Modelo": [
                                    "Modelo A (Global)",
                                    "Modelo B (Ratio de Pago)",
                                ],
                                "Piso de caja mínimo": [caja_a.min(), caja_b.min()],
                                "Deuda final restante": [deuda_a[-1], deuda_b[-1]],
                                "% de deuda pagada": [
                                    100 * (1 - deuda_a[-1] / D0_total_c),
                                    100 * (1 - deuda_b[-1] / D0_total_c),
                                ],
                            }
                        )
                        st.dataframe(resumen, use_container_width=True, hide_index=True)

                        fig, ax_caja = plt.subplots(figsize=(10, 6))
                        ax_deuda = ax_caja.twinx()
                        semanas = list(range(1, horizonte + 1))
                        nombres = ["Modelo A (Global)", "Modelo B (Ratio de Pago)"]
                        colores = ["tab:blue", "tab:green"]
                        cajas = [caja_a, caja_b]
                        deudas_din = [deuda_a, deuda_b]

                        for nombre, color, caja_arr, deuda_arr in zip(
                            nombres, colores, cajas, deudas_din
                        ):
                            ax_caja.plot(
                                semanas,
                                caja_arr,
                                marker="o",
                                color=color,
                                linestyle="-",
                                label=f"Caja - {nombre}",
                            )
                            ax_deuda.plot(
                                semanas,
                                deuda_arr,
                                marker="s",
                                color=color,
                                linestyle="--",
                                alpha=0.6,
                                label=f"Deuda - {nombre}",
                            )

                        ax_caja.axhline(
                            L_min,
                            color="green",
                            linestyle=":",
                            linewidth=1,
                            label="L_min (colchón cómodo)",
                        )
                        ax_caja.axhline(
                            L_survival,
                            color="red",
                            linestyle=":",
                            linewidth=1,
                            label="L_survival (piso de emergencia)",
                        )
                        ax_caja.set_xlabel("Semana")
                        ax_caja.set_ylabel("Caja proyectada ($)")
                        ax_deuda.set_ylabel("Deuda restante ($)")
                        ax_caja.set_title(
                            "Trade-off caja vs. deuda (Modelo A vs. Modelo B)"
                        )

                        l1, e1 = ax_caja.get_legend_handles_labels()
                        l2, e2 = ax_deuda.get_legend_handles_labels()
                        ax_caja.legend(
                            l1 + l2,
                            e1 + e2,
                            loc="center left",
                            bbox_to_anchor=(1.15, 0.5),
                            fontsize=8,
                        )
                        plt.tight_layout()
                        st.pyplot(fig)

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
