import pandas as pd
import streamlit as st
from datetime import date, timedelta
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from openpyxl.utils import get_column_letter

ARCHIVO_DATOS = "Datos flujo de efectivo.xlsx"
ARCHIVO_SALIDA = "Flujo de efectivo.xlsx"

MESES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
    5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
    9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
}


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

        azul_encabezado = PatternFill(fill_type="solid", fgColor="1F4E78")
        azul_claro = PatternFill(fill_type="solid", fgColor="D9EAF7")
        azul_flujo = PatternFill(fill_type="solid", fgColor="9DC3E6")

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
            cell.fill = azul_encabezado

        for row in ws.iter_rows():
            if row[0].value in ["TOTAL INGRESOS", "TOTAL EGRESOS", "SALDO"]:
                for cell in row:
                    cell.font = Font(bold=True)
                    cell.fill = azul_claro

            if row[0].value == "FLUJO NETO":
                for cell in row:
                    cell.font = Font(bold=True, size=13)
                    cell.fill = azul_flujo

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


st.set_page_config(page_title="Flujo de efectivo", layout="wide")
st.title("Sistema de flujo de efectivo")

df = pd.read_excel(ARCHIVO_DATOS)

df["inicio_semana"] = pd.to_datetime(df["inicio_semana"], dayfirst=True)
df["tipo"] = df["tipo"].astype(str).str.strip()
df["categoría"] = df["categoría"].astype(str).str.strip()
df["descripción"] = df["descripción"].astype(str).str.strip()

if "inicio_semana" not in st.session_state:
    st.session_state.inicio_semana = date.today()

if "fin_semana" not in st.session_state:
    st.session_state.fin_semana = date.today() + timedelta(days=5)

if "concepto_seleccionado" not in st.session_state:
    st.session_state.concepto_seleccionado = None

if "categoria_seleccionada" not in st.session_state:
    st.session_state.categoria_seleccionada = None


tab1, tab2, tab3 = st.tabs([
    "Agregar movimiento",
    "Eliminar movimiento",
    "Flujo de efectivo"
])


with tab1:
    st.subheader("Agregar movimiento")

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

    st.markdown("---")

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


with tab2:
    st.subheader("Eliminar movimiento")

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


with tab3:
    st.subheader("Flujo de efectivo actualizado")

    df_actual = pd.read_excel(ARCHIVO_DATOS)
    flujo = generar_flujo(df_actual)

    st.dataframe(flujo, use_container_width=True)

    with open(ARCHIVO_SALIDA, "rb") as archivo:
        st.download_button(
            label="Descargar Excel",
            data=archivo,
            file_name=ARCHIVO_SALIDA,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )