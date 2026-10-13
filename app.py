import io
import os
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Calculadora de Tarimas", layout="centered")

st.title("Calculadora de Porcentajes por Tarima")

# Archivo de almacenamiento persistente
EXCEL_FILE = "registros_tarimas.xlsx"

# Traducción de días de la semana a español
DIAS_SEMANA = {
    "Monday": "Lunes",
    "Tuesday": "Martes",
    "Wednesday": "Miércoles",
    "Thursday": "Jueves",
    "Friday": "Viernes",
    "Saturday": "Sábado",
    "Sunday": "Domingo",
}


def cargar_datos():
    if os.path.exists(EXCEL_FILE):
        return pd.read_excel(EXCEL_FILE)
    else:
        return pd.DataFrame(
            columns=[
                "Día",
                "Fecha",
                "Hora",
                "Nº / ID Tarima",
                "Peso Total (FI)",
                "Merma",
                "% Merma",
                "Bueno",
                "% Bueno",
                "No Colocable",
                "% No Colocable",
            ]
        )


def generar_excel_estilizado(df):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reporte de Tarimas"

    # Mostrar líneas de la cuadrícula
    ws.views.sheetView[0].showGridLines = True

    # 1. Título del Reporte
    ws.merge_cells("A1:K1")
    title_cell = ws["A1"]
    title_cell.value = "REPORTE Y CONTROL DE TARIMAS"
    title_cell.font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(
        start_color="1F4E79", end_color="1F4E79", fill_type="solid"
    )
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35

    # 2. Encabezados con colores por categoría
    headers = [
        "Nº / ID Tarima",
        "Día",
        "Fecha",
        "Hora",
        "Peso Total (FI)",
        "Merma",
        "% Merma",
        "Bueno",
        "% Bueno",
        "No Colocable",
        "% No Colocable",
    ]

    header_fills = [
        PatternFill(
            start_color="2F5597", end_color="2F5597", fill_type="solid"
        ),  # Azul - Identificación
        PatternFill(
            start_color="2F5597", end_color="2F5597", fill_type="solid"
        ),
        PatternFill(
            start_color="2F5597", end_color="2F5597", fill_type="solid"
        ),
        PatternFill(
            start_color="2F5597", end_color="2F5597", fill_type="solid"
        ),
        PatternFill(
            start_color="215E21", end_color="215E21", fill_type="solid"
        ),  # Verde Oscuro - Peso Total
        PatternFill(
            start_color="C00000", end_color="C00000", fill_type="solid"
        ),  # Rojo - Merma
        PatternFill(
            start_color="C00000", end_color="C00000", fill_type="solid"
        ),
        PatternFill(
            start_color="385723", end_color="385723", fill_type="solid"
        ),  # Verde - Bueno
        PatternFill(
            start_color="385723", end_color="385723", fill_type="solid"
        ),
        PatternFill(
            start_color="C55A11", end_color="C55A11", fill_type="solid"
        ),  # Naranja - No Colocable
        PatternFill(
            start_color="C55A11", end_color="C55A11", fill_type="solid"
        ),
    ]

    ws.row_dimensions[3].height = 28

    header_border = Border(
        left=Side(style="thin", color="FFFFFF"),
        right=Side(style="thin", color="FFFFFF"),
        top=Side(style="medium", color="1F4E79"),
        bottom=Side(style="medium", color="1F4E79"),
    )

    for col_num, (header_text, fill) in enumerate(
        zip(headers, header_fills), 1
    ):
        cell = ws.cell(row=3, column=col_num, value=header_text)
        cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        cell.fill = fill
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        cell.border = header_border

    # 3. Filas de datos
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    zebra_fill = PatternFill(
        start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
    )

    for row_idx, row in df.iterrows():
        current_row = row_idx + 4
        ws.row_dimensions[current_row].height = 22
        is_even = current_row % 2 == 0

        def clean_pct(val):
            if isinstance(val, str):
                return float(val.replace("%", "").strip()) / 100.0
            val_float = float(val)
            return val_float / 100.0 if val_float > 1 else val_float

        valores = [
            row.get("Nº / ID Tarima", ""),
            row.get("Día", ""),
            row.get("Fecha", ""),
            row.get("Hora", ""),
            float(row.get("Peso Total (FI)", 0)),
            float(row.get("Merma", 0)),
            clean_pct(row.get("% Merma", 0)),
            float(row.get("Bueno", 0)),
            clean_pct(row.get("% Bueno", 0)),
            float(row.get("No Colocable", row.get("Procesado", 0))),
            clean_pct(row.get("% No Colocable", row.get("% Procesado", 0))),
        ]

        for col_idx, val in enumerate(valores, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=val)
            cell.font = Font(name="Calibri", size=11)
            cell.border = thin_border

            if is_even:
                cell.fill = zebra_fill

            if col_idx in [1, 2, 3, 4]:
                cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )
            elif col_idx in [5, 6, 8, 10]:
                cell.alignment = Alignment(
                    horizontal="right", vertical="center"
                )
                cell.number_format = "#,##0.0"
            elif col_idx in [7, 9, 11]:
                cell.alignment = Alignment(
                    horizontal="right", vertical="center"
                )
                cell.number_format = "0.0%"

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = 16

    wb.save(output)
    return output.getvalue()


# --- ENTRADAS DE DATOS ---
st.subheader("Ingreso de Datos de la Tarima")

tarima_id = st.text_input("Número / ID de Tarima o Lote:", value="")
fi = st.number_input("Peso Total Inicial (FI):", min_value=0.0, value=0.0)
merma = st.number_input("Peso de Merma:", min_value=0.0, value=0.0)
bueno = st.number_input("Peso de Bueno:", min_value=0.0, value=0.0)
no_colocable = st.number_input("Peso No Colocable:", min_value=0.0, value=0.0)

# VALIDACIONES
if fi <= 0:
    st.info(
        "Ingresa un Peso Total Inicial (FI) mayor a 0 para calcular los porcentajes."
    )
else:
    suma_partes = merma + bueno + no_colocable

    if suma_partes > fi:
        st.error(
            f"Error: La suma de las partes ({suma_partes:.2f}) sobrepasa el Peso Total Inicial ({fi:.2f})."
        )
    else:
        # Cálculos de Porcentajes
        pct_merma = (merma / fi) * 100
        pct_bueno = (bueno / fi) * 100
        pct_no_colocable = (no_colocable / fi) * 100

        # Vista previa del cálculo
        data_preview = {
            "Concepto": ["Merma", "Bueno", "No Colocable"],
            "Peso (lb/kg)": [merma, bueno, no_colocable],
            "Porcentaje (%)": [
                f"{pct_merma:.1f}%",
                f"{pct_bueno:.1f}%",
                f"{pct_no_colocable:.1f}%",
            ],
        }
        st.subheader("Resumen de Porcentajes")
        st.table(pd.DataFrame(data_preview))

        # Botón Guardar
        if st.button("Guardar Registro en Excel"):
            df_existente = cargar_datos()

            ahora = datetime.now()
            dia_esp = DIAS_SEMANA[ahora.strftime("%A")]
            fecha_str = ahora.strftime("%Y-%m-%d")
            hora_str = ahora.strftime("%H:%M:%S")

            nuevo_registro = pd.DataFrame(
                [
                    {
                        "Día": dia_esp,
                        "Fecha": fecha_str,
                        "Hora": hora_str,
                        "Nº / ID Tarima": (
                            tarima_id if tarima_id else "Sin Identificador"
                        ),
                        "Peso Total (FI)": fi,
                        "Merma": merma,
                        "% Merma": f"{pct_merma:.1f}%",
                        "Bueno": bueno,
                        "% Bueno": f"{pct_bueno:.1f}%",
                        "No Colocable": no_colocable,
                        "% No Colocable": f"{pct_no_colocable:.1f}%",
                    }
                ]
            )

            df_actualizado = pd.concat(
                [df_existente, nuevo_registro], ignore_index=True
            )
            df_actualizado.to_excel(EXCEL_FILE, index=False)
            st.success("¡Registro guardado exitosamente!")
            st.rerun()

# --- HISTORIAL Y BORRADO ---
st.divider()
st.subheader("Historial de Registros Guardados")

df_historial = cargar_datos()

if not df_historial.empty:
    st.dataframe(df_historial)

    # Eliminar registro
    st.markdown("### Eliminar un registro")

    opciones_eliminar = [
        f"Fila {idx + 1} - {row['Nº / ID Tarima']} ({row['Fecha']} {row['Hora']})"
        for idx, row in df_historial.iterrows()
    ]

    registro_a_eliminar = st.selectbox(
        "Selecciona la fila a eliminar:", opciones_eliminar
    )

    if st.button("Eliminar Fila Seleccionada"):
        indice_eliminar = opciones_eliminar.index(registro_a_eliminar)
        df_historial = df_historial.drop(index=indice_eliminar).reset_index(
            drop=True
        )
        df_historial.to_excel(EXCEL_FILE, index=False)
        st.success("¡Registro eliminado!")
        st.rerun()

    st.divider()

    # Botón de descarga con formato estilizado
    excel_binario = generar_excel_estilizado(df_historial)

    st.download_button(
        label="Descargar Reporte Formateado en Excel",
        data=excel_binario,
        file_name="Reporte_Tarimas_Formateado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Aún no hay registros guardados en el historial.")
