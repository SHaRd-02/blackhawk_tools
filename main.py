import io  # Requerido para el manejo de archivos en memoria
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

st.title("Reportes Parker")

# 1. Display the file uploader widget
uploaded_file = st.file_uploader("Archivo Empleados", type=["xls"])

# 2. Process the file if it exists
if uploaded_file is not None:
    # A. Leer el archivo sin headers
    df = pd.read_excel(uploaded_file, engine="xlrd", header=None)
    rows_total = len(df)

    # B. Limpieza de filas inicial
    df = df.iloc[1 : (rows_total - 1)]
    df = df.reset_index(drop=True)

    # C. REASIGNAR LOS HEADERS REALES
    df.columns = df.iloc[0]
    df = df.iloc[1:]
    df = df.reset_index(drop=True)

    # Renombrar columnas
    df = df.rename(
        columns={
            "Ti me": "Time",
            "VendorItem": "Item Description",
            "UOM (CField1)": "UOM",
            "OnHand": "Qty",
            "IssueCost": "Price",
            "ExtIssCost": "Total Cost",
        }
    )

    # Seleccionar solo las columnas deseadas
    wanted_columns = [
        "Crib",
        "Date",
        "Time",
        "Item Code",
        "Item Description",
        "UOM",
        "Qty",
        "Pkg",
        "Price",
        "Total Cost",
        "Employee",
        "Dept.",
    ]
    df = df[wanted_columns]

    # --- Convertir columna numérica para los cálculos ---
    df["Price"] = pd.to_numeric(df["Price"], errors="coerce").fillna(0)
    df["Total Cost"] = pd.to_numeric(df["Total Cost"] * -1, errors="coerce").fillna(0)

    # Crear las dos tablas dinámicas de resumen
    df_dept = (
        df.pivot_table(
            index="Dept.", values="Total Cost", aggfunc="sum", fill_value=0
        )
        .reset_index()
        .rename(columns={"Total Cost": "Costo Total"})
    )

    df_crib = (
        df.pivot_table(
            index="Crib", values="Total Cost", aggfunc="sum", fill_value=0
        )
        .reset_index()
        .rename(columns={"Total Cost": "Costo Total"})
    )

    # Display data principal en Streamlit
    st.write("### Preview of Uploaded Data:")
    st.dataframe(df)

    # Mostrar resúmenes en la interfaz
    st.write("### 📊 Resúmenes Generados")
    col1, col2 = st.columns(2)
    with col1:
        st.write("**Por Departamento**")
        st.dataframe(df_dept)
    with col2:
        st.write("**Por Crib**")
        st.dataframe(df_crib)

    # ========================================================
    # 3. GENERAR REPORTE ÚNICO EN EXCEL CON ESTILOS Y POSICIONES
    # ========================================================
    buffer = io.BytesIO()

    # Definir estilos de OpenPyXL
    header_fill = PatternFill(
        start_color="1F4E78", end_color="1F4E78", fill_type="solid"
    )
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    total_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    center_alignment = Alignment(horizontal="center", vertical="center")
    accounting_format = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        sheet_name = "Reporte Automatizado"

        # 1. Escribir la tabla principal dejando la fila 1 libre para la suma superior
        df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=1, startcol=0)

        worksheet = writer.sheets[sheet_name]

        # --- APLICAR ESTILOS A LA TABLA PRINCIPAL ---
        # Encabezados en Fila 2 (A2:L2)
        for col_num in range(1, len(df.columns) + 1):
            cell = worksheet.cell(row=2, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_alignment

        # Suma total arriba de la columna Total Cost (Columna J, Fila 1)
        main_start_row = 3
        main_end_row = main_start_row + len(df) - 1
        worksheet.cell(
            row=1, column=10, value=f"=SUM(J{main_start_row}:J{main_end_row})"
        )
        worksheet.cell(row=1, column=10).font = Font(
            name="Calibri", size=11, bold=True
        )
        worksheet.cell(row=1, column=10).number_format = accounting_format

        # Aplicar formato Accounting a columnas Price (I) y Total Cost (J)
        for r in range(main_start_row, main_end_row + 1):
            worksheet.cell(row=r, column=9).number_format = accounting_format
            worksheet.cell(row=r, column=10).number_format = accounting_format

        # ----------------------------------------------------
        # BLOQUE 1: TABLA DEPARTAMENTO
        # ----------------------------------------------------
        worksheet.cell(row=1, column=16, value="DEPARTMENT").font = Font(
            name="Calibri", size=11, bold=True
        )
        worksheet.cell(row=2, column=16, value="Sum of Total Cost").font = Font(
            name="Calibri", size=10, italic=True
        )

        df_dept.to_excel(
            writer, index=False, sheet_name=sheet_name, startrow=2, startcol=15
        )

        # Estilo headers Dept
        for col_idx in [16, 17]:
            c = worksheet.cell(row=3, column=col_idx)
            c.fill = header_fill
            c.font = header_font
            c.alignment = center_alignment

        dept_start_row = 4
        dept_end_row = dept_start_row + len(df_dept) - 1
        dept_total_row = dept_end_row + 1

        # Grand Total Dept
        worksheet.cell(row=dept_total_row, column=16, value="Grand Total")
        worksheet.cell(
            row=dept_total_row,
            column=17,
            value=f"=SUM(Q{dept_start_row}:Q{dept_end_row})",
        )

        for col_idx in [16, 17]:
            c = worksheet.cell(row=dept_total_row, column=col_idx)
            c.fill = header_fill
            c.font = total_font

        for r in range(dept_start_row, dept_total_row + 1):
            worksheet.cell(row=r, column=17).number_format = accounting_format

        # ----------------------------------------------------
        # BLOQUE 2: TABLA CRIB
        # ----------------------------------------------------
        title_crib_row = dept_total_row + 4

        worksheet.cell(row=title_crib_row, column=16, value="CRIB").font = Font(
            name="Calibri", size=11, bold=True
        )
        worksheet.cell(
            row=title_crib_row + 1, column=16, value="Sum of Total Cost"
        ).font = Font(name="Calibri", size=10, italic=True)

        df_crib.to_excel(
            writer,
            index=False,
            sheet_name=sheet_name,
            startrow=title_crib_row + 1,
            startcol=15,
        )

        # Estilo headers Crib
        for col_idx in [16, 17]:
            c = worksheet.cell(row=title_crib_row + 2, column=col_idx)
            c.fill = header_fill
            c.font = header_font
            c.alignment = center_alignment

        crib_start_row = title_crib_row + 3
        crib_end_row = crib_start_row + len(df_crib) - 1
        crib_total_row = crib_end_row + 1

        # Grand Total Crib
        worksheet.cell(row=crib_total_row, column=16, value="Grand Total")
        worksheet.cell(
            row=crib_total_row,
            column=1,
            value=f"=SUM(Q{crib_start_row}:Q{crib_end_row})",
        )  # Corregido a columna 17 (Q)
        worksheet.cell(
            row=crib_total_row,
            column=17,
            value=f"=SUM(Q{crib_start_row}:Q{crib_end_row})",
        )

        for col_idx in [16, 17]:
            c = worksheet.cell(row=crib_total_row, column=col_idx)
            c.fill = header_fill
            c.font = total_font

        for r in range(crib_start_row, crib_total_row + 1):
            worksheet.cell(row=r, column=17).number_format = accounting_format

        # Ajuste automático del ancho de todas las columnas de la hoja
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = col[0].column_letter
            worksheet.column_dimensions[col_letter].width = max(
                max_len + 3, 12
            )

    buffer.seek(0)

    # Botón de descarga en la interfaz de Streamlit
    st.download_button(
        label="📥 Descargar Reporte Automatizado",
        data=buffer,
        file_name="reporte_parker_automatizado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )