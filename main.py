import io  # Requerido para el manejo de archivos en memoria
import pandas as pd
import streamlit as st

st.title("Reportes Parker")

# 1. Display the file uploader widget
uploaded_file = st.file_uploader("Archivo Empleados", type=["xls"])

# Definimos el estilo correcto usando 'th' para los headers
header_style = {
    "selector": "th",  # 'th' es exclusivo para la fila de títulos
    "props": [
        ("background-color", "#1F4E78"),  # Azul oscuro
        ("color", "white"),  # Texto blanco
        ("font-weight", "bold"),  # Texto en negrita
        ("text-align", "center"),  # Centrar texto
    ],
}

# 2. Process the file if it exists
if uploaded_file is not None:
    # A. Leer el archivo sin headers
    df = pd.read_excel(uploaded_file, engine="xlrd", header=None)
    rows_total = len(df)

    # B. Limpieza de filas inicial (tu lógica original)
    df = df.iloc[1 : (rows_total - 1)]
    df = df.reset_index(drop=True)

    # C. REASIGNAR LOS HEADERS REALES (¡Primero hacemos la estructura!)
    # Tomamos la nueva primera fila (índice 0) que contiene los nombres reales
    df.columns = df.iloc[0]
    # Eliminamos esa fila de los datos y reiniciamos el índice
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

    # D. APLICAR EL ESTILO AL FINAL (Una vez que el DataFrame está limpio y estructurado)
    df_estilizado = df.style.set_table_styles([header_style])

    # Display data (Streamlit mostrará la tabla con el header de color azul)
    st.write("### Preview of Uploaded Data:")
    st.dataframe(df_estilizado)

    # ========================================================
    # 3. GENERAR ARCHIVO EXCEL Y BOTÓN DE DESCARGA
    # ========================================================
    # Crear buffer en memoria
    buffer = io.BytesIO()

    # Escribir a Excel controlando el formato y ancho de columnas
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_estilizado.to_excel(
            writer, index=False, sheet_name="Reporte Final"
        )

        # Ajuste automático del ancho de las columnas para que no se corte el texto
        worksheet = writer.sheets["Reporte Final"]
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = col[0].column_letter
            worksheet.column_dimensions[col_letter].width = max(
                max_len + 3, 12
            )

    buffer.seek(0)

    # Botón de descarga en la interfaz de Streamlit
    st.download_button(
        label="📥 Descargar Reporte en Excel",
        data=buffer,
        file_name="reporte_parker_final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
