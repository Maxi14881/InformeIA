import base64
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.formatting.rule import FormulaRule, ColorScaleRule
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Generador Informe QAStudio",
    page_icon="📤",
)

# Convierte la imagen a base64
def image_to_base64(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode()

# URL de la imagen y enlace a YouTube
# URL de la imagen y enlace a YouTube
logo_path = "binit.png"
link = "https://binit.tech/"
logo_base64 = image_to_base64(logo_path)

# Mostrar imagen enlazada
st.markdown(
    f'<a href="{link}" target="_blank">'
    f'<img src="data:image/jpeg;base64,{logo_base64}" style="width:100%;"/>'
    '</a>',
    unsafe_allow_html=True
)

# CSS personalizado
st.markdown(
    """
    <style>
        .stApp {
            background-color: #E0DBDB;
        }
        padding: 10px 15px;
        border-radius: 8px;
        font-weight: bold;
        color: white; /* Color del texto en blanco */
        background-color: #333333; /* Gris oscuro para el fondo */
        border: none;
        cursor: pointer;
        transition: background-color 0.3s ease;
        }
    </style>
    """,
    unsafe_allow_html=True
)

st.title("Generador de Informe QAStudio")

uploaded_file = st.file_uploader("Sube tu archivo Excel de origen", type=["xlsx", "xls"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()
    df['Resultado'] = pd.to_numeric(df['Resultado'], errors='coerce')

    version_agente = str(df['Versión del agente'].iloc[0]) if 'Versión del agente' in df.columns else "N/A"

    pivot = pd.pivot_table(
        df,
        index='Id del caso de prueba',
        columns='Código de ejecución',
        values='Resultado',
        aggfunc='mean',
        fill_value=0
    )

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        encabezado = pd.DataFrame({
            'A': [
                '', '', '', 
                'ANÁLISIS GENERAL DEL ASISTENTE', 'Versión del asistente', 'Fecha de ejecución', 'Cantidad de ejecuciones', '',
                'CRITERIO ACEPTACIÓN - REGLA 80 - 80', 'Total de casos analizados', 'Total de casos con efectividad >= 80%', '% casos > 80%',
                '', 
                'CRITERIO ACEPTACIÓN - NO FAILS', 'Total de casos sin casos de éxito'
            ],
            'B': ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '']
        })
        encabezado.to_excel(writer, index=False, header=False, startrow=0)
        pivot.to_excel(writer, startrow=16)

    output.seek(0)
    wb = load_workbook(output)
    ws = wb.active
    ws.title = "RESUMEN"
    ws["A2"] = "Informe de pruebas IA"
    ws["A2"].font = Font(bold=True, size=18)

    ws["A4"].font = Font(bold=True)
    ws["A9"].font = Font(bold=True)
    ws["A13"].font = Font(bold=True)

    ws["B4"] = '=IF(AND(B11>0.8,B15=0),"PASSED","FAILED")'
    ws["B5"] = version_agente
    ws["B6"] = datetime.today().strftime("%d/%m/%Y")
    ws["B7"] = df["Código de ejecución"].nunique()  # Cantidad de ejecuciones únicas

    fila_inicio = 18
    num_filas = pivot.shape[0]
    fila_fin = fila_inicio + num_filas - 1
    rango_g = f"G{fila_inicio}:G{fila_fin}"

    ws["B10"] = f'=COUNTA({rango_g})'
    ws["B11"] = f'=COUNTIF({rango_g},">=80%")'
    ws["B12"] = '=B11/B10'
    ws["B12"].number_format = "0%"
    ws["B15"] = f'=COUNTIF({rango_g},"=0")'

    ws["G17"] = "Total Gral"
    ws["G17"].font = Font(bold=True)

    col_inicio = 2
    col_fin = col_inicio + pivot.shape[1] - 1

    for i in range(num_filas):
        fila_excel = fila_inicio + i
        col_letra_ini = ws.cell(row=fila_excel, column=col_inicio).column_letter
        col_letra_fin = ws.cell(row=fila_excel, column=col_fin).column_letter
        ws[f"G{fila_excel}"] = f"=AVERAGE({col_letra_ini}{fila_excel}:{col_letra_fin}{fila_excel})"

    fila_total = fila_inicio + num_filas
    ws.cell(row=fila_total, column=1, value="Promedio columna").font = Font(bold=True)
    for col in range(col_inicio, col_fin + 2):
        col_letra = ws.cell(row=fila_inicio, column=col).column_letter
        ws.cell(row=fila_total, column=col, value=f"=AVERAGE({col_letra}{fila_inicio}:{col_letra}{fila_inicio + num_filas - 1})")
        ws.cell(row=fila_total, column=col).number_format = "0%"

    min_row = fila_inicio
    max_row = ws.max_row
    min_col = col_inicio
    max_col = col_fin + 1

    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            if isinstance(cell.value, (int, float)) or (isinstance(cell.value, str) and cell.value.startswith("=")):
                cell.number_format = "0%"

    ws.conditional_formatting.add(
        "B4",
        FormulaRule(formula=['B4="PASSED"'], fill=PatternFill(start_color="63BE7B", end_color="63BE7B", fill_type="solid"))
    )
    ws.conditional_formatting.add(
        "B4",
        FormulaRule(formula=['B4="FAILED"'], fill=PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid"))
    )
    ws.conditional_formatting.add(
        "B12",
        FormulaRule(formula=['B12>=0.8'], fill=PatternFill(start_color="63BE7B", end_color="63BE7B", fill_type="solid"))
    )
    ws.conditional_formatting.add(
        "B12",
        FormulaRule(formula=['B12<0.8'], fill=PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid"))
    )
    ws.conditional_formatting.add(
        "B15",
        FormulaRule(formula=['B15=0'], fill=PatternFill(start_color="63BE7B", end_color="63BE7B", fill_type="solid"))
    )
    ws.conditional_formatting.add(
        "B15",
        FormulaRule(formula=['B15<>0'], fill=PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid"))
    )

    col_letra_ini = get_column_letter(min_col)
    col_letra_fin = get_column_letter(max_col)
    rango_color = f"{col_letra_ini}{min_row}:{col_letra_fin}{fila_fin}"

    ws.conditional_formatting.add(
        rango_color,
        ColorScaleRule(
            start_type='num', start_value=0, start_color='FFFC7B7B',
            mid_type='num', mid_value=0.5, mid_color='FFFFFF00',
            end_type='num', end_value=1, end_color='FF63BE7B'
        )
    )

    thin = Side(border_style="thin", color="000000")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)
    for row in ws.iter_rows(min_row=4, max_row=16, min_col=1, max_col=2):
        for cell in row:
            cell.border = border

    # --- Copiar hoja "Ejecuciones" si existe ---
    xls = pd.ExcelFile(uploaded_file)
    if "Ejecuciones" in xls.sheet_names:
        df_ejec = pd.read_excel(xls, sheet_name="Ejecuciones")
        ws_ejec = wb.create_sheet("EJECUCIONES")
        for i, col in enumerate(df_ejec.columns, 1):
            ws_ejec.cell(row=1, column=i, value=col)
        for row_idx, row in enumerate(df_ejec.itertuples(index=False), 2):
            for col_idx, value in enumerate(row, 1):
                ws_ejec.cell(row=row_idx, column=col_idx, value=value)

    # --- Copiar hoja "CRITERIOS ACEPTACIÓN" ---
    columnas_criterios = ["Id del caso de prueba", "Caso de prueba", "Resultado esperado"]
    if all(col in df.columns for col in columnas_criterios):
        criterios_df = df[columnas_criterios].drop_duplicates().sort_values(by="Id del caso de prueba")
        ws_criterios = wb.create_sheet("CRITERIOS ACEPTACIÓN")
        for i, col in enumerate(criterios_df.columns, 1):
            ws_criterios.cell(row=1, column=i, value=col)
        for row_idx, row in enumerate(criterios_df.itertuples(index=False), 2):
            for col_idx, value in enumerate(row, 1):
                ws_criterios.cell(row=row_idx, column=col_idx, value=value)

    # Guardar el archivo final
    output_final = BytesIO()
    wb.save(output_final)
    output_final.seek(0)

    st.success("¡Archivo generado correctamente!")
    st.download_button(
        label="Descargar QAStudio Result.xlsx",
        data=output_final,
        file_name="QAStudio Result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
