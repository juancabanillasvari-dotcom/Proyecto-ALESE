import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Seguimiento de Unidades - ALESE", layout="wide")

st.title("🚗 Control de Estado de Unidades — Cliente ALESE")
st.markdown("Sube tu reporte nativo de Excel para obtener el estado clasificado, correlativo y ordenado de más antiguo a más actual.")

uploaded_file = st.file_uploader("Sube tu archivo Excel (ej. Automate_diario.xlsx)", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        sheet_target = 'SeguimientoServicio' if 'SeguimientoServicio' in pd.ExcelFile(uploaded_file).sheet_names else 0
        df = pd.read_excel(uploaded_file, sheet_name=sheet_target)

    # Filtrar solo cliente ALESE y unidades que NO estén RETIRADAS
    df_alese = df[
        (df['Cliente'].astype(str).str.contains('ALESE', case=False, na=False)) &
        (df['Ubicacion'].astype(str).str.upper() != 'RETIRADO')
    ].copy()

    if df_alese.empty:
        st.warning("No se encontraron unidades en almacén para el cliente ALESE en el archivo subido.")
    else:
        piv = df_alese.pivot_table(
            index=['Cliente', 'Almacen', 'Vin', 'Marca', 'Modelo', 'Color Vehiculo'],
            columns='SubServicio',
            values=['Estado Vehiculo', 'Fecha Termino', 'Fecha Ingreso'],
            aggfunc='first'
        )
        piv.columns = [f'{col[0]}_{col[1]}' for col in piv.columns]
        piv = piv.reset_index()

        def evaluar_estado_unidad(row):
            pdi_final_est = str(row.get('Estado Vehiculo_PDI FINAL', '')).upper()
            pdi_inicial_est = str(row.get('Estado Vehiculo_PDI INICIAL COMPLETO', '')).upper()

            if pdi_final_est.startswith("FINALIZADO"):
                return "FINALIZADO"
            if "EN COLA" in pdi_inicial_est or pdi_inicial_est in ["NAN", "NONE", ""]:
                return "EN COLA"
            if "EN COLA" in pdi_final_est or "EN PROCESO" in pdi_final_est or "EN PROCESO" in pdi_inicial_est:
                return "EN PROCESO"
            return "EN COLA"

        piv['Estado Vehiculo'] = piv.apply(evaluar_estado_unidad, axis=1)

        # Asignar fecha cruda para ordenamiento estricto
        def asignar_fecha_raw(row):
            estado = row['Estado Vehiculo']
            if estado in ["EN COLA", "EN PROCESO"]:
                return row.get('Fecha Ingreso_PDI INICIAL COMPLETO', row.get('Fecha Ingreso_PDI FINAL', None))
            else:
                return row.get('Fecha Termino_PDI FINAL', row.get('Fecha Termino_PDI INICIAL COMPLETO', None))

        piv['Fecha_Raw'] = piv.apply(asignar_fecha_raw, axis=1)
        
        # Conversión a datetime interpretando el día primero (dayfirst=True) para orden cronológico estricto
        piv['Fecha_Datetime'] = pd.to_datetime(piv['Fecha_Raw'], errors='coerce', dayfirst=True)

        # Extraer exactamente los 10 primeros caracteres del archivo nativo de izquierda a derecha
        def extraer_10_chars(row):
            val = row['Fecha_Raw']
            if pd.isna(val) or str(val).strip() in ["", "nan", "NaT", "None"]:
                return ""
            return str(val).strip()[:10]

        piv['Fecha_Str'] = piv.apply(extraer_10_chars, axis=1)

        # Ordenar de menor a mayor (más antiguo a más actual)
        estado_order = {"FINALIZADO": 1, "EN PROCESO": 2, "EN COLA": 3}
        piv['Orden_Estado'] = piv['Estado Vehiculo'].map(estado_order)
        piv = piv.sort_values(by=['Orden_Estado', 'Fecha_Datetime'], ascending=[True, True])

        # Métricas principales
        st.markdown("---")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("📦 Total en Almacén", len(piv))
        m2.metric("🟢 Finalizadas", (piv['Estado Vehiculo'] == "FINALIZADO").sum())
        m3.metric("🟡 En Proceso", (piv['Estado Vehiculo'] == "EN PROCESO").sum())
        m4.metric("🔴 En Cola", (piv['Estado Vehiculo'] == "EN COLA").sum())
        st.markdown("---")

        # Función para renderizar tablas HTML estilizadas y listas para copiar al correo
        def render_styled_table(df_sub, col_fecha_name, header_color):
            df_display = pd.DataFrame({
                'N°': range(1, len(df_sub) + 1),
                'Cliente': df_sub['Cliente'],
                'Almacen': df_sub['Almacen'],
                'Vin': df_sub['Vin'],
                'Marca': df_sub['Marca'],
                'Modelo': df_sub['Modelo'],
                'Color Vehiculo': df_sub['Color Vehiculo'],
                col_fecha_name: df_sub['Fecha_Str']
            })

            # HTML + CSS para formateo corporativo compatible con copia a Outlook/Gmail
            html = f"""
            <style>
                .custom-table {{
                    width: 100%;
                    border-collapse: collapse;
                    font-family: 'Calibri', 'Segoe UI', sans-serif;
                    font-size: 13px;
                    margin-bottom: 20px;
                }}
                .custom-table th {{
                    background-color: {header_color};
                    color: white;
                    font-weight: bold;
                    text-align: center;
                    padding: 8px 12px;
                    border: 1px solid #d9d9d9;
                }}
                .custom-table td {{
                    padding: 6px 10px;
                    border: 1px solid #d9d9d9;
                }}
                .custom-table tr:nth-child(even) {{
                    background-color: #f9fafb;
                }}
                .center {{ text-align: center; }}
                .left {{ text-align: left; }}
            </style>
            <table class="custom-table">
                <thead>
                    <tr>
                        <th>N°</th>
                        <th>Cliente</th>
                        <th>Almacen</th>
                        <th>Vin</th>
                        <th>Marca</th>
                        <th>Modelo</th>
                        <th>Color Vehiculo</th>
                        <th>{col_fecha_name}</th>
                    </tr>
                </thead>
                <tbody>
            """
            for row in df_display.itertuples(index=False):
                html += f"""
                    <tr>
                        <td class="center">{row[0]}</td>
                        <td class="left">{row[1]}</td>
                        <td class="left">{row[2]}</td>
                        <td class="center"><b>{row[3]}</b></td>
                        <td class="left">{row[4]}</td>
                        <td class="left">{row[5]}</td>
                        <td class="left">{row[6]}</td>
                        <td class="center">{row[7]}</td>
                    </tr>
                """
            html += "</tbody></table>"
            return html, df_display

        # Separar por estados
        df_fin_raw = piv[piv['Estado Vehiculo'] == "FINALIZADO"]
        df_proc_raw = piv[piv['Estado Vehiculo'] == "EN PROCESO"]
        df_cola_raw = piv[piv['Estado Vehiculo'] == "EN COLA"]

        # Generar HTMLs estilizados
        html_fin, df_fin_exp = render_styled_table(df_fin_raw, "Fecha Término", "#1F4E79")
        html_proc, df_proc_exp = render_styled_table(df_proc_raw, "Fecha Ingreso", "#C65911")
        html_cola, df_cola_exp = render_styled_table(df_cola_raw, "Fecha Ingreso", "#595959")

        # Mostrar en Streamlit
        st.subheader("🟢 Unidades Finalizadas (Fecha Término)")
        st.markdown(html_fin, unsafe_allow_html=True)

        st.subheader("🟡 Unidades En Proceso (Fecha Ingreso)")
        st.markdown(html_proc, unsafe_allow_html=True)

        st.subheader("🔴 Unidades En Cola (Fecha Ingreso)")
        st.markdown(html_cola, unsafe_allow_html=True)

        # Generar Excel formateado profesionalmente con openpyxl
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_fin_exp.to_excel(writer, index=False, sheet_name='Finalizado')
            df_proc_exp.to_excel(writer, index=False, sheet_name='En Proceso')
            df_cola_exp.to_excel(writer, index=False, sheet_name='En Cola')

            # Darle diseño profesional al Excel generado
            wb = writer.book
            headers_colors = {"Finalizado": "1F4E79", "En Proceso": "C65911", "En Cola": "595959"}
            import openpyxl.styles as opts
            for sheetname in wb.sheetnames:
                ws = wb[sheetname]
                ws.views.sheetView[0].showGridLines = True
                fill_color = headers_colors.get(sheetname, "1F4E79")
                fill = opts.PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
                font = opts.Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                border = opts.Border(
                    left=opts.Side(style='thin', color='D9D9D9'),
                    right=opts.Side(style='thin', color='D9D9D9'),
                    top=opts.Side(style='thin', color='D9D9D9'),
                    bottom=opts.Side(style='thin', color='D9D9D9')
                )
                for cell in ws[1]:
                    cell.fill = fill
                    cell.font = font
                    cell.alignment = opts.Alignment(horizontal="center", vertical="center")
                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        cell.border = border
                        if cell.column in [1, 4, 8]:
                            cell.alignment = opts.Alignment(horizontal="center")
                        else:
                            cell.alignment = opts.Alignment(horizontal="left")

        st.download_button(
            label="📥 Descargar Reporte Excel Ejecutivo (.xlsx)",
            data=output.getvalue(),
            file_name="Reporte_Cliente_ALESE_Ejecutivo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
