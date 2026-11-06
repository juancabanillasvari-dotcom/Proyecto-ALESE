import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Seguimiento de Unidades - ALESE", layout="wide")

st.title("🚗 Control de Estado de Unidades — Cliente ALESE")
st.markdown("Sube tu reporte nativo de Excel para obtener el estado clasificado, correlativo y ordenado estrictamente de más antiguo a más actual.")

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
        
        # Conversión estricta a datetime para orden cronológico de más antiguo a más actual
        piv['Fecha_Datetime'] = pd.to_datetime(piv['Fecha_Raw'], errors='coerce', dayfirst=True)

        # Extraer los 10 primeros caracteres del archivo nativo
        def extraer_10_chars(row):
            val = row['Fecha_Raw']
            if pd.isna(val) or str(val).strip() in ["", "nan", "NaT", "None"]:
                return ""
            return str(val).strip()[:10]

        piv['Fecha_Str'] = piv.apply(extraer_10_chars, axis=1)

        # Ordenar estrictamente de menor a mayor (más antiguo arriba)
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

        # Preparar DataFrames limpios con numeración correlativa N°
        def build_clean_df(df_sub, col_fecha_title):
            if df_sub.empty:
                return pd.DataFrame(columns=['N°', 'Cliente', 'Almacen', 'Vin', 'Marca', 'Modelo', 'Color Vehiculo', col_fecha_title])
            res = pd.DataFrame({
                'N°': range(1, len(df_sub) + 1),
                'Cliente': df_sub['Cliente'],
                'Almacen': df_sub['Almacen'],
                'Vin': df_sub['Vin'],
                'Marca': df_sub['Marca'],
                'Modelo': df_sub['Modelo'],
                'Color Vehiculo': df_sub['Color Vehiculo'],
                col_fecha_title: df_sub['Fecha_Str']
            })
            return res

        df_fin = build_clean_df(piv[piv['Estado Vehiculo'] == "FINALIZADO"], "Fecha Término")
        df_proc = build_clean_df(piv[piv['Estado Vehiculo'] == "EN PROCESO"], "Fecha Ingreso")
        df_cola = build_clean_df(piv[piv['Estado Vehiculo'] == "EN COLA"], "Fecha Ingreso")

        # Visualización nativa impecable en Streamlit
        st.subheader("🟢 Unidades Finalizadas")
        st.dataframe(df_fin, use_container_width=True, hide_index=True)

        st.subheader("🟡 Unidades En Proceso")
        st.dataframe(df_proc, use_container_width=True, hide_index=True)

        st.subheader("🔴 Unidades En Cola")
        st.dataframe(df_cola, use_container_width=True, hide_index=True)

        # Generar Excel estructurado con 3 pestañas profesionales
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_fin.to_excel(writer, index=False, sheet_name='Finalizado')
            df_proc.to_excel(writer, index=False, sheet_name='En Proceso')
            df_cola.to_excel(writer, index=False, sheet_name='En Cola')

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
