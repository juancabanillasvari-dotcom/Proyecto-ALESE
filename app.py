import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Seguimiento de Unidades - ALESE", layout="wide")

st.title("🚗 Control de Estado de Unidades — Cliente ALESE")
st.markdown("Sube tu reporte nativo de Excel para obtener el estado clasificado y correlativo de las unidades en almacén.")

uploaded_file = st.file_uploader("Sube tu archivo Excel (ej. Automate_diario.xlsx)", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        sheet_target = 'SeguimientoServicio' if 'SeguimientoServicio' in pd.ExcelFile(uploaded_file).sheet_names else 0
        df = pd.read_excel(uploaded_file, sheet_name=sheet_target)

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

        # Tomar exactamente los 10 primeros caracteres de la cadena nativa original
        def extraer_10_chars(row):
            estado = row['Estado Vehiculo']
            if estado in ["EN COLA", "EN PROCESO"]:
                val = row.get('Fecha Ingreso_PDI INICIAL COMPLETO', row.get('Fecha Ingreso_PDI FINAL', ''))
            else:
                val = row.get('Fecha Termino_PDI FINAL', row.get('Fecha Termino_PDI INICIAL COMPLETO', ''))
            
            if pd.isna(val) or str(val).strip() in ["", "nan", "NaT", "None"]:
                return ""
            return str(val).strip()[:10]

        piv['Fecha'] = piv.apply(extraer_10_chars, axis=1)

        estado_order = {"FINALIZADO": 1, "EN PROCESO": 2, "EN COLA": 3}
        piv['Orden_Estado'] = piv['Estado Vehiculo'].map(estado_order)
        piv = piv.sort_values(by=['Orden_Estado', 'Fecha'], ascending=[True, True])

        cols_export = ['Cliente', 'Almacen', 'Vin', 'Marca', 'Modelo', 'Color Vehiculo', 'Estado Vehiculo', 'Fecha']
        df_final = piv[cols_export]

        # Métricas principales
        st.markdown("---")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("📦 Total en Almacén", len(df_final))
        m2.metric("🟢 Finalizadas", (df_final['Estado Vehiculo'] == "FINALIZADO").sum())
        m3.metric("🟡 En Proceso", (df_final['Estado Vehiculo'] == "EN PROCESO").sum())
        m4.metric("🔴 En Cola", (df_final['Estado Vehiculo'] == "EN COLA").sum())
        st.markdown("---")

        # Preparar dataframes con numeración correlativa N°
        df_fin = df_final[df_final['Estado Vehiculo'] == "FINALIZADO"].drop(columns=['Estado Vehiculo']).copy()
        df_fin.insert(0, 'N°', range(1, len(df_fin) + 1))

        df_proc = df_final[df_final['Estado Vehiculo'] == "EN PROCESO"].drop(columns=['Estado Vehiculo']).copy()
        df_proc.insert(0, 'N°', range(1, len(df_proc) + 1))

        df_cola = df_final[df_final['Estado Vehiculo'] == "EN COLA"].drop(columns=['Estado Vehiculo']).copy()
        df_cola.insert(0, 'N°', range(1, len(df_cola) + 1))

        st.subheader("🟢 Unidades Finalizadas")
        st.dataframe(df_fin, use_container_width=True)

        st.subheader("🟡 Unidades En Proceso")
        st.dataframe(df_proc, use_container_width=True)

        st.subheader("🔴 Unidades En Cola")
        st.dataframe(df_cola, use_container_width=True)

        # Generar Excel estructurado con 3 pestañas para descarga
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_fin.to_excel(writer, index=False, sheet_name='Finalizado')
            df_proc.to_excel(writer, index=False, sheet_name='En Proceso')
            df_cola.to_excel(writer, index=False, sheet_name='En Cola')

        st.download_button(
            label="📥 Descargar Reporte Excel Clasificado (.xlsx)",
            data=output.getvalue(),
            file_name="Reporte_Cliente_ALESE_Correlativo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
