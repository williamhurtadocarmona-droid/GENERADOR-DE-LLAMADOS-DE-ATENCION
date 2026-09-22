import streamlit as st
import pandas as pd
from docx import Document
import io
from datetime import datetime

st.set_page_config(page_title="Generador de Llamados de Atención", page_icon="📄", layout="centered")

st.title("📄 Generador de Llamados de Atención Disciplinario")
st.write("Sube el reporte de Sofia Plus, escribe tus datos como instructor y genera el documento al instante.")

# 1. Campo para ingresar el Nombre del Instructor
nombre_instructor = st.text_input("Nombre del Instructor:", placeholder="Ej: William Hurtado")

# 2. Cargar el archivo de Excel
excel_file = st.file_uploader("Sube tu reporte en Excel (.xlsx / .xls)", type=["xlsx", "xls"])

if excel_file is not None and nombre_instructor:
    try:
        # Extraer automáticamente la Ficha y el Programa del encabezado fijo de Sofia Plus
        df_encabezado = pd.read_excel(excel_file, nrows=12, header=None)
        
        ficha = str(df_encabezado.iloc[2, 2]).strip() if len(df_encabezado) > 2 else ""
        programa = str(df_encabezado.iloc[5, 2]).strip() if len(df_encabezado) > 5 else ""

        # Leer la tabla de aprendices saltando las 12 filas iniciales del reporte
        df = pd.read_excel(excel_file, skiprows=12)
        st.success("¡Reporte del SENA cargado correctamente!")
        
        # Limpiar nombres de columnas
        df.columns = [str(c).strip() for c in df.columns]
        
        # Identificar automáticamente las columnas del estándar del SENA
        col_nombre = "Nombre" if "Nombre" in df.columns else df.columns[2]
        col_apellido = "Apellidos" if "Apellidos" in df.columns else df.columns[3]
        col_cedula = "Número de" if "Número de" in df.columns else df.columns[1]

        # Combinar Nombre + Apellidos
        df['NOMBRE_COMPLETO'] = df[col_nombre].astype(str) + " " + df[col_apellido].astype(str)
        
        # Lista desplegable con los aprendices
        aprendiz_seleccionado = st.selectbox("Selecciona el Aprendiz:", df['NOMBRE_COMPLETO'].unique())
        
        if aprendiz_seleccionado:
            fila_datos = df[df['NOMBRE_COMPLETO'] == aprendiz_seleccionado].iloc[0]
            cedula_aprendiz = str(fila_datos[col_cedula]).strip()
            
            # Fecha de hoy automática en formato Día/Mes/Año (ej: 22/09/2026)
            fecha_hoy = datetime.now().strftime("%d/%m/%Y")

            with st.expander("Ver datos que se insertarán en la plantilla"):
                st.write(f"**Fecha del Llamado:** {fecha_hoy}")
                st.write(f"**Instructor:** {nombre_instructor}")
                st.write(f"**Aprendiz:** {aprendiz_seleccionado} - C.C. {cedula_aprendiz}")
                st.write(f"**Programa:** {programa}")
                st.write(f"**Ficha:** {ficha}")

            # Botón para generar el Word
            if st.button("🚀 Generar Llamado de Atención"):
                doc = Document("plantilla.docx")
                
                # Mapeo exacto de tus etiquetas en la plantilla de Word
                reemplazos = {
                    "{{NOMBRE DEL PROGRAMA}}": programa,
                    "{{FECHA}}": fecha_hoy,
                    "{{NOMBRES Y CEDULA DEL APRENDIZ}}": f"{aprendiz_seleccionado} - C.C. {cedula_aprendiz}",
                    "{{FICHA}}": ficha,
                    "{{NOMBRE DEL INSTRUCTOR}}": nombre_instructor
                }
                
                # Reemplazo en los párrafos
                for p in doc.paragraphs:
                    for tag, valor in reemplazos.items():
                        if tag in p.text:
                            p.text = p.text.replace(tag, valor)
                            
                # Reemplazo en las tablas del Word
                for tabla in doc.tables:
                    for fila_tabla in tabla.rows:
                        for celda in fila_tabla.cells:
                            for p in celda.paragraphs:
                                for tag, valor in reemplazos.items():
                                    if tag in p.text:
                                        p.text = p.text.replace(tag, valor)
                
                output = io.BytesIO()
                doc.save(output)
                output.seek(0)
                
                st.download_button(
                    label="📥 Descargar Documento Listo",
                    data=output,
                    file_name=f"Llamado_Atencion_{aprendiz_seleccionado}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                st.success("¡Documento generado con la fecha de hoy!")

    except Exception as e:
        st.error(f"Error al leer el reporte: {e}")
elif excel_file is not None and not nombre_instructor:
    st.warning("⚠️ Por favor escribe tu nombre como instructor en la casilla superior.")
