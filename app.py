import streamlit as st
import pandas as pd
from docx import Document
import io
from datetime import datetime
import os

st.set_page_config(page_title="Generador de Llamados de Atención", page_icon="📄", layout="centered")

st.title("📄 Generador de Llamados de Atención Disciplinario")
st.write("Sube el reporte del SENA, escribe tus datos como instructor y genera el documento al instante.")

# 1. Campo para ingresar el Nombre del Instructor
nombre_instructor = st.text_input("Nombre del Instructor:", placeholder="Ej: William Hurtado")

# 2. Cargar la Plantilla de Word (Opcional)
template_file = st.file_uploader("1. Sube tu plantilla de Word (.docx) - Opcional", type=["docx"])

# 3. Cargar el archivo de Excel de Sofia Plus
excel_file = st.file_uploader("2. Sube tu reporte en Excel (.xlsx / .xls)", type=["xlsx", "xls"])

if excel_file is not None and nombre_instructor:
    try:
        # Definir la fuente de la plantilla
        doc_source = None
        if template_file is not None:
            doc_source = template_file
        elif os.path.exists("Llamado de atencion.docx"):
            doc_source = "Llamado de atencion.docx"
        elif os.path.exists("plantilla.docx"):
            doc_source = "plantilla.docx"
            
        if doc_source is None:
            st.error("⚠️ No se encontró la plantilla de Word. Por favor sube el archivo .docx arriba.")
        else:
            # Extraer automáticamente Ficha y Programa del reporte de Sofia Plus
            df_encabezado = pd.read_excel(excel_file, nrows=12, header=None)
            
            ficha = str(df_encabezado.iloc[2, 2]).strip() if len(df_encabezado) > 2 else ""
            programa = str(df_encabezado.iloc[5, 2]).strip() if len(df_encabezado) > 5 else ""

            # Leer la tabla de aprendices saltando las 12 filas del reporte
            df = pd.read_excel(excel_file, skiprows=12)
            st.success("¡Reporte del SENA cargado correctamente!")
            
            # Limpiar nombres de columnas
            df.columns = [str(c).strip() for c in df.columns]
            
            # Identificar columnas del estándar SENA
            col_nombre = "Nombre" if "Nombre" in df.columns else df.columns[2]
            col_apellido = "Apellidos" if "Apellidos" in df.columns else df.columns[3]
            col_cedula = "Número de" if "Número de" in df.columns else df.columns[1]

            # Combinar Nombre + Apellidos
            df['NOMBRE_COMPLETO'] = df[col_nombre].astype(str) + " " + df[col_apellido].astype(str)
            
            # Lista desplegable de aprendices
            aprendiz_seleccionado = st.selectbox("3. Selecciona el Aprendiz:", df['NOMBRE_COMPLETO'].unique())
            
            if aprendiz_seleccionado:
                fila_datos = df[df['NOMBRE_COMPLETO'] == aprendiz_seleccionado].iloc[0]
                cedula_aprendiz = str(fila_datos[col_cedula]).strip()
                
                # Fecha actual automática
                fecha_hoy = datetime.now().strftime("%d/%m/%Y")

                with st.expander("Ver datos que se reemplazarán en el documento"):
                    st.write(f"**Fecha del Llamado:** {fecha_hoy}")
                    st.write(f"**Instructor:** {nombre_instructor}")
                    st.write(f"**Aprendiz:** {aprendiz_seleccionado} - C.C. {cedula_aprendiz}")
                    st.write(f"**Programa:** {programa}")
                    st.write(f"**Ficha:** {ficha}")

                # Botón de Generación
                if st.button("🚀 Generar Llamado de Atención"):
                    if isinstance(doc_source, str):
                        doc = Document(doc_source)
                    else:
                        doc = Document(io.BytesIO(doc_source.getvalue()))
                    
                    reemplazos = {
                        "{{NOMBRE DEL PROGRAMA}}": programa,
                        "{{FECHA}}": fecha_hoy,
                        "{{NOMBRES Y CEDULA DEL APRENDIZ}}": f"{aprendiz_seleccionado} - C.C. {cedula_aprendiz}",
                        "{{FICHA}}": ficha,
                        "{{NOMBRE DEL INSTRUCTOR}}": nombre_instructor
                    }
                    
                    # Reemplazo en párrafos
                    for p in doc.paragraphs:
                        for tag, valor in reemplazos.items():
                            if tag in p.text:
                                p.text = p.text.replace(tag, valor)
                                
                    # Reemplazo en tablas
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
                    st.success("¡Documento generado exitosamente!")

    except Exception as e:
        st.error(f"Error procesando el archivo: {e}")
elif excel_file is not None and not nombre_instructor:
    st.warning("⚠️ Escribe tu nombre como instructor para continuar.")
