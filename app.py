import streamlit as st
import pandas as pd
from docx import Document
import io
from datetime import datetime
import os

st.set_page_config(page_title="Generador de Llamados de Atención - SENA", page_icon="📄", layout="centered")

st.title("📄 Generador de Llamados de Atención SENA")
st.write("Sube tu reporte de Sofia Plus, selecciona el tipo de llamado, completa los datos y genera el documento al instante.")

# 1. Selección del tipo de llamado
tipo_llamado = st.selectbox(
    "Selecciona el Tipo de Llamado de Atención:",
    ["Primer Llamado de Atención Académico", "Segundo Llamado de Atención Académico/Disciplinario"]
)

# 2. Campo para ingresar el Nombre del Instructor y Área
col1, col2 = st.columns(2)
with col1:
    nombre_instructor = st.text_input("Nombre del Instructor:", placeholder="Ej: William Hurtado")
with col2:
    area_formacion = st.text_input("Área de Formación / Programa:", value="Área Metalmecánica")

# 3. Carga de la Plantilla de Word específica (Opcional)
st.markdown("---")
st.subheader("📁 Plantilla y Reporte")
template_file = st.file_uploader("Sube tu plantilla de Word (.docx) personalizada - Opcional", type=["docx"])

# 4. Cargar el archivo de Excel de Sofia Plus
excel_file = st.file_uploader("Sube tu reporte en Excel de Sofia Plus (.xlsx / .xls)", type=["xlsx", "xls"])

if excel_file is not None and nombre_instructor:
    try:
        # Definir la fuente de la plantilla según la elección o archivos locales
        doc_source = None
        if template_file is not None:
            doc_source = template_file
        else:
            if "Primer" in tipo_llamado:
                if os.path.exists("1er LLAMADO DE ATENCION ACADEMICO.docx"):
                    doc_source = "1er LLAMADO DE ATENCION ACADEMICO.docx"
                elif os.path.exists("plantilla.docx"):
                    doc_source = "plantilla.docx"
            else:
                if os.path.exists("2do LLAMADO DE ATENCION.docx"):
                    doc_source = "2do LLAMADO DE ATENCION.docx"
                elif os.path.exists("plantilla.docx"):
                    doc_source = "plantilla.docx"
                    
        # Si aún no hay fuente, buscar cualquier docx en el directorio
        if doc_source is None:
            archivos_docx = [f for f in os.listdir() if f.endswith(".docx")]
            if archivos_docx:
                doc_source = archivos_docx[0]

        if doc_source is None:
            st.error("⚠️ No se encontró ninguna plantilla de Word. Por favor sube un archivo .docx arriba.")
        else:
            # Extraer automáticamente Ficha y Programa del reporte de Sofia Plus
            df_encabezado = pd.read_excel(excel_file, nrows=12, header=None)
            
            ficha = str(df_encabezado.iloc[2, 2]).strip() if len(df_encabezado) > 2 and pd.notna(df_encabezado.iloc[2, 2]) else "N/A"
            programa = str(df_encabezado.iloc[5, 2]).strip() if len(df_encabezado) > 5 and pd.notna(df_encabezado.iloc[5, 2]) else "N/A"

            # Leer la tabla de aprendices saltando las 12 filas del reporte
            df = pd.read_excel(excel_file, skiprows=12)
            st.success("¡Reporte de Sofia Plus cargado correctamente!")
            
            # Limpiar nombres de columnas
            df.columns = [str(c).strip() for c in df.columns]
            
            # Detección de columnas
            col_tipo_doc = next((c for c in df.columns if "tipo" in c.lower()), df.columns[0] if len(df.columns) > 0 else "")
            col_cedula = next((c for c in df.columns if any(k in c.lower() for k in ["número", "documento", "identificación", "cedula"]) and c != col_tipo_doc), df.columns[1] if len(df.columns) > 1 else df.columns[0])
            col_nombre = next((c for c in df.columns if "nombre" in c.lower() and "apellido" not in c.lower()), df.columns[2] if len(df.columns) > 2 else df.columns[0])
            col_apellido = next((c for c in df.columns if "apellido" in c.lower()), df.columns[3] if len(df.columns) > 3 else df.columns[0])

            # Combinar Nombre + Apellidos
            df['NOMBRE_COMPLETO'] = df[col_nombre].astype(str) + " " + df[col_apellido].astype(str)
            
            st.markdown("---")
            aprendiz_seleccionado = st.selectbox("👤 Selecciona el Aprendiz:", df['NOMBRE_COMPLETO'].unique())
            
            if aprendiz_seleccionado:
                fila_datos = df[df['NOMBRE_COMPLETO'] == aprendiz_seleccionado].iloc[0]
                cedula_aprendiz = str(fila_datos[col_cedula]).strip()
                tipo_doc_aprendiz = str(fila_datos[col_tipo_doc]).strip() if col_tipo_doc in df.columns else "C.C."
                
                st.markdown("### 📝 Descripción de la Evidencia o Hechos")
                texto_evidencia = st.text_area(
                    "Escribe aquí la descripción de la falta o lo ocurrido para la tabla:",
                    value="Informe del instructor y registro de observación directa sobre el uso inadecuado de un dispositivo de sonido (parlante) a alto volumen dentro del ambiente de formación (aula/taller), alterando la convivencia y el desarrollo normal de las actividades académicas.",
                    height=120
                )

                # Datos adicionales si es segundo llamado
                fecha_primer_llamado = ""
                if "Segundo" in tipo_llamado:
                    st.markdown("### 📋 Datos adicionales para el Segundo Llamado")
                    fecha_primer_llamado = st.text_input("Fecha del Primer Llamado de Atención:", value="DD/MM/AAAA")

                fecha_hoy = datetime.now().strftime("%d/%m/%Y")

                with st.expander("🔍 Ver datos que se reemplazarán en el documento"):
                    st.write(f"**Tipo de Llamado:** {tipo_llamado}")
                    st.write(f"**Fecha Actual:** {fecha_hoy}")
                    st.write(f"**Instructor:** {nombre_instructor} ({area_formacion})")
                    st.write(f"**Aprendiz:** {aprendiz_seleccionado} ({tipo_doc_aprendiz} {cedula_aprendiz})")
                    st.write(f"**Programa:** {programa}")
                    st.write(f"**Ficha:** {ficha}")
                    st.write(f"**Evidencia:** {texto_evidencia}")
                    if "Segundo" in tipo_llamado:
                        st.write(f"**Fecha 1er Llamado:** {fecha_primer_llamado}")

                st.markdown("---")
                if st.button("🚀 Generar y Descargar Documento"):
                    if isinstance(doc_source, str):
                        doc = Document(doc_source)
                    else:
                        doc = Document(io.BytesIO(doc_source.getvalue()))
                    
                    reemplazos = {
                        "{{PROGRAMA}}": programa,
                        "{{NOMBRE DEL PROGRAMA}}": programa,
                        "{{CIUDAD}}": "Barranquilla",
                        "{{FECHA}}": fecha_hoy,
                        "{{NOMBRE DEL APRENDIZ}}": aprendiz_seleccionado,
                        "{{TIPO DE DOCUEMNTO}}": tipo_doc_aprendiz,
                        "{{TIPO DE DOCUMENTO}}": tipo_doc_aprendiz,
                        "{{NUMERO}}": cedula_aprendiz,
                        "{{NUMERO DE FICHA}}": ficha,
                        "{{FICHA}}": ficha,
                        "{{Nombre del instructor}}": nombre_instructor,
                        "{{NOMBRE DEL INSTRUCTOR}}": nombre_instructor,
                        "{{FECHA DEL PRIMER LLAMADO}}": fecha_primer_llamado,
                        "{{DESCRIPCION DE LOS HECHOS}}": texto_evidencia
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
                                    if "Informe del instructor y registro de observación directa" in p.text and "{{DESCRIPCION DE LOS HECHOS}}" not in p.text:
                                        p.text = texto_evidencia
                    
                    output = io.BytesIO()
                    doc.save(output)
                    output.seek(0)
                    
                    nombre_archivo_salida = f"Llamado_Atencion_{aprendiz_seleccionado.replace(' ', '_')}.docx"
                    
                    st.download_button(
                        label="📥 Descargar Documento Listo (.docx)",
                        data=output,
                        file_name=nombre_archivo_salida,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    st.success("¡Documento generado exitosamente con los datos del reporte!")

    except Exception as e:
        st.error(f"Error procesando el archivo: {e}")
elif excel_file is not None and not nombre_instructor:
    st.warning("⚠️ Escribe tu nombre como instructor para continuar.")
