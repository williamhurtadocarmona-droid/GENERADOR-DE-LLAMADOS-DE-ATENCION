import streamlit as st
import pandas as pd
from docx import Document
import io
from fpdf import FPDF

# Configuración de la página
st.set_page_config(page_title="Generador de Llamados de Atención", page_icon="📄", layout="centered")

st.title("📄 Generador de Llamados de Atención")
st.write("Sube tu plantilla de Word y el archivo de Excel para generar los documentos automáticamente en Word o PDF.")

# 1. Subir la plantilla de Word
plantilla_file = st.file_uploader("1. Sube tu plantilla de Word (.docx)", type=["docx"])

# 2. Subir el archivo de Excel
excel_file = st.file_uploader("2. Sube el reporte de Excel (.xlsx)", type=["xlsx"])

# Función para convertir el contenido extraído de Word a PDF usando FPDF
def convertir_docx_a_pdf(doc_bytes):
    doc = Document(io.BytesIO(doc_bytes))
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=11)
    
    for p in doc.paragraphs:
        texto = p.text.strip()
        if texto:
            # Codificar texto para evitar problemas con caracteres especiales en FPDF
            texto_limpio = texto.encode('latin-1', 'replace').decode('latin-1')
            pdf.multi_cell(0, 8, txt=texto_limpio)
            pdf.ln(2)
            
    for table in doc.tables:
        for row in table.rows:
            fila_texto = " | ".join([cell.text.strip() for cell in row.cells])
            if fila_texto.strip():
                fila_limpia = fila_texto.encode('latin-1', 'replace').decode('latin-1')
                pdf.multi_cell(0, 7, txt=fila_limpia)
        pdf.ln(4)
        
    return pdf.output(dest='S').encode('latin-1')

if excel_file:
    try:
        df = pd.read_excel(excel_file)
        st.success("¡Archivo de Excel cargado correctamente!")
        
        # Selección de aprendiz
        if 'NOMBRE' in df.columns or 'APRENDIZ' in df.columns:
            col_nombre = 'NOMBRE' if 'NOMBRE' in df.columns else 'APRENDIZ'
            aprendiz_seleccionado = st.selectbox("Selecciona el aprendiz:", df[col_nombre].dropna().unique())
            
            fila_datos = df[df[col_nombre] == aprendiz_seleccionado].iloc[0]
            
            if st.button("🚀 Generar Documentos"):
                # Abrir plantilla subida o la guardada por defecto
                if plantilla_file:
                    doc = Document(plantilla_file)
                else:
                    doc = Document("Llamado de atencion.docx")
                
                # Reemplazo de etiquetas
                datos_reemplazo = {f"{{{{{col}}}}}": str(val) for col, val in fila_datos.items()}
                
                for p in doc.paragraphs:
                    for tag, val in datos_reemplazo.items():
                        if tag in p.text:
                            p.text = p.text.replace(tag, val)
                            
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for p in cell.paragraphs:
                                for tag, val in datos_reemplazo.items():
                                    if tag in p.text:
                                        p.text = p.text.replace(tag, val)
                
                # Guardar resultado en memoria (Word)
                buffer_word = io.BytesIO()
                doc.save(buffer_word)
                buffer_word.seek(0)
                word_bytes = buffer_word.getvalue()
                
                st.subheader("🎉 Documento generado con éxito")
                
                col1, col2 = st.columns(2)
                
                # Botón Descargar Word
                with col1:
                    st.download_button(
                        label="📥 Descargar Word (.docx)",
                        data=word_bytes,
                        file_name=f"Llamado_Atencion_{aprendiz_seleccionado}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                
                # Botón Descargar PDF
                with col2:
                    try:
                        pdf_bytes = convertir_docx_a_pdf(word_bytes)
                        st.download_button(
                            label="📕 Descargar PDF (.pdf)",
                            data=pdf_bytes,
                            file_name=f"Llamado_Atencion_{aprendiz_seleccionado}.pdf",
                            mime="application/pdf"
                        )
                    except Exception as e:
                        st.warning("No se pudo generar el botón de PDF, pero el archivo de Word está listo.")
                        
        else:
            st.error("No se encontró una columna llamada 'NOMBRE' o 'APRENDIZ' en el Excel.")
    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
