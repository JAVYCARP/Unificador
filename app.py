import streamlit as st

# --- BLOQUEO DE CONTRASEÑA ---
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    st.subheader("🔒 Acceso Restringido")
    clave = st.text_input("Introduce la clave de acceso:", type="password")
    if st.button("Ingresar"):
        if clave == "MiClaveSecreta123": # <--- Pon la clave que quieras aquí
            st.session_state["autenticado"] = True
            st.rerun()
        else:
            st.error("❌ Clave incorrecta.")
    st.stop()
# -----------------------------

# (Aquí abajo dejas pegado el resto de tu código de unificación que ya te funcionaba)

import streamlit as st
import pandas as pd
import io

# CONFIGURACIÓN VISUAL DE LA PÁGINA
st.set_page_config(page_title="Unificador de Excel Protegido", page_icon="📊", layout="centered")

st.title("📊 Unificador de Archivos Excel")
st.write("Sube los archivos que deseas consolidar. El procesamiento se realiza de forma segura en el servidor.")

# Las pestañas que tu lógica busca dentro de los archivos
SOLAPAS_OBJETIVO = ['TdC', 'Filas de TdC', 'Materiales']

# 1. COMPONENTE WEB PARA SUBIR MÚLTIPLES ARCHIVOS A LA VEZ
archivos_subidos = st.file_uploader(
    "Arrastra o selecciona todos los archivos Excel aquí", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True  # Permite subir muchos archivos juntos
)

# Si el usuario ya subió al menos un archivo
if archivos_subidos:
    st.success(f"📂 Se han cargado {len(archivos_subidos)} archivos con éxito.")
    
    # 2. BOTÓN WEB PARA ENCIENDER TU LÓGICA DE NEGOCIO
    if st.button("🚀 Ejecutar Unificación Segura"):
        
        # Diccionario para agrupar los datos en memoria
        datos_agrupados = {solapa: [] for solapa in SOLAPAS_OBJETIVO}
        
        # Barra de progreso visual para el usuario web
        progreso = st.progress(0)
        status_text = st.empty()
        
        # Procesar cada archivo subido
        for index, archivo_objeto in enumerate(archivos_subidos):
            nombre_archivo = archivo_objeto.name
            status_text.text(f"Procesando: {nombre_archivo}...")
            
            try:
                # Leemos el archivo directamente desde la memoria del navegador
                with pd.ExcelFile(archivo_objeto) as excel_obj:
                    for solapa in SOLAPAS_OBJETIVO:
                        if solapa in excel_obj.sheet_names:
                            df = pd.read_excel(excel_obj, sheet_name=solapa)
                            
                            # TU CORRECCIÓN: Eliminar filas completamente vacías
                            df = df.dropna(how='all')
                            
                            if not df.empty:
                                df['Fuente_Archivo'] = nombre_archivo
                                datos_agrupados[solapa].append(df)
            except Exception as e:
                st.error(f"Error al leer {nombre_archivo}: {e}")
            
            # Actualizar barra de progreso en la pantalla del usuario
            progreso.progress((index + 1) / len(archivos_subidos))
            
        status_text.text("Generando archivo unificado final...")
        
        # Crear el archivo Excel en la memoria del servidor (sin escribir archivos en disco)
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            unificacion_exitosa = False
            
            for solapa, lista_dfs in datos_agrupados.items():
                if lista_dfs:
                    # TU CORRECCIÓN: Concatenar ignorando índices
                    df_final = pd.concat(lista_dfs, ignore_index=True)
                    
                    # TU CORRECCIÓN: Eliminar duplicados ignorando la columna origen
                    columnas_datos = [c for c in df_final.columns if c != 'Fuente_Archivo']
                    df_final = df_final.drop_duplicates(subset=columnas_datos, keep='first')
                    
                    # Escribir en el archivo temporal
                    df_final.to_excel(writer, sheet_name=solapa, index=False)
                    st.toast(f"✓ {solapa}: Unificada con éxito sin duplicados.", icon="✅")
                    unificacion_exitosa = True
                else:
                    st.warning(f"⚠ Pestaña '{solapa}': No se encontraron datos en los archivos subidos.")
        
        # Preparar los datos binarios para la descarga
        data_final = output.getvalue()
        
        if unificacion_exitosa:
            st.success("✨ ¡Unificación completada con éxito!")
            
            # 3. BOTÓN DE DESCARGA SEGURO
            st.download_button(
                label="📥 Descargar Excel Unificado",
                data=data_final,
                file_name="unificado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No se pudo unificar ninguna pestaña. Verifica que los archivos tengan los nombres de solapas correctos.")
