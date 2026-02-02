import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime
import time

st.set_page_config(page_title="Stock Lautaro App", layout="wide")

# --- 1. GESTIÓN DE SESIÓN Y LOGIN ---
if 'admin_logged_in' not in st.session_state:
    st.session_state.admin_logged_in = False

def check_admin_password():
    """Verifica la contraseña contra los secrets de Streamlit"""
    # Asegúrate de configurar [admin] password = "..." en tus secrets
    if st.session_state.password_input == st.secrets["admin"]["password"]:
        st.session_state.admin_logged_in = True
        st.session_state.password_input = "" 
    else:
        st.error("❌ Contraseña incorrecta")

def logout():
    st.session_state.admin_logged_in = False

# --- 2. FUNCIONES DE UTILERÍA ---
def parse_product_code(code_str):
    try:
        if pd.isna(code_str): return "S/D", "S/D", "S/D"
        parts = str(code_str).strip().split('.')
        if len(parts) < 2: return code_str, "N/A", "N/A"
        modelo = parts[0]
        suffix = parts[1]
        if len(suffix) >= 6:
            color = suffix[:3]
            talle = suffix[-3:]
        else:
            color = suffix
            talle = "N/A"
        return modelo, color, talle
    except:
        return str(code_str), "Err", "Err"

def color_rows(val):
    if '✅ OK' in str(val): return 'background-color: #d4edda; color: #155724' # Verde
    if '🔴 Falta' in str(val): return 'background-color: #f8d7da; color: #721c24' # Rojo
    if '🟣 Sobra' in str(val): return 'background-color: #e2d9f3; color: #4a148c' # Violeta
    return ''

# --- 3. CONEXIÓN A GOOGLE SHEETS ---
# ttl=0 asegura que siempre traiga los datos frescos de la nube al cargar
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    try:
        df = conn.read(worksheet="Hoja1", ttl=0)
        if 'Nombre_Lote' not in df.columns:
            return pd.DataFrame(columns=['Nombre_Lote', 'Estado_Stock', 'Código', 'Modelo_Ref', 'Color_Code', 'Talle_Code'])
        return df
    except:
        return pd.DataFrame(columns=['Nombre_Lote', 'Estado_Stock', 'Código'])

df_master = load_data()

# --- 4. BARRA LATERAL (MENU) ---
st.sidebar.title("Menú Principal")

opciones_menu = ["📂 Continuar / Editar Stock", "➕ Iniciar NUEVO Stock"]

if st.session_state.admin_logged_in:
    opciones_menu.append("🛑 Zona Admin (Borrar)")
    st.sidebar.success("🔓 Modo Admin: ON")
    if st.sidebar.button("Salir de Admin"):
        logout()
        st.rerun()
else:
    with st.sidebar.expander("🔐 Acceso Admin"):
        st.text_input("Contraseña:", type="password", key="password_input", on_change=check_admin_password)

modo = st.sidebar.radio("Ir a:", opciones_menu)

st.title("👕 Sistema de Control de Stock")

# ==========================================
# ZONA A: INICIAR NUEVO STOCK
# ==========================================
if modo == "➕ Iniciar NUEVO Stock":
    st.header("Cargar Nuevo Lote")
    st.markdown("Sube un archivo Excel/CSV para comenzar un conteo nuevo en un sector.")
    
    col1, col2 = st.columns(2)
    with col1:
        nuevo_lote = st.text_input("Nombre del Lote / Sector", placeholder="Ej: Deposito_Estanteria_1")
    with col2:
        archivo = st.file_uploader("Subir archivo", type=['csv', 'xlsx'])

    if st.button("🚀 Procesar y Crear", type="primary") and archivo and nuevo_lote:
        if nuevo_lote in df_master['Nombre_Lote'].unique():
            st.error(f"⚠️ Ya existe un lote llamado '{nuevo_lote}'. Usa otro nombre o ve a 'Continuar'.")
        else:
            try:
                # Lectura flexible
                try:
                    df_new = pd.read_csv(archivo, header=3, encoding='latin-1')
                    if 'Código' not in df_new.columns: raise Exception()
                except:
                    archivo.seek(0)
                    df_new = pd.read_csv(archivo)
                
                # Procesamiento
                df_new.columns = [c.strip() for c in df_new.columns]
                parsed = df_new['Código'].apply(parse_product_code)
                df_new['Modelo_Ref'] = [x[0] for x in parsed]
                df_new['Color_Code'] = [x[1] for x in parsed]
                df_new['Talle_Code'] = [x[2] for x in parsed]
                
                df_new['Nombre_Lote'] = nuevo_lote
                df_new['Estado_Stock'] = "Pendiente"
                df_new['Cant_Diferencia'] = 0
                df_new['Fecha_Inicio'] = datetime.now().strftime("%Y-%m-%d %H:%M")

                # Limpieza de columnas vacías
                df_new = df_new.dropna(axis=1, how='all')
                
                # Concatenar y guardar en la nube
                updated_master = pd.concat([df_master, df_new], ignore_index=True)
                conn.update(worksheet="Hoja1", data=updated_master)
                
                st.success(f"✅ Lote '{nuevo_lote}' creado correctamente.")
                st.balloons()
                time.sleep(2)
                st.rerun()
            except Exception as e:
                st.error(f"Error al procesar el archivo: {e}")

# ==========================================
# ZONA B: CONTINUAR / EDITAR (Aquí ocurre la magia)
# ==========================================
elif modo == "📂 Continuar / Editar Stock":
    lotes = [x for x in df_master['Nombre_Lote'].unique() if str(x) != 'nan']
    
    if not lotes:
        st.info("👋 No hay stocks activos. Ve a 'Iniciar NUEVO Stock' para comenzar.")
    else:
        col_sel, col_btn = st.columns([3, 1])
        with col_sel:
            seleccion = st.selectbox("Selecciona el Lote a trabajar:", lotes)
        with col_btn:
            st.write("") # Espaciado
            st.write("") 
            if st.button("🔄 Actualizar Datos"):
                st.cache_data.clear()
                st.rerun()
        
        # Filtramos datos del lote seleccionado
        # IMPORTANTE: Al filtrar así, df_lote es una copia. Al editarla no rompemos nada.
        df_lote = df_master[df_master['Nombre_Lote'] == seleccion].copy()
        
        # Orden visual
        if 'Modelo_Ref' in df_lote.columns:
            df_lote = df_lote.sort_values(by=['Modelo_Ref', 'Color_Code', 'Talle_Code'])

        # Métricas de avance
        total = len(df_lote)
        pendientes = len(df_lote[df_lote['Estado_Stock'] == 'Pendiente'])
        progreso = int(((total - pendientes) / total) * 100)
        st.progress(progreso)
        st.caption(f"Progreso: {progreso}% ({total - pendientes}/{total} items procesados)")

        # Configuración del Editor
        cols_cfg = {
            "Código": st.column_config.TextColumn("Código", disabled=True),
            "Descripción": st.column_config.TextColumn("Desc.", disabled=True),
            "Estado_Stock": st.column_config.SelectboxColumn(
                "Estado (Click para editar)", 
                options=["Pendiente", "✅ OK", "🔴 Falta", "🟣 Sobra"],
                required=True
            ),
            "Cant_Diferencia": st.column_config.NumberColumn("Cant. Dif.", min_value=0, step=1)
        }
        
        cols_show = ['Modelo_Ref', 'Color_Code', 'Talle_Code', 'Descripción', 'Estado_Stock', 'Cant_Diferencia']
        # Intersección segura de columnas
        cols_final = [c for c in cols_show if c in df_lote.columns]

        st.write("### 📝 Planilla de Edición")
        # El data_editor permite cambiar valores que YA estaban cargados antes
        edited_lote = st.data_editor(
            df_lote[cols_final],
            column_config=cols_cfg,
            use_container_width=True,
            height=500,
            hide_index=True,
            key=f"editor_{seleccion}" # Clave única para evitar conflictos
        )
        
        # BOTÓN DE GUARDADO
        if st.button("💾 Guardar Cambios en la Nube", type="primary"):
            # 1. Aplicamos los cambios del editor al dataframe local del lote
            # Esto actualiza las filas que María modificó (de Falta a OK, etc)
            df_lote.update(edited_lote)
            
            # 2. Reconstrucción del Master
            # Tomamos todo lo que NO es este lote del master original
            df_clean_master = df_master[df_master['Nombre_Lote'] != seleccion]
            
            # Le pegamos nuestro lote actualizado
            df_final_upload = pd.concat([df_clean_master, df_lote], ignore_index=True)
            
            # 3. Subimos a Google Sheets (Sobrescribe la hoja completa)
            conn.update(worksheet="Hoja1", data=df_final_upload)
            
            st.success("✅ ¡Cambios guardados! El Google Sheet ha sido actualizado.")
            time.sleep(1)
            st.rerun() # Recargamos para ver los datos frescos

        # RESUMEN DE DIFERENCIAS (Solo muestra lo que no está OK ni Pendiente)
        st.divider()
        st.write("### 🚨 Reporte de Diferencias")
        diferencias = df_lote[df_lote['Estado_Stock'].isin(['🔴 Falta', '🟣 Sobra'])]
        
        if not diferencias.empty:
            st.dataframe(
                diferencias[cols_final].style.map(color_rows, subset=['Estado_Stock']),
                use_container_width=True, 
                hide_index=True
            )
        else:
            st.info("No hay diferencias reportadas por el momento (Todo OK o Pendiente).")

# ==========================================
# ZONA C: ADMIN - BORRADO
# ==========================================
elif modo == "🛑 Zona Admin (Borrar)":
    st.header("⚠️ Zona de Peligro: Eliminar Stocks")
    
    lotes = [x for x in df_master['Nombre_Lote'].unique() if str(x) != 'nan']
    
    if not lotes:
        st.write("No hay stocks para borrar.")
    else:
        lote_borrar = st.selectbox("Selecciona el Lote a eliminar:", lotes)
        st.warning(f"Esto borrará todas las filas de '{lote_borrar}' del Google Sheet para siempre.")
        
        if st.button(f"🔥 ELIMINAR DEFINITIVAMENTE '{lote_borrar}'"):
            # Filtrar para dejar todo MENOS el lote seleccionado
            df_nuevo_master = df_master[df_master['Nombre_Lote'] != lote_borrar]
            
            # Actualizar Google Sheets
            conn.update(worksheet="Hoja1", data=df_nuevo_master)
            
            st.success("Lote eliminado.")
            time.sleep(1)
            st.rerun()
