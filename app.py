import streamlit as st
import pandas as pd
import unicodedata
import io
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------
# 1. CONFIGURACIÓN DE PÁGINA Y ESTILOS
# ---------------------------------------------------------
st.set_page_config(
    page_title="Portal de Coordinación Electoral",
    page_icon="🗳️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main { background-color: #f8fafc; }
    .login-header { text-align: center; padding: 2rem 0; }
    .login-header h1 { color: #1e293b; font-weight: 700; margin-bottom: 0.5rem; }
    .login-header p { color: #64748b; font-size: 1rem; }
    .metric-card {
        background-color: #ffffff;
        border-radius: 12px;
        padding: 1.25rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        border: 1px solid #e2e8f0;
        text-align: center;
    }
    .metric-card h3 { color: #64748b; font-size: 0.875rem; font-weight: 600; text-transform: uppercase; margin-bottom: 0.5rem; }
    .metric-card .number { color: #0f172a; font-size: 2rem; font-weight: 700; }
    .user-banner {
        background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
        color: white;
        padding: 1.5rem 2rem;
        border-radius: 12px;
        margin-bottom: 2rem;
        box-shadow: 0 10px 15px -3px rgba(59, 130, 246, 0.2);
    }
    .user-banner h2 { color: white !important; margin: 0; font-weight: 700; }
    .user-banner p { color: #dbeafe; margin-top: 0.25rem; margin-bottom: 0; }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 2. FUNCIONES DE NORMALIZACIÓN AUTOMÁTICA DE COLUMNAS
# ---------------------------------------------------------
def normalizar_texto(texto):
    """Quita tildes, espacios extremos y pasa a mayúsculas."""
    if pd.isna(texto):
        return ""
    s = str(texto).strip()
    s = unicodedata.normalize('NFKD', s).encode('ASCII', 'ignore').decode('utf-8')
    return s.upper()

def estandarizar_columnas(df):
    """Detecta automáticamente variaciones de nombres de columnas y las renombra."""
    column_mapping = {}
    for col in df.columns:
        col_norm = normalizar_texto(col)
        
        if "CEDULA" in col_norm or "NRO" in col_norm or "IDENTIFI" in col_norm:
            column_mapping[col] = "CEDULA"
        elif "NOMBRE" in col_norm or "APELLIDO" in col_norm:
            column_mapping[col] = "APELLIDOS Y NOMBRES"
        elif "ASIGNADO" in col_norm:
            column_mapping[col] = "ASIGNADO"
        elif "CANTON" in col_norm:
            column_mapping[col] = "CANTON"
        elif "TEL" in col_norm or "CEL" in col_norm:
            column_mapping[col] = "TELEFONO"
        elif "OBSERVAC" in col_norm or "ERROR" in col_norm or "NOVEDAD" in col_norm:
            column_mapping[col] = "OBSERVACIONES"
        elif "TIPO" in col_norm or "CARGO" in col_norm:
            column_mapping[col] = "TIPO"
            
    return df.rename(columns=column_mapping)

def limpiar_cedula(val):
    """Limpia espacios, elimina decimales (.0) y preserva el 0 inicial si falta."""
    if pd.isna(val):
        return ""
    s = str(val).strip().split('.')[0]
    if len(s) == 9:
        s = s.zfill(10)
    return s

# ---------------------------------------------------------
# 3. CARGA DE DATOS EXCEL INICIALES
# ---------------------------------------------------------
@st.cache_data
def load_data():
    try:
        df_miembros = pd.read_excel("CANTONALES_1.xlsx", dtype=str, header=1)
        df_coord_recinto = pd.read_excel("COOR_RECINTO_LLAMAR_CAPACITAR.xlsx", dtype=str)
        df_coord_mesa = pd.read_excel("COOR_MESA_LLAMAR1_CAPACITAR.xlsx", dtype=str)

        df_miembros = estandarizar_columnas(df_miembros)
        df_coord_recinto = estandarizar_columnas(df_coord_recinto)
        df_coord_mesa = estandarizar_columnas(df_coord_mesa)

        tablas = [
            ("CANTONALES_1.xlsx", df_miembros),
            ("COOR_RECINTO_LLAMAR_CAPACITAR.xlsx", df_coord_recinto),
            ("COOR_MESA_LLAMAR1_CAPACITAR.xlsx", df_coord_mesa)
        ]

        for nombre_archivo, df in tablas:
            if "CEDULA" in df.columns:
                df["CEDULA"] = df["CEDULA"].apply(limpiar_cedula)
            else:
                st.error(
                    f"⚠️ En el archivo **{nombre_archivo}** no se encontró una columna válida de Cédula. "
                    f"Columnas detectadas en el Excel: `{list(df.columns)}`"
                )

        df_coord_recinto["TIPO"] = "Coordinador de Recinto"
        df_coord_mesa["TIPO"] = "Coordinador de Mesa"

        return df_miembros, df_coord_recinto, df_coord_mesa

    except Exception as e:
        st.error(f"❌ Error al abrir los archivos Excel: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

df_miembros, df_coord_recinto, df_coord_mesa = load_data()

# ---------------------------------------------------------
# 4. INICIALIZACIÓN DEL ESTADO DE SESIÓN
# ---------------------------------------------------------
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "user_info" not in st.session_state:
    st.session_state["user_info"] = None
if "observaciones_dict" not in st.session_state:
    st.session_state["observaciones_dict"] = {}
if "custom_coordinadores" not in st.session_state:
    st.session_state["custom_coordinadores"] = None

# ---------------------------------------------------------
# 5. VENTANAS EMERGENTES (MODALES)
# ---------------------------------------------------------
@st.dialog("📥 Importar Nómina")
def abrir_dialog_importar(df_actual, canton_asignado):
    st.write(f"Suba un archivo Excel (`.xlsx`, `.xls`) o CSV (`.csv`) perteneciente a la región **{canton_asignado}**.")
    
    archivo_cargado = st.file_uploader(
        "Seleccione el archivo a cargar:",
        type=["xlsx", "xls", "csv"],
        key="uploader_modal"
    )

    if archivo_cargado is not None:
        try:
            if archivo_cargado.name.endswith('.csv'):
                df_nuevo = pd.read_csv(archivo_cargado, dtype=str)
            else:
                df_nuevo = pd.read_excel(archivo_cargado, dtype=str)

            df_nuevo = estandarizar_columnas(df_nuevo)
            if "CEDULA" in df_nuevo.columns:
                df_nuevo["CEDULA"] = df_nuevo["CEDULA"].apply(limpiar_cedula)

            # ---------------------------------------------------------
            # VALIDACIÓN DE CANTÓN/REGIÓN PERMITIDA
            # ---------------------------------------------------------
            canton_user_norm = normalizar_texto(canton_asignado)

            if "CANTON" in df_nuevo.columns:
                cantones_presentes = df_nuevo["CANTON"].dropna().apply(normalizar_texto).unique()
                
                # Filtrar únicamente registros pertenecientes al cantón asignado
                df_nuevo["CANTON_TMP_NORM"] = df_nuevo["CANTON"].apply(normalizar_texto)
                registros_totales = len(df_nuevo)
                df_nuevo = df_nuevo[df_nuevo["CANTON_TMP_NORM"] == canton_user_norm].copy()
                df_nuevo.drop(columns=["CANTON_TMP_NORM"], inplace=True)
                registros_validos = len(df_nuevo)

                if registros_validos == 0:
                    st.error(
                        f"❌ **Acceso Restringido:** El archivo contiene datos de cantón(es) distinto(s) "
                        f"({', '.join(cantones_presentes)}) y ningún registro pertenece a su cantón asignado (**{canton_asignado}**)."
                    )
                    return
                elif registros_validos < registros_totales:
                    st.warning(
                        f"⚠️ **Filtrado Regional Aplicado:** Se ignoraron {registros_totales - registros_validos} registro(s) "
                        f"pertenecientes a otros cantones. Solo se procesarán los **{registros_validos}** registros de **{canton_asignado}**."
                    )
            else:
                df_nuevo["CANTON"] = canton_asignado
                st.info(f"ℹ️ El archivo cargado no especificaba cantón. Se le asignó automáticamente la región **{canton_asignado}**.")

            # Selección de modalidad de importación
            modo_importacion = st.radio(
                "Elija el modo de importación:",
                [
                    "🔄 Sobrescribir nómina existente",
                    "➕ Agregar solo columnas y datos modificados (con sufijo '(nuevo)' y 'NUEVO REGISTRO')"
                ],
                index=0
            )

            st.divider()

            if "Sobrescribir" in modo_importacion:
                st.write("🔍 **Previsualización de los datos a sobrescribir:**")
                st.dataframe(df_nuevo.head(5), use_container_width=True)

                st.warning("⚠️ **Confirmación requerida:** Esta acción reemplazará la información actual del cantón.")
                
                c_conf1, c_conf2 = st.columns(2)
                with c_conf1:
                    if st.button("✅ Confirmar y Sobrescribir", type="primary", use_container_width=True):
                        if "OBSERVACIONES" in df_nuevo.columns and "CEDULA" in df_nuevo.columns:
                            for _, row in df_nuevo.iterrows():
                                cedula = str(row["CEDULA"])
                                obs = str(row["OBSERVACIONES"]).strip() if pd.notna(row["OBSERVACIONES"]) else ""
                                if obs:
                                    st.session_state["observaciones_dict"][cedula] = obs

                        st.session_state["custom_coordinadores"] = df_nuevo
                        st.success(f"🎉 Nómina del cantón {canton_asignado} sobrescrita con éxito.")
                        st.rerun()
                with c_conf2:
                    if st.button("❌ Cancelar", use_container_width=True):
                        st.rerun()

            else:
                # ---------------------------------------------------------
                # FILTRADO EXCLUSIVO DE COLUMNAS QUE SUFRIERON MODIFICACIONES
                # ---------------------------------------------------------
                dict_actual = df_actual.set_index("CEDULA").to_dict(orient="index") if "CEDULA" in df_actual.columns else {}
                
                filas_modificadas = []
                for _, row in df_nuevo.iterrows():
                    cedula = row["CEDULA"]
                    if cedula not in dict_actual:
                        # Registro totalmente nuevo: agregar campos nuevos presentes
                        fila_dict = {"CEDULA": cedula, "NUEVO REGISTRO": "Sí"}
                        for col in row.index:
                            if col != "CEDULA":
                                val_nuevo = str(row[col]).strip() if pd.notna(row[col]) else ""
                                if val_nuevo:
                                    fila_dict[f"{col} (nuevo)"] = val_nuevo
                        filas_modificadas.append(fila_dict)
                    else:
                        # Registro existente: incluir ÚNICAMENTE las columnas que tuvieron cambios
                        actual_row = dict_actual[cedula]
                        fila_dict = {"CEDULA": cedula}
                        hay_cambio = False
                        for col in row.index:
                            if col != "CEDULA":
                                val_nuevo = str(row[col]).strip() if pd.notna(row[col]) else ""
                                val_actual = str(actual_row.get(col, "")).strip() if pd.notna(actual_row.get(col, "")) else ""
                                if val_nuevo != val_actual and val_nuevo != "":
                                    fila_dict[f"{col} (nuevo)"] = val_nuevo
                                    hay_cambio = True
                        if hay_cambio:
                            fila_dict["NUEVO REGISTRO"] = "Sí"
                            filas_modificadas.append(fila_dict)

                if len(filas_modificadas) == 0:
                    st.info("ℹ️ Toda la información del archivo importado ya existe idéntica en el sistema. No se detectaron modificaciones ni registros nuevos.")
                else:
                    df_nuevo_renamed = pd.DataFrame(filas_modificadas)

                    # Fusionar solo agregando las columnas que efectivamente experimentaron cambios
                    df_merged = pd.merge(df_actual, df_nuevo_renamed, on="CEDULA", how="outer")
                    df_merged["NUEVO REGISTRO"] = df_merged["NUEVO REGISTRO"].fillna("No")

                    st.write(f"🔍 **Previsualización ({len(filas_modificadas)} registro(s) con cambios o datos nuevos):**")
                    st.dataframe(df_merged.head(10), use_container_width=True)

                    c_conf1, c_conf2 = st.columns(2)
                    with c_conf1:
                        if st.button("✅ Confirmar y Agregar Información Nueva", type="primary", use_container_width=True):
                            st.session_state["custom_coordinadores"] = df_merged
                            st.success(f"🎉 Se incorporaron solo las columnas modificadas de {len(filas_modificadas)} registro(s).")
                            st.rerun()
                    with c_conf2:
                        if st.button("❌ Cancelar", use_container_width=True):
                            st.rerun()

        except Exception as e:
            st.error(f"❌ Error al procesar el archivo: {e}")

@st.dialog("📤 Exportar Nómina")
def abrir_dialog_exportar(df_exportar, canton_asignado):
    st.write("Seleccione el formato en el que desea descargar la nómina actual con las observaciones guardadas:")
    
    # ---------------------------------------------------------
    # GENERACIÓN DE EXCEL ESTILIZADO (OpenPyXL)
    # ---------------------------------------------------------
    buffer_excel = io.BytesIO()
    with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
        df_exportar.to_excel(writer, index=False, sheet_name='Coordinadores')
        
        ws = writer.sheets['Coordinadores']
        
        # Declaración de estilos (colores, bordes y fuentes)
        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        
        row_even_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        row_odd_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        
        thin_border = Border(
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'),
            bottom=Side(style='thin', color='CBD5E1')
        )
        
        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Encabezados
        ws.row_dimensions[1].height = 28
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = align_center
            cell.border = thin_border

        # Filas de datos
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column), start=2):
            ws.row_dimensions[row_idx].height = 24
            fill_actual = row_even_fill if row_idx % 2 == 0 else row_odd_fill
            
            for cell in row:
                cell.fill = fill_actual
                cell.border = thin_border
                cell.alignment = align_center
                cell.font = Font(name="Calibri", size=10)

        # Ajuste automático de ancho de columna
        for col in ws.columns:
            max_len = 0
            for cell in col:
                val = str(cell.value or '')
                if len(val) > max_len:
                    max_len = len(val)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 16)

    buffer_excel.seek(0)

    # Buffer CSV
    buffer_csv = df_exportar.to_csv(index=False).encode('utf-8-sig')

    col_e1, col_e2 = st.columns(2)
    with col_e1:
        st.download_button(
            label="📊 Descargar Excel (.xlsx)",
            data=buffer_excel,
            file_name=f"Nomina_Coordinadores_{canton_asignado}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col_e2:
        st.download_button(
            label="📄 Descargar CSV (.csv)",
            data=buffer_csv,
            file_name=f"Nomina_Coordinadores_{canton_asignado}.csv",
            mime="text/csv",
            use_container_width=True
        )

# ---------------------------------------------------------
# 6. LÓGICA DE ACCESO (LOGIN)
# ---------------------------------------------------------
def login(cedula, password):
    cedula_clean = limpiar_cedula(cedula)
    password_clean = str(password).strip()

    if password_clean != "cne.2026":
        st.error("❌ Contraseña incorrecta. Intente de nuevo.")
        return

    if "CEDULA" not in df_miembros.columns:
        st.error("❌ No fue posible procesar la columna de cédulas desde el Excel.")
        return

    match = df_miembros[df_miembros["CEDULA"] == cedula_clean]

    if not match.empty:
        st.session_state["authenticated"] = True
        st.session_state["user_info"] = match.iloc[0].to_dict()
        st.rerun()
    else:
        st.error(f"❌ La cédula '{cedula_clean}' no está registrada como Miembro Cantonal.")

def logout():
    st.session_state["authenticated"] = False
    st.session_state["user_info"] = None
    st.session_state["custom_coordinadores"] = None
    st.rerun()

# ---------------------------------------------------------
# 7. NAVEGACIÓN Y VISTAS
# ---------------------------------------------------------
if not st.session_state.get("authenticated", False):
    # --- VISTA DE LOGIN ---
    col1, col2, col3 = st.columns([1, 1.5, 1])

    with col2:
        st.markdown("""
        <div class="login-header">
            <h1>Sistema Electoral Cantonal</h1>
            <p>Acceso exclusivo para Miembros Cantonales</p>
        </div>
        """, unsafe_allow_html=True)

        with st.form("login_form", clear_on_submit=False):
            st.subheader("Iniciar Sesión")

            username_input = st.text_input(
                "Número de Cédula (Usuario)",
                placeholder="Ej. 1312345678"
            )
            
            password_input = st.text_input(
                "Contraseña",
                type="password",
                placeholder="Ingresar contraseña"
            )

            submit_btn = st.form_submit_button("Ingresar al Sistema", use_container_width=True)

            if submit_btn:
                if username_input and password_input:
                    login(username_input, password_input)
                else:
                    st.warning("⚠️ Por favor ingrese su número de cédula y contraseña.")

else:
    # --- PANEL CANTONAL ---
    user = st.session_state["user_info"]
    canton_asignado = user.get("ASIGNADO", user.get("CANTON", "NO ASIGNADO"))

    # Barra lateral
    with st.sidebar:
        st.title("🛡️ Panel Cantonal")
        st.markdown(f"**Usuario:** {user.get('APELLIDOS Y NOMBRES', 'N/D')}")
        st.markdown(f"**Cédula:** {user.get('CEDULA', 'N/D')}")
        st.markdown(f"**Región Asignada:** `{canton_asignado}`")
        st.divider()
        if st.button("🚪 Cerrar Sesión", use_container_width=True):
            logout()

    # Banner Superior
    st.markdown(f"""
    <div class="user-banner">
        <h2>Bienvenido(a), {user.get('APELLIDOS Y NOMBRES', 'Usuario')}</h2>
        <p>Coordinación Cantonal asignada a: <strong>{canton_asignado}</strong></p>
    </div>
    """, unsafe_allow_html=True)

    # Carga de datos activos
    if st.session_state["custom_coordinadores"] is not None:
        coordinadores_base = st.session_state["custom_coordinadores"].copy()
    else:
        recintos_canton = df_coord_recinto[df_coord_recinto["CANTON"] == canton_asignado].copy() if "CANTON" in df_coord_recinto.columns else pd.DataFrame()
        mesas_canton = df_coord_mesa[df_coord_mesa["CANTON"] == canton_asignado].copy() if "CANTON" in df_coord_mesa.columns else pd.DataFrame()
        coordinadores_base = pd.concat([recintos_canton, mesas_canton], ignore_index=True)

    # Filtrar por el cantón del usuario autenticado si existe la columna CANTON
    if "CANTON" in coordinadores_base.columns:
        df_canton_activo = coordinadores_base[coordinadores_base["CANTON"] == canton_asignado].copy()
    else:
        df_canton_activo = coordinadores_base.copy()

    # Mapeo de observaciones
    df_canton_activo["OBSERVACIONES"] = df_canton_activo["CEDULA"].map(
        st.session_state["observaciones_dict"]
    ).fillna(df_canton_activo.get("OBSERVACIONES", ""))

    # Métricas
    total_recintos = len(df_canton_activo[df_canton_activo.get("TIPO", "") == "Coordinador de Recinto"]) if "TIPO" in df_canton_activo.columns else 0
    total_mesas = len(df_canton_activo[df_canton_activo.get("TIPO", "") == "Coordinador de Mesa"]) if "TIPO" in df_canton_activo.columns else 0
    total_general = len(df_canton_activo)

    m1, m2, m3 = st.columns(3)
    with m1:
        st.markdown(f"""
        <div class="metric-card">
            <h3>Total Coordinadores</h3>
            <div class="number">{total_general}</div>
        </div>
        """, unsafe_allow_html=True)
    with m2:
        st.markdown(f"""
        <div class="metric-card">
            <h3>Coordinadores de Recinto</h3>
            <div class="number">{total_recintos}</div>
        </div>
        """, unsafe_allow_html=True)
    with m3:
        st.markdown(f"""
        <div class="metric-card">
            <h3>Coordinadores de Mesa</h3>
            <div class="number">{total_mesas}</div>
        </div>
        """, unsafe_allow_html=True)

    st.write("")
    st.write("")

    # Tabla y Filtros
    st.subheader(f"📋 Nómina de Coordinadores - Cantón {canton_asignado}")

    f_col1, f_col2 = st.columns([3, 1])
    with f_col1:
        busqueda = st.text_input("🔍 Buscar por Nombre o Cédula", placeholder="Escriba para filtrar...")
    with f_col2:
        filtro_tipo = st.selectbox("Filtrar por Tipo", ["Todos", "Coordinador de Recinto", "Coordinador de Mesa"])

    df_resultado = df_canton_activo.copy()

    # Filtrar datos por tipo
    if filtro_tipo != "Todos" and "TIPO" in df_resultado.columns:
        df_resultado = df_resultado[df_resultado["TIPO"] == filtro_tipo].copy()

    # Búsqueda general
    if busqueda and not df_resultado.empty:
        df_resultado = df_resultado[
            df_resultado["APELLIDOS Y NOMBRES"].astype(str).str.contains(busqueda, case=False, na=False) |
            df_resultado["CEDULA"].astype(str).str.contains(busqueda, case=False, na=False)
        ]

    # Construcción dinámica de columnas visibles
    base_cols = ["APELLIDOS Y NOMBRES", "CANTON", "CEDULA", "TELEFONO"]
    if filtro_tipo != "Todos":
        base_cols.append("TIPO")
    base_cols.append("OBSERVACIONES")

    # Incluir columnas adicionales detectadas en el DataFrame que sí contienen modificaciones
    extra_cols = [c for c in df_resultado.columns if c not in base_cols]
    cols_a_mostrar = base_cols + extra_cols

    cols_existentes = [c for c in cols_a_mostrar if c in df_resultado.columns]

    if not df_resultado.empty and len(cols_existentes) > 0:
        # Configuración para deshabilitar edición en columnas no deseadas
        col_config = {
            "OBSERVACIONES": st.column_config.TextColumn(
                "OBSERVACIONES / ERRORES",
                help="Escriba cualquier error o novedad en la información de esta fila",
                width="large"
            )
        }
        for col in cols_existentes:
            if col != "OBSERVACIONES":
                col_config[col] = st.column_config.Column(disabled=True)

        df_editado = st.data_editor(
            df_resultado[cols_existentes],
            column_config=col_config,
            use_container_width=True,
            hide_index=True,
            key="editor_nomina"
        )

        st.write("")
        
        # BARRA DE ACCIONES (Guardar, Importar, Exportar juntos)
        col_btn1, col_btn2, col_btn3, _ = st.columns([1.5, 1.5, 1.5, 3])
        
        with col_btn1:
            if st.button("💾 Guardar Observaciones", type="primary", use_container_width=True):
                for _, row in df_editado.iterrows():
                    cedula = str(row["CEDULA"])
                    obs = str(row["OBSERVACIONES"]).strip() if pd.notna(row["OBSERVACIONES"]) else ""
                    if obs:
                        st.session_state["observaciones_dict"][cedula] = obs
                    elif cedula in st.session_state["observaciones_dict"]:
                        del st.session_state["observaciones_dict"][cedula]
                
                st.success("✅ Observaciones guardadas exitosamente.")
                st.rerun()

        with col_btn2:
            if st.button("📥 Importar Nómina", use_container_width=True):
                abrir_dialog_importar(df_canton_activo, canton_asignado)

        with col_btn3:
            if st.button("📤 Exportar Nómina", use_container_width=True):
                df_exp = df_canton_activo.copy()
                df_exp["OBSERVACIONES"] = df_exp["CEDULA"].map(
                    st.session_state["observaciones_dict"]
                ).fillna("")
                abrir_dialog_exportar(df_exp, canton_asignado)

    else:
        st.warning("No se encontraron registros con los criterios seleccionados.")