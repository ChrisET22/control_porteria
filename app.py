import datetime
import os
import shutil
import urllib.parse
import pandas as pd
import streamlit as st

EXCEL_FILE = "registro_porteria.xlsx"
CARPETA_HISTORIAL = "Historial_Diario"

# Listas de opciones predeterminadas
OPCIONES_EMPRESA = [
    "CASATORO RENAULT",
    "CASATORO FORD",
    "CASATORO VOLKSWAGEN",
    "CASATORO BONAPARTE",
    "BONAPARTE",
    "MELOLLEVO FINANZAUTO",
    "MELOLLEVO FINANDINA",
    "MELOLLEVO CASATORO",
    "CASATORO FUERA DE ESTANDAR",
    "FUERA DE ESTANDAR",
    "FINANDINA",
    "FINANZAUTO",
    "INCHAPE",
    "DERCO",
    "PRACO",
    "MATRASE",
    "VIAUTOS",
    "CARFIAO",
    "ALIANZA",
    "AUTOCOM",
    "ALCALA",
    "BRACHOAUTOS",
    "AUTOLAND",
    "MOTORYSA",
    "EQUIRENT",
    "BELLPI",
    "TOYOTA",
    "DAVIVIENDA",
    "AMBACAR",
    "ARVAL",
    "METROKIA",
    "COOVITEL",
    "RIDDARA",
    "CASATORO MAZDA",
]

OPCIONES_PROCEDIMIENTO = [
    "RETIRAR VEHICULO/S",
    "INSPECCIONAR VEHICULO/S",
    "INSTALACION/DESINSTALACION ACCESORIOS",
    "INSTALACION GPS",
    "INSTALACION PELICULAS",
    "RECOGER RECICLAJE",
    "ENTREGAR ENVIO",
    "RECOGER ENCOMIENDA",
    "RECOGER ENVIO",
    "RECOGER MERCANCIA",
    "DEJAR MERCANCIA",
]


def inicializar_excel():
  if not os.path.exists(EXCEL_FILE):
    df = pd.DataFrame(
        columns=[
            "Fecha",
            "Nombre",
            "Cédula",
            "Empresa/Patio",
            "Procedimiento",
            "Placas/VIN",
            "Hora Ingreso",
            "Hora Salida",
            "Estado",
        ]
    )
    df.to_excel(EXCEL_FILE, index=False)


inicializar_excel()

st.set_page_config(
    page_title="Control Portería", page_icon="🛃", layout="wide"
)

# Estilo personalizado para el botón verde suave
st.markdown(
    """
    <style>
    div.stButton > button[kind="primary"] {
        background-color: #4CAF50 !important;
        color: white !important;
        border: none !important;
        font-weight: bold !important;
        font-size: 16px !important;
        padding: 10px 24px !important;
        border-radius: 8px !important;
        transition: 0.3s !important;
    }
    div.stButton > button[kind="primary"]:hover {
        background-color: #45a049 !important;
        border: none !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("🛃 Control de Acceso - Portería")

# Inicializar estados de formulario
if "ing_nombre" not in st.session_state:
  st.session_state.ing_nombre = ""
if "ing_cedula" not in st.session_state:
  st.session_state.ing_cedula = ""
if "ing_placas" not in st.session_state:
  st.session_state.ing_placas = ""
if "manual_empresa" not in st.session_state:
  st.session_state.manual_empresa = False
if "manual_procedimiento" not in st.session_state:
  st.session_state.manual_procedimiento = False


def guardar_registro():
  nombre = st.session_state.get("ing_nombre", "").strip()
  cedula = st.session_state.get("ing_cedula", "").strip()

  if not nombre or not cedula:
    st.session_state.msg_error = (
        "⚠️ Por favor ingrese al menos el Nombre y la Cédula."
    )
    return

  # Obtener valor de empresa
  if st.session_state.manual_empresa:
    empresa_val = st.session_state.get("txt_empresa", "")
  else:
    sel = st.session_state.get("sel_empresa", "")
    empresa_val = "" if sel.startswith("--") else sel

  # Obtener valor de procedimiento
  if st.session_state.manual_procedimiento:
    procedimiento_val = st.session_state.get("txt_procedimiento", "")
  else:
    sel_p = st.session_state.get("sel_procedimiento", "")
    procedimiento_val = "" if sel_p.startswith("--") else sel_p

  placas = st.session_state.get("ing_placas", "")

  fecha_actual = datetime.datetime.now().strftime("%Y-%m-%d")
  hora_actual = datetime.datetime.now().strftime("%I:%M %p")

  df = pd.read_excel(EXCEL_FILE, dtype=str).fillna("")

  nuevo_registro = {
      "Fecha": fecha_actual,
      "Nombre": nombre,
      "Cédula": cedula,
      "Empresa/Patio": empresa_val,
      "Procedimiento": procedimiento_val,
      "Placas/VIN": placas,
      "Hora Ingreso": hora_actual,
      "Hora Salida": "",
      "Estado": "Ingresó",
  }

  df = pd.concat([df, pd.DataFrame([nuevo_registro])], ignore_index=True)
  df.to_excel(EXCEL_FILE, index=False)

  # Mensaje de éxito
  st.session_state.msg_exito = f"✅ Ingreso registrado para {nombre}"

  # Limpiar campos de texto para el nuevo registro
  st.session_state.ing_nombre = ""
  st.session_state.ing_cedula = ""
  st.session_state.ing_placas = ""
  st.session_state.manual_empresa = False
  st.session_state.manual_procedimiento = False

  if "txt_empresa" in st.session_state:
    st.session_state.txt_empresa = ""
  if "txt_procedimiento" in st.session_state:
    st.session_state.txt_procedimiento = ""
  if "sel_empresa" in st.session_state:
    st.session_state.sel_empresa = "-- SELECCIONAR O BUSCAR EN LISTA --"
  if "sel_procedimiento" in st.session_state:
    st.session_state.sel_procedimiento = "-- SELECCIONAR O BUSCAR EN LISTA --"


# ---- FORMULARIO DE INGRESO ----
st.subheader("📝 Registrar Nuevo Ingreso")

col_f1, col_f2 = st.columns(2)

with col_f1:
  st.text_input("Nombre Completo:", key="ing_nombre")
  st.text_input("Cédula:", key="ing_cedula")

  # Manejo Híbrido Dinámico para Empresa / Patio
  if not st.session_state.manual_empresa:
    st.selectbox(
        "Empresa / Patio Destino (Seleccione o escriba abajo):",
        options=["-- SELECCIONAR O BUSCAR EN LISTA --"] + OPCIONES_EMPRESA,
        key="sel_empresa",
    )
    chk_empresa = st.checkbox(
        "✏️ Escribir empresa/patio manualmente (no está en la lista)",
        key="chk_emp",
    )
    if chk_empresa:
      st.session_state.manual_empresa = True
      st.rerun()
  else:
    st.text_input("Escriba la Empresa / Patio Destino:", key="txt_empresa")
    if st.button("⬅️ Volver a lista de opciones", key="btn_volver_emp"):
      st.session_state.manual_empresa = False
      st.rerun()

with col_f2:
  # Manejo Híbrido Dinámico para Procedimiento
  if not st.session_state.manual_procedimiento:
    st.selectbox(
        "Procedimiento / Labor (Seleccione o escriba abajo):",
        options=["-- SELECCIONAR O BUSCAR EN LISTA --"] + OPCIONES_PROCEDIMIENTO,
        key="sel_procedimiento",
    )
    chk_procedimiento = st.checkbox(
        "✏️ Escribir procedimiento manualmente (no está en la lista)",
        key="chk_proc",
    )
    if chk_procedimiento:
      st.session_state.manual_procedimiento = True
      st.rerun()
  else:
    st.text_input("Escriba el Procedimiento / Labor:", key="txt_procedimiento")
    if st.button("⬅️ Volver a lista de opciones", key="btn_volver_proc"):
      st.session_state.manual_procedimiento = False
      st.rerun()

  st.text_input("Placas / VIN:", key="ing_placas")

st.write("")
st.button(
    "🟢 REGISTRAR INGRESO",
    type="primary",
    use_container_width=True,
    on_click=guardar_registro,
)

# Mostrar alertas de estado tras el guardado
if "msg_exito" in st.session_state:
  st.success(st.session_state.msg_exito)
  del st.session_state.msg_exito

if "msg_error" in st.session_state:
  st.error(st.session_state.msg_error)
  del st.session_state.msg_error

# ---- TABLA DE REGISTROS (ORDEN MÁS RECIENTE ARRIBA) ----
st.write("---")
st.subheader("📋 Control de Personal del Día")

df_registros = pd.read_excel(EXCEL_FILE, dtype=str).fillna("")

if not df_registros.empty:
  for idx in reversed(df_registros.index):
    row = df_registros.loc[idx]
    col_datos, col_wa, col_salida, col_eliminar = st.columns([4, 1, 1, 1])

    ya_salio = row["Estado"] == "Se Retiró"

    with col_datos:
      if ya_salio:
        st.markdown(
            f"~~**{row['Nombre']}** ({row['Cédula']})~~ &nbsp;|&nbsp; 🏢 Destino:"
            f" **{row['Empresa/Patio']}** &nbsp;|&nbsp; 📋 Labor:"
            f" **{row['Procedimiento']}** &nbsp;|&nbsp; 🚛 Placas:"
            f" **{row['Placas/VIN']}**  \n🔴 **RETIRADO** (Entrada:"
            f" {row['Hora Ingreso']} | Salida: {row['Hora Salida']})"
        )
      else:
        st.markdown(
            f"**{row['Nombre']}** ({row['Cédula']}) &nbsp;|&nbsp; 🏢 Destino:"
            f" **{row['Empresa/Patio']}** &nbsp;|&nbsp; 📋 Labor:"
            f" **{row['Procedimiento']}** &nbsp;|&nbsp; 🚛 Placas:"
            f" **{row['Placas/VIN']}**  \n🟢 **EN INSTALACIONES** (Entrada:"
            f" {row['Hora Ingreso']})"
        )

    with col_wa:
      estado_txt = (
          f"🔴 RETIRADO (Salida: {row['Hora Salida']})"
          if ya_salio
          else "🟢 EN INSTALACIONES"
      )
      mensaje_wa = (
          f"🛃 *CONTROL DE ACCESO PORTERÍA*\n\n"
          f"👤 *Nombre:* {row['Nombre']}\n"
          f"🆔 *Cédula:* {row['Cédula']}\n"
          f"🏢 *Destino:* {row['Empresa/Patio']}\n"
          f"📋 *Procedimiento:* {row['Procedimiento']}\n"
          f"🚛 *Placas/VIN:* {row['Placas/VIN']}\n"
          f"⏰ *Hora Ingreso:* {row['Hora Ingreso']}\n"
          f"📊 *Estado:* {estado_txt}"
      )
      url_wa = f"https://wa.me/?text={urllib.parse.quote(mensaje_wa)}"

      st.markdown(
          f"""
            <a href="{url_wa}" target="_blank">
                <button style="
                    background-color:#25D366; 
                    color:white; 
                    padding:6px 12px; 
                    border:none; 
                    border-radius:5px; 
                    width:100%; 
                    font-size:13px; 
                    font-weight:bold; 
                    cursor:pointer;
                    margin-top:2px;">
                    📲 Compartir
                </button>
            </a>
        """,
          unsafe_allow_html=True,
      )

    with col_salida:
      if ya_salio:
        st.button("✔️ Retirado", key=f"salida_{idx}", disabled=True)
      else:
        if st.button("🔴 Salida", key=f"salida_{idx}"):
          df_registros.loc[idx, "Hora Salida"] = (
              datetime.datetime.now().strftime("%I:%M %p")
          )
          df_registros.loc[idx, "Estado"] = "Se Retiró"
          df_registros.to_excel(EXCEL_FILE, index=False)
          st.rerun()

    with col_eliminar:
      if st.button("🗑️ Borrar", key=f"borrar_{idx}"):
        st.session_state[f"confirm_borrar_{idx}"] = True

    if st.session_state.get(f"confirm_borrar_{idx}", False):
      st.warning(f"⚠️ ¿Eliminar registro de **{row['Nombre']}**?")
      col_si, col_no = st.columns(2)
      with col_si:
        if st.button("Sí, eliminar", key=f"confirm_si_{idx}"):
          df_registros = df_registros.drop(idx)
          df_registros.to_excel(EXCEL_FILE, index=False)
          del st.session_state[f"confirm_borrar_{idx}"]
          st.rerun()
      with col_no:
        if st.button("Cancelar", key=f"confirm_no_{idx}"):
          del st.session_state[f"confirm_borrar_{idx}"]
          st.rerun()

    st.markdown(
        "<hr style='margin: 6px 0; border-top: 1px solid #e6e6e6;'>",
        unsafe_allow_html=True,
    )
else:
  st.info("No hay registros en la lista actual.")

# ---- CIERRE DE JORNADA Y ARCHIVADO ----
st.write("---")
st.subheader("🏁 Cierre de Jornada / Nuevo Día")

with st.expander("📂 Opciones de Finalización de Día"):
  st.write(
      "Al cerrar la jornada, el registro actual se guardará en la carpeta"
      f" **'{CARPETA_HISTORIAL}'** organizado por la fecha de hoy, y la"
      " pantalla quedará limpia para la siguiente jornada."
  )

  if st.button("📁 ARCHIVAR Y EMPEZAR NUEVO DÍA"):
    if not df_registros.empty:
      if not os.path.exists(CARPETA_HISTORIAL):
        os.makedirs(CARPETA_HISTORIAL)

      fecha_hoy = datetime.datetime.now().strftime("%Y-%m-%d")
      archivo_destino = os.path.join(
          CARPETA_HISTORIAL, f"Registro_Porteria_{fecha_hoy}.xlsx"
      )

      shutil.copy(EXCEL_FILE, archivo_destino)

      df_vacio = pd.DataFrame(columns=df_registros.columns)
      df_vacio.to_excel(EXCEL_FILE, index=False)

      st.success(
          f"🎉 Jornada archivada exitosamente como '{archivo_destino}'. La"
          " pantalla está lista para un nuevo día."
      )
      st.rerun()
    else:
      st.warning(
          "El registro actual ya está vacío. No hay datos para archivar."
      )