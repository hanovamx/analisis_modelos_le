import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(page_title="Análisis de Rangos de Tamaño", layout="wide")
st.title("Análisis de Rangos de Tamaño por Familia")

@st.cache_data
def load_data():
    df_rangos = pd.read_excel("productos_con_rangos_tamano.xlsx")
    logica_rangos = pd.read_excel("logica_rangos_familias.xlsx")
    return df_rangos, logica_rangos

df, logica = load_data()

familias = sorted(df["Familia"].dropna().unique())
familia = st.sidebar.selectbox("Selecciona una familia", familias)
df_fam = df[df["Familia"] == familia]

# Mostrar lógica usada
logica_texto = logica.loc[logica["Familia"] == familia, "Lógica_Rango"].values
if len(logica_texto) > 0:
    st.info(f"Lógica de rangos para **{familia}**: {logica_texto[0]}")
else:
    st.warning("No hay lógica registrada para esta familia.")

st.subheader(f"Distribución de tamaños para {familia}")

# Solo si hay datos numéricos
if df_fam["Tamaño_num"].notnull().sum() > 0:
    fig, ax = plt.subplots(figsize=(8,4))
    df_fam[df_fam["Tamaño_num"].notnull()]["Tamaño_num"].hist(bins=15, ax=ax)
    ax.set_xlabel("Tamaño (cm)")
    ax.set_ylabel("Cantidad de productos vendidos")
    st.pyplot(fig)
else:
    st.info("No hay tamaños numéricos para graficar en esta familia.")

# Tabla de rangos
st.subheader("Tabla de productos y rango asignado")
st.dataframe(df_fam[["idProducto", "Tamaño", "Tamaño_num", "Rango_Tamaño"]], use_container_width=True)

# Botón de descarga
import io

buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
    df_fam.to_excel(writer, index=False)
buffer.seek(0)

st.download_button(
    label="Descargar Excel de familia",
    data=buffer,
    file_name=f"productos_rangos_{familia}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.markdown("---")
st.markdown("La lógica utilizada para cada familia puede editarse en el script Python. Si necesitas un rango especial, indícalo ahí o avísame y te lo ayudo a personalizar.")
