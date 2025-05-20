import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io

st.set_page_config(page_title="Análisis de Productos Vendidos", layout="wide")
st.title("Análisis de Ventas: Descripciones Limpias y Outliers Agrupados")

@st.cache_data
def load_data():
    return pd.read_excel("productos_limpios.xlsx")

df = load_data()

familias = sorted(df["Familia"].dropna().unique())
familia = st.sidebar.selectbox("Selecciona una familia", familias)
df_fam = df[df["Familia"] == familia]

n_otros = df_fam["Es_Otros"].sum()
n_total = len(df_fam)
porc_otros = (n_otros / n_total) * 100 if n_total > 0 else 0

st.header(f"Familia seleccionada: {familia}")

st.subheader("Tabla de descripciones (limpias y finales) con cantidad vendida")
st.dataframe(df_fam[["Descripcion_clean", "Descripcion_final", "CantidadVendida"]].sort_values("CantidadVendida", ascending=False), use_container_width=True)

st.markdown(f"**Productos agrupados como 'OTROS':** {n_otros} de {n_total} ({porc_otros:.1f}%)")

st.subheader("Distribución de cantidad vendida por descripción limpia")
fig, ax = plt.subplots(figsize=(10, 4))
ax.boxplot(df_fam["CantidadVendida"])
ax.set_title(f"Boxplot de ventas ({familia})")
ax.set_ylabel("Cantidad Vendida")
st.pyplot(fig)

outliers = df_fam[df_fam["Es_Otros"] == 1]
if not outliers.empty:
    st.subheader("Descripciones agrupadas como 'OTROS'")
    st.dataframe(outliers[["Descripcion_clean", "Descripcion_final", "CantidadVendida"]], use_container_width=True)
else:
    st.info("No hay agrupaciones 'OTROS' en esta familia.")

# Botón de descarga de Excel
buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
    df_fam.to_excel(writer, index=False)
buffer.seek(0)
st.download_button(
    label="Descargar Excel de resultados",
    data=buffer,
    file_name=f"productos_limpios_{familia}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.markdown("---")
st.markdown("Puedes consultar el mapeo de fuzzy matching y el log de outliers en los archivos `log_fuzzy_matching.xlsx` y `log_outliers.xlsx` generados por el script de análisis.")

st.info("Este análisis usa una lógica adaptativa: agrupa solo los menos vendidos según percentil y un mínimo absoluto, y nunca agrupa familias muy pequeñas. Puedes ajustar MINIMO_ABSOLUTO y PERCENTIL en el script de análisis para personalizar la cantidad de agrupados.")
