import pandas as pd
import pyodbc
from fuzzywuzzy import process, fuzz

# Conexión SQL Server
conn = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server};'
    'SERVER=localhost,1433;'
    'DATABASE=JoyDB;'
    'UID=sa;'
    'PWD=Tec12345!'
)

# --- QUERY: Vendidos (uno por cada venta, cantidad=1) y no vendidos (cantidad=0) ---
query = """
SELECT
  P.idProducto,
  F.Nombre as Familia,
  fd.Nombre AS Descripcion,
  1 as CantidadVendida
FROM PV_VentasLineas VL
LEFT JOIN IN_Productos P ON VL.idProducto = P.idProducto
LEFT JOIN CAT_Familias F ON P.idFamilia = F.idFamilia
LEFT JOIN CAT_FamiliasDescripciones FD ON P.idDescripcion = FD.idDescripcion
WHERE P.idProducto IN(
    SELECT idProducto
    FROM PV_VentasLineas VL JOIN PV_Ventas V ON VL.idVenta = V.idVenta
    WHERE Fecha >= '20240101' AND Fecha <= '20241231'
)
UNION ALL
SELECT
  P.idProducto,
  F.Nombre AS Familia,
  FD.Nombre AS Descripcion,
  0 AS CantidadVendida
FROM IN_Productos P
LEFT JOIN CAT_Familias F ON P.idFamilia = F.idFamilia
LEFT JOIN CAT_FamiliasDescripciones FD ON P.idDescripcion = FD.idDescripcion
WHERE P.idProducto IN (
    SELECT DISTINCT idProducto
    FROM IN_FacturasDet
    WHERE FechaReg >= '20240101' AND FechaReg <= '20241231'
)
AND P.idProducto NOT IN (
    SELECT DISTINCT VL.idProducto
    FROM PV_VentasLineas VL
    INNER JOIN PV_Ventas V ON VL.idVenta = V.idVenta
    WHERE V.Fecha >= '20240101' AND V.Fecha <= '20241231'
)
"""

df = pd.read_sql(query, conn)
conn.close()

# Limpieza y normalización
df["Descripcion"] = df["Descripcion"].astype(str).str.strip().str.upper()
df["Familia"] = df["Familia"].astype(str).str.strip().str.upper()

# Agrupa por producto, familia y descripción para sumar ventas (por si hay varias filas para el mismo producto)
ventas = df.groupby(['idProducto', 'Familia', 'Descripcion'], as_index=False)['CantidadVendida'].sum()

# Fuzzy Matching por familia
descripcion_map = {}
for familia, grupo in ventas.groupby("Familia"):
    descripciones = grupo["Descripcion"].tolist()
    asignado = set()
    for desc in descripciones:
        if desc in asignado:
            continue
        similars = process.extract(desc, descripciones, scorer=fuzz.token_sort_ratio, limit=None)
        cluster = [d for d, score in similars if score >= 88]
        for d in cluster:
            asignado.add(d)
            descripcion_map[(familia, d)] = desc

ventas["Descripcion_clean"] = ventas.apply(
    lambda row: descripcion_map.get((row["Familia"], row["Descripcion"]), row["Descripcion"]),
    axis=1
)

# Suma por familia + descripción limpia para obtener la venta total por agrupamiento
ventas_clean = ventas.groupby(["Familia", "Descripcion_clean"]).agg({"CantidadVendida": "sum"}).reset_index()

# Lógica adaptativa para agrupación
MINIMO_ABSOLUTO = 3      # Ajusta según tu criterio
PERCENTIL = 0.10         # Percentil bajo para cada familia

agrupados = []
outliers_log = []

for familia, grupo in ventas_clean.groupby("Familia"):
    if len(grupo) <= 5:
        grupo["Descripcion_final"] = grupo["Descripcion_clean"]
        grupo["Es_Otros"] = 0
    else:
        limite_percentil = grupo["CantidadVendida"].quantile(PERCENTIL)
        umbral_familia = max(limite_percentil, MINIMO_ABSOLUTO)
        def sugerir_nombre(desc):
            return f"OTROS {familia.upper()}"
        grupo["Descripcion_final"] = grupo.apply(
            lambda row: sugerir_nombre(row["Descripcion_clean"]) if (row["CantidadVendida"] <= umbral_familia) else row["Descripcion_clean"],
            axis=1
        )
        grupo["Es_Otros"] = grupo.apply(
            lambda row: 1 if row["Descripcion_final"].startswith("OTROS ") and row["Descripcion_final"] != row["Descripcion_clean"] else 0,
            axis=1
        )
    for _, r in grupo.iterrows():
        if r["Descripcion_final"] != r["Descripcion_clean"]:
            outliers_log.append({
                "Familia": familia,
                "Descripcion_clean": r["Descripcion_clean"],
                "CantidadVendida": r["CantidadVendida"],
                "Descripcion_final": r["Descripcion_final"]
            })
    agrupados.append(grupo)

ventas_final = pd.concat(agrupados)

# Output para Streamlit
ventas_final.to_excel("productos_limpios.xlsx", index=False)
print("Archivo productos_limpios.xlsx generado con productos vendidos y NO vendidos.")

# Log de clusters fuzzy
fuzzy_log = pd.DataFrame([
    {"Familia": f, "Descripcion_original": d, "Descripcion_cluster": c}
    for (f, d), c in descripcion_map.items()
])
fuzzy_log.to_excel("log_fuzzy_matching.xlsx", index=False)

# Log de outliers agrupados
pd.DataFrame(outliers_log).to_excel("log_outliers.xlsx", index=False)
