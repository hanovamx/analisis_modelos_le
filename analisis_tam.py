import pandas as pd
import pyodbc
import re
import numpy as np

# Conexión SQL Server
conn = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server};'
    'SERVER=localhost,1433;'
    'DATABASE=JoyDB;'
    'UID=sa;'
    'PWD=Tec12345!'
)

query = """
SELECT
  P.idProducto,
  F.Nombre AS Familia,
  FT.Nombre AS Tamaño
FROM PV_VentasLineas VL
LEFT JOIN IN_Productos P ON VL.idProducto = P.idProducto
LEFT JOIN CAT_Familias F ON P.idFamilia = F.idFamilia
LEFT JOIN CAT_FamiliasTamaño FT ON P.idFamiliaTamaño = FT.idFamiliaTamaño
WHERE VL.idVentaLinea IN(
    SELECT idVentaLinea
    FROM PV_VentasLineas VL 
    JOIN PV_Ventas V ON VL.idVenta = V.idVenta
    WHERE Fecha >= '20240101' AND Fecha <= '20241231'
) 
AND FT.Nombre IS NOT NULL
ORDER BY P.idProducto ASC
"""

df = pd.read_sql(query, conn)
conn.close()

df["Tamaño"] = df["Tamaño"].astype(str).str.strip().str.upper()
df["Familia"] = df["Familia"].astype(str).str.strip().str.upper()

# === Función para extraer valor numérico (en cms) ===
def extraer_cm(valor):
    if pd.isnull(valor):
        return np.nan
    m = re.search(r'([\d\.]+)\s*CM[S]?', valor)
    if m:
        return float(m.group(1))
    m = re.search(r'([\d\.]+)\s*$', valor)  # Por si es solo el número
    if m:
        return float(m.group(1))
    # Ejemplos especiales (grande, chico, etc.) los puedes mapear aparte si gustas
    return np.nan

df["Tamaño_num"] = df["Tamaño"].apply(extraer_cm)

# === Definir lógica de rangos por familia ===
familia_rangos = {}
asignaciones = []

for familia, grupo in df.groupby("Familia"):
    # ¿Cuántos valores numéricos hay?
    tamanos_validos = grupo.dropna(subset=["Tamaño_num"])["Tamaño_num"]
    logica = ""
    if len(tamanos_validos) == 0:
        # Para familias sin valores numéricos, solo asigna el valor textual
        grupo["Rango_Tamaño"] = grupo["Tamaño"]
        logica = "Solo valores textuales. No se asignó rango automático."
    else:
        # Opciones de lógica: usar percentiles, clustering, o reglas fijas para algunas familias
        # Ejemplo para cadenas (puedes editar)
        if "CADENA" in familia:
            # 40-44 cm = corta, 45-50 = estándar, >50 = larga
            def rango_cadena(x):
                if pd.isnull(x):
                    return grupo["Tamaño"]
                elif x < 45:
                    return "CORTA"
                elif 45 <= x <= 50:
                    return "ESTÁNDAR"
                else:
                    return "LARGA"
            grupo["Rango_Tamaño"] = grupo["Tamaño_num"].apply(rango_cadena)
            logica = "CORTA: <45cm, ESTÁNDAR: 45-50cm, LARGA: >50cm"
        elif "MEDALLA" in familia or "DIJE" in familia or "CRUZ" in familia:
            # Usa percentiles por familia para pequeño, mediano, grande
            q1 = tamanos_validos.quantile(0.33)
            q2 = tamanos_validos.quantile(0.66)
            def rango_pct(x):
                if pd.isnull(x):
                    return grupo["Tamaño"]
                elif x <= q1:
                    return "PEQUEÑO"
                elif x <= q2:
                    return "MEDIANO"
                else:
                    return "GRANDE"
            grupo["Rango_Tamaño"] = grupo["Tamaño_num"].apply(rango_pct)
            logica = f"PEQUEÑO: ≤{q1:.2f}cm, MEDIANO: >{q1:.2f}–≤{q2:.2f}cm, GRANDE: >{q2:.2f}cm"
        elif "GARGANTILLA" in familia:
            # Lógica similar a cadenas
            def rango_gargantilla(x):
                if pd.isnull(x):
                    return grupo["Tamaño"]
                elif x < 45:
                    return "CORTA"
                elif 45 <= x <= 50:
                    return "ESTÁNDAR"
                else:
                    return "LARGA"
            grupo["Rango_Tamaño"] = grupo["Tamaño_num"].apply(rango_gargantilla)
            logica = "CORTA: <45cm, ESTÁNDAR: 45-50cm, LARGA: >50cm"
        else:
            # Por defecto usa cuartiles (puedes tunear familia por familia después)
            q1 = tamanos_validos.quantile(0.25)
            q3 = tamanos_validos.quantile(0.75)
            def rango_def(x):
                if pd.isnull(x):
                    return grupo["Tamaño"]
                elif x <= q1:
                    return "PEQUEÑO"
                elif x <= q3:
                    return "MEDIANO"
                else:
                    return "GRANDE"
            grupo["Rango_Tamaño"] = grupo["Tamaño_num"].apply(rango_def)
            logica = f"PEQUEÑO: ≤{q1:.2f}cm, MEDIANO: >{q1:.2f}–≤{q3:.2f}cm, GRANDE: >{q3:.2f}cm"
    familia_rangos[familia] = logica
    asignaciones.append(grupo)

# Unir todo
df_rangos = pd.concat(asignaciones)

# Exporta resultados y lógica
df_rangos.to_excel("productos_con_rangos_tamano.xlsx", index=False)
pd.DataFrame([
    {"Familia": fam, "Lógica_Rango": logica}
    for fam, logica in familia_rangos.items()
]).to_excel("logica_rangos_familias.xlsx", index=False)

print("¡Listo! Resultados exportados como 'productos_con_rangos_tamano.xlsx' y 'logica_rangos_familias.xlsx'.")
