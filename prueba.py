#Importación de librerías
import pandas as pd
import numpy as np
from datetime import datetime
import re

#Carga de datos
base = ("https://view.officeapps.live.com/op/view.aspx?src=https%3A%2F%2Fraw.githubusercontent.com%2Fkarol-vann%2FPruebaTecnicaBC%2Frefs%2Fheads%2Fmain%2Fprueba.xlsx&wdOrigin=BROWSELINK")
exportación = ("C:/Users/kpote/Documents/Personal/prueba/Resultados_VaR.xlsx")

fecha_valoracion = pd.to_datetime("2026-04-13")

#Lectura de datos
posiciones = pd.read_excel(base, sheet_name="Posiciones")
icolcap = pd.read_excel(base, sheet_name="ICOLCAP")
titulos = pd.read_excel(base, sheet_name="Info. Títulos")
ibr = pd.read_excel(base, sheet_name="C_IBR")
c_cec = pd.read_excel(base, sheet_name="CEC")
c_cecuvr = pd.read_excel(base, sheet_name="CECUVR")
c_ois_proy = pd.read_excel(base, sheet_name="C_OISCOP_Proy")

#------- funciones auxiliares -------

def crear_interpolador_curva(df_fila_curva, sufijo):
    """
    Transforma una fila de Excel con nodos en las columnas en un 
    interpolador funcional.
    """
    # 1. Filtrar columnas que contienen el sufijo (ej: 'CEC' o 'UVR')
    cols_nodos = [c for c in df_fila_curva.columns if sufijo in c]
    
    # 2. Extraer los números (días) de los nombres de las columnas
    # N1080CEC -> 1080
    dias = [int(re.findall(r'\d+', c)[0]) for c in cols_nodos]
    
    # 3. Obtener los valores de la primera fila
    tasas = df_fila_curva[cols_nodos].iloc[0].values
    
    # 4. Ordenar puntos para np.interp
    puntos = sorted(zip(dias, tasas))
    x, y = zip(*puntos)
    
    # 5. Retornar una función lambda que interpola sobre estos puntos
    return lambda d: np.interp(d, x, y)

# Definir la lógica de asignación masiva
def asignar_tasa_mercado(row):
    dias = row['dias_vencimiento']
    if pd.isna(dias) or dias <= 0:
        return 0.0
    
    # Si es TES Pesos o moneda COP, usar curva CEC
    if "TES" in row['Activo'] and row['Moneda'] == "COP":
        return interp_cec(dias)
    
    # Si es TES UVR o moneda UVR, usar curva CECUVR
    elif "TES" in row['Activo'] and row['Moneda'] == "UVR":
        return interp_uvr(dias)
    
    # Si es Swap, usar curva OIS
    elif "Swap" in row['Activo']:
        return interp_ois(dias)
    
    return row['tasa_cupon']

def obtener_nodos_disponibles(df_curva, sufijo):
    cols = [c for c in df_curva.columns if sufijo in c]
    # Extrae solo el número: 'N1080CEC' -> 1080
    return sorted([int(re.findall(r'\d+', c)[0]) for c in cols])

def asignar_factor_riesgo(row, dic_nodos):
    """Busca el nodo más cercano en la curva que corresponde al activo"""
    macaulay_años = row['Duracion_Macaulay']
    if macaulay_años <= 0:
        return "ICOLCAP"  #para el ETF iShares COLCAP
    
    # 1. Identificar sufijo de curva
    if "TES" in str(row['Activo']).upper():
        sufijo = "CEC" if row['Moneda'] == "COP" else "CECUVR"
    elif "SWAP" in str(row['Activo']).upper():
        sufijo = "OISCOP_Proy"
    else:
        sufijo = "CEC" # Por defecto o para CDT
    
    # 2. Obtener lista de nodos para esa curva
    nodos_disponibles = dic_nodos.get(sufijo, [])
    if not nodos_disponibles:
        return "N/A"
    
    # 3. Encontrar el número de nodo más cercano a la duración (en días)
    dias_objetivo = macaulay_años * 365
    num_nodo = min(nodos_disponibles, key=lambda x: abs(x - dias_objetivo))
    
    return f"N{num_nodo}{sufijo}"


def calcular_duracion_modificada(row):
    if row['tasa_mercado'] == 0 or pd.isna(row['dias_vencimiento']):
        return 1, 0.0
    base_dias = 360 if ("CDT" in row['Activo'] or "Swap" in row['Activo']) else 365
    t_años = row['dias_vencimiento'] / base_dias
    tasa_mercado = row['tasa_mercado']
    cupon = row['tasa_cupon']
    
    # Renta Fija Tradicional (TES)
    if "TES" in row['Activo']:
        # Maculay simplificada para bonos bullet
        macaulay = ((1+tasa_mercado)/tasa_mercado) - ((1+tasa_mercado + t_años*(cupon-tasa_mercado)) / (cupon*((1+tasa_mercado)**t_años - 1) + tasa_mercado))
        modificada = macaulay / (1 + tasa_mercado)
        return modificada,macaulay
    
    # Renta Variable Proyectada (CDT IBR)
    elif "CDT" in row['Activo']:
        macaulay = 0.25 # pago trimestral
        modificada = macaulay / (1 + tasa_mercado)
        return modificada,macaulay
    
    # Derivados (Swap IBR)
    elif "Swap" in row['Activo']:
        dur_fija = (1 - (1 + tasa_mercado)**-t_años) / tasa_mercado
        dur_variable = 0.0027 # Overnight 1/360
        macaulay = dur_fija - dur_variable
        modificada = (dur_fija - dur_variable) / (1 + tasa_mercado)
        return modificada, macaulay
    
    return 1, 0.0



#Limpieza de datos

icolcap = icolcap.sort_values(by="Fecha", ascending=False)
icolcap = icolcap.rename(columns={"NAV por acción": "ICOLCAP"})
#cambiar nombre de columna Row por Fecha
ibr = ibr.rename(columns={"Row": "Fecha"})
ibr = ibr.sort_values(by="Fecha", ascending=False)
c_cec = c_cec.rename(columns={"Row": "Fecha"})
c_cec = c_cec.sort_values(by="Fecha", ascending=False)
c_cecuvr = c_cecuvr.rename(columns={"Row": "Fecha"})
c_cecuvr = c_cecuvr.sort_values(by="Fecha", ascending=False)
c_ois_proy = c_ois_proy.rename(columns={"Row": "Fecha"})
c_ois_proy = c_ois_proy.sort_values(by="Fecha", ascending=False)

#--Tratamiento de valores faltantes--
c_cec = c_cec.fillna(method='ffill')
c_cecuvr = c_cecuvr.fillna(method='ffill')
c_ois_proy = c_ois_proy.fillna(method='ffill')
icolcap = icolcap.fillna(method='ffill')

print("Conteo de valores faltantes por nodo para CEC:")
print(c_cec.isnull().sum())
print("\nConteo de valores faltantes por nodo para CECUVR:")
print(c_cecuvr.isnull().sum())
print("\nConteo de valores faltantes por nodo para OIS:")
print(c_ois_proy.isnull().sum())
print("\nConteo de valores faltantes para ICOLCAP:")
print(icolcap.isnull().sum())

#Filtrar por fecha de valoración
c_cec_dia = c_cec[c_cec["Fecha"] == fecha_valoracion]
c_cecuvr_dia = c_cecuvr[c_cecuvr["Fecha"] == fecha_valoracion]
c_ois_proy_dia = c_ois_proy[c_ois_proy["Fecha"] == fecha_valoracion]
posiciones.loc[posiciones["ID"] == "A6", "ID"] = "A5" 

#----------- Base de portafolio unificada con información de títulos -----------

#Base de portafolio unificada con información de títulos
base_portafolio = posiciones.merge(titulos, on="ID", how="left")
base_portafolio = base_portafolio.drop(columns=["Activo_x"]).rename(columns={"Activo_y": "Activo"})
base_portafolio.insert(1, "Activo", base_portafolio.pop("Activo")) #dejar solo una columna de Activo
#separar numero de letra
base_portafolio["tasa_cupon"]  = base_portafolio["Tasa / Cupón"].str.extract(r'(\d+\.?\d*)').astype(float) / 100
ibr3m_cdt = ibr["N93IBR"].iloc[0] #tomamos el ibr 3m para hallar el valor del cupon

#calcular cupon para CDT IBR
base_portafolio["tasa_cupon"] = np.where(
    base_portafolio['Indexador'] == 'IBR 3M',
   (1 + ibr3m_cdt) * (1 + base_portafolio["tasa_cupon"]) - 1, 
    base_portafolio['tasa_cupon']                  
)

#calcular dias al vencimiento
base_portafolio["dias_vencimiento"] = (base_portafolio["Fecha Vencimiento"] - fecha_valoracion).dt.days
#valor de mercado actual

# Interpolación para hallar la tasa
interp_cec = crear_interpolador_curva(c_cec_dia, "CEC")
interp_uvr = crear_interpolador_curva(c_cecuvr_dia, "CECUVR")
interp_ois   = crear_interpolador_curva(c_ois_proy_dia, "OISCOP_Proy")

# Aplicar a todo el portafolio
base_portafolio['tasa_mercado'] = base_portafolio.apply(asignar_tasa_mercado, axis=1)
base_portafolio['Duracion_Modificada'], base_portafolio['Duracion_Macaulay'] = zip(*base_portafolio.apply(calcular_duracion_modificada, axis=1))

# --- CONFIGURACIÓN DE NODOS POR CURVA ---
nodos_cec = obtener_nodos_disponibles(c_cec, "CEC")
nodos_uvr = obtener_nodos_disponibles(c_cecuvr, "UVR")
nodos_ois = obtener_nodos_disponibles(c_ois_proy, "OIS")

# --- asignar nodo de filtro para cada activo ---
dict_nodos_referencia = {
    "CEC": nodos_cec,
    "CECUVR": nodos_uvr,
    "OISCOP_Proy": nodos_ois
}

# Creamos la nueva columna con el factor asociado a cada activo
base_portafolio['factor_riesgo'] = base_portafolio.apply(
    asignar_factor_riesgo, 
    args=(dict_nodos_referencia,), 
    axis=1
)

#Dejar en un sola moneda la posición para el calculo del VAR

"""
condiciones y resultados para calcular la posición en pesos dependiendo de si es 
UVR o ETF iShares COLCAP, si no es ninguno de los dos se deja el valor original de la posición en número
"""
condiciones = [
    (base_portafolio["Moneda"] == "UVR"),
    (base_portafolio["Activo"] == "ETF iShares COLCAP")
]

resultados = [
    base_portafolio["Posición (Número)"] * 410.5604,
    base_portafolio["Posición (Número)"] * icolcap["ICOLCAP"].iloc[0]
]

base_portafolio["Posicion_Nominal"] = np.select(condiciones, resultados, default=base_portafolio["Posición (Número)"])

base_depurada = base_portafolio[["ID", "Activo", "Tipo", "Moneda", "Posicion_Nominal", "tasa_cupon", "tasa_mercado", "Duracion_Modificada", "factor_riesgo"]]

#%%----------------- VAR PARAMETRICO ----------------------


#------------1 MATRIZ DE POSICIONES EN PESOS-------------
#unificar precios y calcular retornnos
df_precios = pd.concat([c_cec.set_index("Fecha"), c_cecuvr.set_index("Fecha"),
            icolcap.set_index("Fecha"),c_ois_proy.set_index("Fecha")], axis=1)             
retornos_historicos = df_precios.pct_change().dropna()

# ------------ 2 CALCULO DE VOLATILIDADES CON EWMA -------------
# Parámetro de decaimiento para EWMA
lambda_ewma = 0.94 
com = lambda_ewma / (1 - lambda_ewma)
# la varianza (retornos al cuadrado con media móvil exponencial)
varianza_ewma = retornos_historicos.pow(2).ewm(com=com, adjust=False).mean()
volatilidades_ewma = np.sqrt(varianza_ewma.iloc[-1])
# crear un diccionario para mapear al portafolio
mapa_vols_ewma = volatilidades_ewma.to_dict()
#asignar la volatilidad a cada activo según su factor de riesgo
base_depurada['vol_ewma'] = base_depurada['factor_riesgo'].map(mapa_vols_ewma)

#------------ 3 MARTRIZ DE CORRELACIÓN -------------
matriz_correlacion = retornos_historicos.corr()
factores_en_portafolio = base_depurada['factor_riesgo'].unique()
matriz_correlacion_portafolio = matriz_correlacion.loc[factores_en_portafolio, factores_en_portafolio]

# -------------4 VAR INDIVIDUAL-------------

nivel_confianza = 2.33  # para un nivel de confianza del 99%
base_depurada['VaR_Individual_1dia'] = (base_depurada['Posicion_Nominal'] * base_depurada['Duracion_Modificada']
                                         * base_depurada['vol_ewma'] * nivel_confianza* np.sqrt(1)
)
base_depurada['VaR_Individual_1dia%'] = (base_depurada['VaR_Individual_1dia'] / base_depurada['Posicion_Nominal']) * 100

# -------------5 VAR PORTAFOLIO-------------
# vector de VaR individual
vector_var = base_depurada['VaR_Individual_1dia'].values
# multiplicación matricial para obtener el VaR total del portafolio
# vector_var.dot(corr) multiplica el vector por la matriz, luego .dot(vector_var por el vector traspuesto)
var_portafolio_cuadrado = np.dot(np.dot(vector_var, matriz_correlacion_portafolio), vector_var.T)
var_total_portafolio = np.sqrt(var_portafolio_cuadrado)
var_porcentaje = (var_total_portafolio / base_depurada['Posicion_Nominal'].sum()) * 100

print(f"VaR Paramétrico a 1 día del portafolio: {var_total_portafolio:.2f} pesos")
print(f"VaR Paramétrico a 1 día del portafolio como porcentaje: {var_porcentaje:.2f}%")

#----------------- COMPONENT VAR----------------------
# (VaR * Correlación) que tanto se anula o potencia el riesgo de cada activo por su relación con los demás activos del poortafolio
vector_x_corr = np.dot(vector_var, matriz_correlacion_portafolio)
# Contribución marginal de cada activo al var
base_depurada['Contribucion_Riesgo'] = (vector_x_corr * vector_var) / var_total_portafolio
base_depurada['%_Contribucion'] = (base_depurada['Contribucion_Riesgo'] / var_total_portafolio) * 100

# Ordenar para ver el mayor contribuyente
top_riesgo = base_depurada.sort_values(by='%_Contribucion', ascending=False)
print(top_riesgo[['Activo', 'Posicion_Nominal', '%_Contribucion']])

# ---------------- ESTRESS TESTING INVERSO ----------------------
valor_total_portafolio = base_depurada['Posicion_Nominal'].sum() 
umbral_perdida = valor_total_portafolio * 0.05 #definir un umbral de pérdida del 5%
#buscamos un multiplicador que teniendo en cuenta los mov actuales del mercado, nos de una pérdida del 5% del portafolio
# Calculamos la exposición total al riesgo (Valor en Riesgo por cada 1% de movimiento en el mercado)
exposicion_total = (base_depurada['Posicion_Nominal'] * base_depurada['Duracion_Modificada']).sum()
shock_necesario = umbral_perdida / exposicion_total 
print(f"Shock simultáneo necesario para generar una pérdida del 5%: {shock_necesario:.2%}")

#Dataframe para resultados
resultados = pd.DataFrame({
    "Métrica": ["VaR Portafolio (1 día)", "VaR Portafolio (%)", "Shock Simultáneo para Pérdida del 5%"],
    "Valor": [f"{var_total_portafolio:.2f}", f"{var_porcentaje:.2f}", f"{shock_necesario:.2%}"],
    "expresion": ["COP", "%", "%"]
})

#Exportar resultados a diferentes hojas de Excel
with pd.ExcelWriter("Resultados_VaR.xlsx") as writer:
    base_depurada.to_excel(writer, sheet_name="Base Depurada", index=False)
    matriz_correlacion_portafolio.to_excel(writer, sheet_name="Matriz Correlacion")
    resultados.to_excel(writer, sheet_name="Resultados Resumen", index=False)
    top_riesgo.to_excel(writer, sheet_name="Top Contribuidores", index=False)

