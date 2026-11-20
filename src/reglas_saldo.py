# -*- coding: utf-8 -*-
"""
===============================================================================
 REPORTE DE VALIDACIÓN DE SALDOS – "REGLAS DE SALDO "
===============================================================================
 OBJETIVO
 --------
 Leer UN archivo de remuneraciones (formato ANCHO: 1 fila por trabajador,
 1 columna por concepto) y, por cada trabajador, evaluar las reglas de saldo
 (/S01, /S02, 1D51, 2H43 y tope del 15% de la renta imponible) para indicar
 qué ajuste hay que hacer en el sistema de remuneraciones.

 FLUJO (bloques en orden de ejecución)
 -------------------------------------
   0. Importaciones
   1. Configuración  <- ÚNICO bloque que normalmente se edita
   2. Utilidades     (funciones pequeñas reutilizadas por todo el script)
   3. Carga          (lee el archivo, detecta el encabezado, limpia nombres)
   4. Preparación    (nombres oficiales, quita columnas no necesarias, tipos,
                      validación de filas)
   5. Columnas calculadas (Max_permitido, IT45, voluntarios, etc.)
   6. Motor de reglas (decide la acción y los montos de cada trabajador)
   7. Métricas       (consolidados para QA y gestión)
   8. Exportación    (Excel con colores, lista de ajustes y leyenda)
   9. Programa principal (orquesta todo lo anterior)

===============================================================================
"""

# =============================================================================
# 0. IMPORTACIONES
# =============================================================================
import os
import re
import unicodedata
import warnings
from datetime import datetime

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=FutureWarning)

try:  # En Colab / Jupyter, display() muestra tablas; fuera de ahí usamos print
    from IPython.display import display
except ImportError:  # pragma: no cover
    def display(obj):
        print(obj)


# =============================================================================
# 1. CONFIGURACIÓN
#    Todo lo que puede cambiar entre ejecuciones está aquí. No tocar el resto.
# =============================================================================
VERSION_SCRIPT = "2.1 (QA)"

# --- 1.1 Origen del archivo de datos -----------------------------------------
# 'upload' -> files.upload() de Colab | 'drive' -> Google Drive | 'local' -> ruta fija
MODO_CARGA = "upload"
RUTA_DATOS_DRIVE = "/content/drive/MyDrive/datos_remuneraciones.xlsx"
RUTA_DATOS_LOCAL = "datos_remuneraciones.xlsx"
HOJA_DATOS = 0                      # índice o nombre de la hoja con los datos

# --- 1.2 Parámetros de negocio -----------------------------------------------
PORCENTAJE_TOPE = 0.15              # Max_permitido = 15% * Tope Imponible (RI)
TOLERANCIA_CLP = 1                  # diferencia aceptable al comparar montos ($)
DESCUENTOS_EN_ABSOLUTO = True       # compara magnitudes (ignora el signo)

# --- 1.3 Nombres canónicos de columnas clave ---------------------------------
COL_ID = "N. Personal"
COL_DIAS, COL_PAGO, COL_RI = "dias", "pago", "RI"
COL_S01 = "Sobrepago"                       # /S01
COL_S02 = "Sobrepago Mes Anterior"          # /S02
COL_1D51 = "Devolución Pago Sin Causa"      # 1D51
COL_2H43 = "Compensación Sobrepago"         # 2H43

ALIAS_COLUMNAS = {                          # DEFINIR 'X': alias
    "Días Trab.": COL_DIAS,
    "Líquido pagado": COL_PAGO,
    "Tope Imponible": COL_RI,
}
COLS_OBLIGATORIAS = [COL_ID, COL_DIAS, COL_PAGO, COL_RI]
COLS_MONTO_OPCIONALES = [COL_S01, COL_S02, COL_1D51, COL_2H43]  # si faltan -> 0
COLS_NUMERICAS_EXTRA = [COL_DIAS, COL_PAGO, COL_RI]

# --- 1.4 Simplificación y formato de la salida -------------------------------
# Columnas que NO se necesitan en este reporte (se eliminan al cargar).
# Nota: 'Cod. Ceco' se incluye junto a 'Centro de Costo'; quitarlo de la lista si se requiere.
COLS_A_ELIMINAR = [
    "Rut", "Sociedad", "Cod. Div.", "División", "Cod. Sdiv", "Sub-División",
    "Cod. Ceco", "Centro de Costo", "Cod. Posición", "Posición", "Tipo Contrato",
    "Fecha Ingreso", "Fecha Fin Contrato", "Plan Horario", "Cod. Ine", "Forma Pago",
    "Cod. Banco", "Banco", "Cuenta", "Cod. AFP", "AFP", "Cod. Isapre", "Isapre",
    "Días Lic.", "Días Aus.",
]
EXPORTAR_AUXILIARES = False         # True -> agrega al Resultado las sumas por grupo (trazabilidad QA)
EXPORTAR_COLUMNAS_QA = False        # True -> agrega Casos_Aplicados y Estado al Resultado

COLOR_CREADAS_ENC = "9DC3E6"        # color del ENCABEZADO de las columnas creadas por el script (azul)
COLOR_SALDOS_ENC = "FFD966"         # color del ENCABEZADO de /S01, 2H43, /S02 y 1D51 (amarillo)

# Orden de columnas del Resultado y de Lista_Ajustes (van después de N. Personal y Nombre).
# Las demás columnas del archivo (haberes y descuentos individuales) quedan a continuación.
ORDEN_COLUMNAS = [
    COL_DIAS, COL_RI, COL_PAGO, COL_S01, COL_2H43, COL_S02, COL_1D51,
    "Max_permitido", "Comp_2H43_Calculado", "Nuevo_1D51", "PROX_MES_1D51",
    "Dentro_15_Post_Ajuste", "Dentro_15%", "Dscto_voluntario_total", "IT45",
    "Otros_Obligatorio", "Excede_Tope_Post", "Acciones", "Detalle_Descuentos",
]
FORMATO_MONTO = "#,##0"

# --- 1.5 Grupos de conceptos (transcritos de la hoja REGLAS) ------------------
# --- Grupos transcritos de la hoja REGLAS (typos corregidos) ----------------
G_VOL_IT15 = ["Mutual de Seguros", "Seg Vida Chilena Consol", "Descuento Casino",
              "Descuento Upgrade Celular", "Descuento Teleton"]
G_VOL_IT14 = ["Seg Complementario Salud", "Descuento FALP",
              "Convenio Gym Smartfit", "Seguro de Salud"]
G_TOTAL_VOL = G_VOL_IT15 + G_VOL_IT14
G_VOL_IT45 = ["Prestamo al Personal", "Autorización de Descuento"]
G_OBLIG_IT14 = ["Leasing Los Andes", "Retención Judicial"]
G_OBLIG_IT15 = ["CCAF Los Heroes", "Credito Social 1", "Credito Social Aval",
                "Seguro de Vida CCAF", "Servicios Medicos Fonasa"]
G_OBLIG_IT3213 = ["Ahorro AFP (Cuenta dos)", "Aho Previs Ley 19.768 - B",
                  "Ahorro previsional vol."]
G_ELIMINAR = ["Mutual de Seguros", "Seg Vida Chilena Consol",
              "Descuento Upgrade Celular", "Descuento Teleton", "Descuento FALP",
              "Leasing Los Andes", "Retención Judicial", "CCAF Los Heroes",
              "Credito Social 1", "Credito Social Aval", "Seguro de Vida CCAF",
              "Servicios Medicos Fonasa", "Ahorro AFP (Cuenta dos)",
              "Aho Previs Ley 19.768 - B", "Ahorro previsional vol."]
G_NO_TRANSF = ["Seg Complementario Salud", "Descuento Casino",
               "Convenio Gym Smartfit", "Seguro de Salud"]
G_DENTRO_15 = ["Sobrepago Mes Anterior", "Prestamo al Personal", "Mutual de Seguros",
               "Seg Complementario Salud", "Seg Vida Chilena Consol",
               "Descuento Casino", "Descuento FALP", "Autorización de Descuento",
               "Devolución Pago Sin Causa", "Convenio Gym Smartfit",
               "Seguro de Salud", "Descuento Upgrade Celular"]
G_OTROS_OBLIG = G_OBLIG_IT15 + G_OBLIG_IT14 + G_OBLIG_IT3213
G_BONOS = ["Bono Apoyo", "Bono Beneficio Nacim Hijo", "Bono Complemento de Costo",
           "Bono Enc Satisfaccion", "Bono Especial", "Bono Fijo", "Bono Gestión",
           "Bono Higiene", "Bono Mantención", "Bono Matrimonio", "Bono Nocturno",
           "Bono Promociones", "Bono Trainee", "Bono Traslado", "Bono de Asistencia",
           "Bono de Merma", "Bono de Recepción Camión", "Bono de Reemplazo",
           "Bono de Servicio", "Bono fallecimiento", "Bono meta de Venta",
           "Bono Resultado Operación", "Bono de Zona", "Diferencia Bono",
           "Bono Administracion", "Bono Tratamiento", "Bono Recon. Antiguedad",
           "Bono Vivienda", "Bono Reclutador", "Bono Atención Cliente", "Bono Gráfica"]
G_VARIABLES = ["Horas Extras FT (50%)", "Horas Extras PT (50%)",
               "Horas Extras Domingo FT", "Horas Extras Domingo PT",
               "Recargo 30% Domingo FT", "Recargo 30% Domingo PT",
               "Compensacion Feriado FT", "Comp. Feriado Irrenun. FT",
               "Comp. Feriado Irrenun. PT"]
G_ASIGNACIONES = ["Asignacion Movilizacion", "Asignación Colación",
                  "Asignacion de Caja Fija", "Dif Asignacion de Caja",
                  "Asignaciones Maternales", "Asignaciones Retroactivas",
                  "Asignación de Colación", "Asignación de Zona"]

GRUPOS = {
    "Descto_Vol_IT_15": G_VOL_IT15, "Descto_Vol_IT_14": G_VOL_IT14,
    "Total_Descto_Vol": G_TOTAL_VOL, "Descto_Vol_IT_45": G_VOL_IT45,
    "Descto_Oblig_IT_14": G_OBLIG_IT14, "Descto_Oblig_IT_15": G_OBLIG_IT15,
    "Descto_Oblig_IT_3213": G_OBLIG_IT3213, "Descto_eliminar": G_ELIMINAR,
    "Descto_No_Transf": G_NO_TRANSF, "Dentro_15%": G_DENTRO_15,
    "Otros_Obligatorio": G_OTROS_OBLIG, "Bonos": G_BONOS,
    "Variables": G_VARIABLES, "Asignaciones": G_ASIGNACIONES,
}

# Sumas auxiliares por grupo. Se calculan para el motor de reglas y, por defecto,
# NO se exportan (EXPORTAR_AUXILIARES). La FALP no tiene columna propia: ya forma
# parte de G_VOL_IT14 y se conserva su columna original 'Descuento FALP'.
AUX_GRUPOS = {
    "Vol_IT14": G_VOL_IT14, "Vol_IT15": G_VOL_IT15,
    "Oblig_IT14": G_OBLIG_IT14, "Oblig_IT15": G_OBLIG_IT15, "Oblig_IT3213": G_OBLIG_IT3213,
    "Descto_eliminar_total": G_ELIMINAR, "Descto_No_Transf_total": G_NO_TRANSF,
    "Total_Bonos": G_BONOS, "Total_Variables": G_VARIABLES, "Total_Asignaciones": G_ASIGNACIONES,
}
AUX_COLS = list(AUX_GRUPOS.keys())
QA_COLS = ["Casos_Aplicados", "Estado"]

# --- 1.6 Vocabulario UNIFICADO de acciones -----------------------------------
# La columna 'Acciones' solo usa estas frases (separadas por ' | ').
A_OK = "OK"
A_SOBREPAGO_OK = "Sobrepago OK"
A_COMPENSAR = "Compensar /S02 con 2H43"
A_AJUSTAR_1D51 = "Ajustar 1D51 del mes"
A_TRASP_1D51 = "Traspasar 1D51 a próximo mes"
A_ELIM_TRASP = "Eliminar o traspasar descuentos"
A_ELIM_ATRASOS = "Eliminar atrasos"
A_REV_DESC = "Revisar descuentos"
A_REV_S01S02_SIN_RI = "Revisar: /S01 + /S02 con días > 0 y sin RI (la compensación con 2H43 requiere RI > 0)"
A_REV_S02_DIAS0_SIN_RI = "Revisar: /S02 con días = 0, sin RI y sin otros descuentos que mover (regla no definida)"
A_REV_2H43 = "Revisar: 2H43 distinto de /S02"
ACCIONES_OK = {A_OK, A_SOBREPAGO_OK}     # acciones que NO requieren intervención

# --- 1.7 Diccionario de conceptos (pestaña DEFINICIONES) ----------------------
DICCIONARIO = [  # (código, concepto) transcrito de la pestaña DEFINICIONES
    ('', 'N. Personal'),
    ('', 'Nombre'),
    ('', 'Rut'),
    ('', 'Sociedad'),
    ('', 'Cod. Div.'),
    ('', 'División'),
    ('', 'Cod. Sdiv'),
    ('', 'Sub-División'),
    ('', 'Cod. Ceco'),
    ('', 'Centro de Costo'),
    ('', 'Cod. Posición'),
    ('', 'Posición'),
    ('', 'Tipo Contrato'),
    ('', 'Fecha Ingreso'),
    ('', 'Fecha Fin Contrato'),
    ('', 'Plan Horario'),
    ('', 'Cod. Ine'),
    ('', 'Forma Pago'),
    ('', 'Cod. Banco'),
    ('', 'Banco'),
    ('', 'Cuenta'),
    ('', 'Cod. AFP'),
    ('', 'AFP'),
    ('', 'Cod. Isapre'),
    ('', 'Isapre'),
    ('', 'Días Trab.'),
    ('', 'Días Lic.'),
    ('', 'Días Aus.'),
    ('', 'Tope Imponible'),
    ('', 'Líquido pagado'),
    ('/S01', 'Sobrepago'),
    ('1D33', 'Atrasos'),
    ('1H01', 'Sueldo Base'),
    ('1H02', 'Gratificación'),
    ('1H03', 'Asignacion Movilizacion'),
    ('1H04', 'Asignación Colación'),
    ('1H05', 'Bono Apoyo'),
    ('1H09', 'Bono Beneficio Nacim Hijo'),
    ('1H13', 'Bono Complemento de Costo'),
    ('1H17', 'Bono Enc Satisfaccion'),
    ('1H18', 'Bono Especial'),
    ('1H21', 'Bono Fijo'),
    ('1H23', 'Bono Gestión'),
    ('1H26', 'Bono Higiene'),
    ('1H29', 'Bono Mantención'),
    ('1H30', 'Bono Matrimonio'),
    ('1H34', 'Bono Nocturno'),
    ('1H35', 'Bono Promociones'),
    ('1H39', 'Bono Trainee'),
    ('1H40', 'Bono Traslado'),
    ('1H43', 'Bono de Asistencia'),
    ('1H45', 'Bono de Merma'),
    ('1H46', 'Bono de Recepción Camión'),
    ('1H47', 'Bono de Reemplazo'),
    ('1H48', 'Bono de Servicio'),
    ('1H50', 'Bono fallecimiento'),
    ('1H51', 'Bono meta de Venta'),
    ('1H54', 'Bono Resultado Operación'),
    ('1H59', 'Horas Extras FT (50%)'),
    ('1H60', 'Horas Extras PT (50%)'),
    ('1H61', 'Horas Extras Domingo FT'),
    ('1H62', 'Horas Extras Domingo PT'),
    ('1H63', 'Recargo 30% Domingo FT'),
    ('1H64', 'Recargo 30% Domingo PT'),
    ('1H65', 'Compensacion Feriado FT'),
    ('1H66', 'Comp. Feriado Irrenun. FT'),
    ('1H67', 'Comp. Feriado Irrenun. PT'),
    ('1H70', 'Bono de Zona'),
    ('1H73', 'Diferencia Bono'),
    ('1H75', 'Diferencia Gratificación'),
    ('1H76', 'Diferencia Sueldo Base'),
    ('1H83', 'Asig. Mov Sala Cuna'),
    ('1H84', 'Asignacion de Caja Fija'),
    ('1H85', 'Asig de Caja Variable'),
    ('1H86', 'Bono Administracion'),
    ('1H88', 'Dif Asignacion de Caja'),
    ('1H89', 'Diferencia Colacion'),
    ('1H90', 'Diferencia Movilizacion'),
    ('1H91', 'Diferencia liquida'),
    ('1H99', 'Asignaciones Maternales'),
    ('2H02', 'Asig Familiares Simple'),
    ('2H03', 'Asignaciones Retroactivas'),
    ('2H11', 'Bono Tratamiento'),
    ('2H14', 'Bono Recon. Antiguedad'),
    ('2H17', 'Bono Vivienda'),
    ('2H20', 'Asig.Movili.conectividad'),
    ('2H21', 'Asig. Compens. Sala Cuna'),
    ('2H30', 'Bono Reclutador'),
    ('2H37', 'Asig. Desgaste Vehículo'),
    ('2H42', 'Asignación de Colación'),
    ('2H47', 'Asig. Comp. Sala Cuna (R)'),
    ('2H49', 'Asignación de Zona'),
    ('2H60', 'Bono Atención Cliente'),
    ('2H66', 'Bono Gráfica'),
    ('/301', 'Cotizacion Salud Obligato'),
    ('/302', 'Cotiz salud adicioNoDeduc'),
    ('/303', 'Cotizacion salud adiciona'),
    ('/312', 'Impuesto sobre la renta'),
    ('/S02', 'Sobrepago Mes Anterior'),
    ('1A01', 'Anticipo Sueldo Mensual'),
    ('1D02', 'CCAF Los Heroes'),
    ('1D03', 'Credito Social 1'),
    ('1D05', 'Credito Social Aval'),
    ('1D06', 'Seguro de Vida CCAF'),
    ('1D12', 'Prestamo al Personal'),
    ('1D15', 'Mutual de Seguros'),
    ('1D16', 'Seg Complementario Salud'),
    ('1D18', 'Seg Vida Chilena Consol'),
    ('1D19', 'Servicios Medicos Fonasa'),
    ('1A08', 'Anticipo bono Promociones'),
    ('1D27', 'Leasing Los Andes'),
    ('1D28', 'Dcto. Cuota Sindical'),
    ('1D29', 'Descuento Casino'),
    ('1D37', 'Contrib.san.obligatoriaF'),
    ('1D46', 'Descuento FALP'),
    ('1D48', 'Autorización de Descuento'),
    ('1D51', 'Devolución Pago Sin Causa'),
    ('1D61', 'Convenio Gym Smartfit'),
    ('1D62', 'Seguro de Salud'),
    ('1D63', 'Descuento Upgrade Celular'),
    ('/GMT', 'Retención Judicial'),
    ('FPA', 'Fondo pensiones AFP'),
    ('COMAFP', 'Comisión AFP'),
    ('CU2', 'Ahorro AFP (Cuenta dos)'),
    ('CES', 'Seg Cesantia Trabajador'),
    ('APV-B', 'Aho Previs Ley 19.768 - B'),
    ('APV-VOL', 'Ahorro previsional vol.'),
    ('2H43', 'Compensación Sobrepago'),
    ('', 'Descuento Teleton'),  # agregado: figura en REGLAS pero no estaba en DEFINICIONES
]


# =============================================================================
# 2. UTILIDADES
#    Funciones pequeñas y sin efectos laterales, usadas por los demás bloques.
# =============================================================================
def normalizar(texto) -> str:
    """Normaliza texto para comparar nombres (minúsculas, sin tildes ni ´ ni espacios dobles).

    Por qué: los datos pueden venir con distinta capitalización o tildes; comparar
    normalizado evita fallas silenciosas. Retorna '' si el valor es nulo.
    """
    if texto is None or (not isinstance(texto, str) and pd.isna(texto)):
        return ""
    t = unicodedata.normalize("NFKD", str(texto))
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = re.sub(r"[´`’']", "", t)
    return re.sub(r"\s+", " ", t).strip().lower()


def limpiar_encabezado(valor) -> str:
    """Limpia un nombre de columna: quita espacios y comas finales ('Días Trab.,' -> 'Días Trab.')."""
    return str(valor).strip().rstrip(",").strip()


def _nombres_en_grupos():
    """Nombres normalizados de todos los conceptos usados en los grupos (siempre son montos)."""
    return {normalizar(n) for g in GRUPOS.values() for n in g}


def redondear_clp(x):
    """Redondeo 'comercial' (0.5 sube), como ROUND() de Excel; numpy usa redondeo bancario."""
    return np.floor(x + 0.5)


def _parsear_monto(v) -> float:
    """Convierte un valor individual a float. Maneja '1.234', '123-', '(123)' y vacíos.

    Supone CLP entero: el punto es separador de miles y la coma, decimal.
    """
    if v is None:
        return np.nan
    if isinstance(v, (int, float, np.integer, np.floating)):
        return float(v)
    s = str(v).strip()
    if s in ("", "-", "nan", "None"):
        return np.nan
    negativo = s.endswith("-") or s.startswith("-") or (s.startswith("(") and s.endswith(")"))
    s = s.strip("-()").replace(".", "").replace(",", ".")
    try:
        x = float(s)
    except ValueError:
        return np.nan
    return -x if negativo else x


def a_numero(serie: pd.Series) -> pd.Series:
    """Convierte una Serie a float de forma segura (no convertibles -> NaN)."""
    if pd.api.types.is_numeric_dtype(serie):
        return serie.astype(float)
    return serie.map(_parsear_monto).astype(float)


def suma_grupo(df: pd.DataFrame, nombres: list):
    """Suma por fila las columnas de un grupo que existan en df.

    Retorna (Serie con la suma, lista de nombres del grupo ausentes en los datos).
    Los ausentes cuentan como 0 y se informan en las alertas.
    """
    mapa = {normalizar(c): c for c in df.columns}
    presentes = [mapa[normalizar(n)] for n in nombres if normalizar(n) in mapa]
    faltan = [n for n in nombres if normalizar(n) not in mapa]
    if not presentes:
        return pd.Series(0.0, index=df.index), faltan
    vals = df[presentes]
    if DESCUENTOS_EN_ABSOLUTO:
        vals = vals.abs()
    return vals.sum(axis=1), faltan


# =============================================================================
# 3. CARGA DE ARCHIVOS
# =============================================================================
def obtener_ruta_datos() -> str:
    """Obtiene la ruta del archivo de remuneraciones según MODO_CARGA (upload | drive | local).

    Las reglas y el diccionario ya están dentro del script: solo se necesita el archivo de datos.
    """
    if MODO_CARGA == "local":
        return RUTA_DATOS_LOCAL
    if MODO_CARGA == "drive":
        from google.colab import drive  # solo existe en Colab
        drive.mount("/content/drive")
        return RUTA_DATOS_DRIVE

    from google.colab import files
    print(">> Sube el archivo de REMUNERACIONES (xlsx/csv)")
    subido = files.upload()
    return next(iter(subido))


def cargar_metadatos() -> dict:
    """Construye los metadatos a partir del diccionario embebido (DICCIONARIO).

    Retorna dict con: 'codigo_a_concepto', 'concepto_por_norm' (nombre normalizado ->
    nombre oficial) y 'conceptos_con_codigo' (conceptos que son montos).
    """
    codigo_a_concepto = {c.upper(): n for c, n in DICCIONARIO if c}
    meta = {
        "codigo_a_concepto": codigo_a_concepto,
        "concepto_por_norm": {normalizar(n): n for _, n in DICCIONARIO},
        "conceptos_con_codigo": set(codigo_a_concepto.values()),
    }
    print(f"[Metadatos] {len(DICCIONARIO)} conceptos en el diccionario embebido "
          f"({len(codigo_a_concepto)} con código)")
    return meta


def validar_reglas_vs_diccionario(meta: dict) -> list:
    """Detecta nombres usados en los grupos de REGLAS que no existen en DEFINICIONES.

    Retorna lista de alertas (str).
    """
    alertas = []
    for grupo, nombres in GRUPOS.items():
        for n in nombres:
            if normalizar(n) not in meta["concepto_por_norm"]:
                alertas.append(f"'{n}' (grupo {grupo}) no existe en DEFINICIONES")
    alertas = list(dict.fromkeys(alertas))
    for a in alertas:
        print("   [!]", a)
    return alertas


def cargar_datos(ruta_datos: str) -> pd.DataFrame:
    """Lee el archivo de datos (xlsx/xls/csv) y detecta automáticamente la fila de encabezado.

    Por qué: los reportes exportados suelen traer títulos antes de la tabla. Se busca en
    las primeras 30 filas la que contiene 'N. Personal'. Los nombres de columna se limpian
    (espacios y comas finales).
    """
    ext = os.path.splitext(ruta_datos)[1].lower()
    if ext in (".csv", ".txt"):
        crudo = pd.read_csv(ruta_datos, header=None, sep=None, engine="python", dtype=object)
    else:
        crudo = pd.read_excel(ruta_datos, sheet_name=HOJA_DATOS, header=None, dtype=object)

    fila_header = None
    for i in range(min(30, len(crudo))):
        if any(normalizar(limpiar_encabezado(v)) == normalizar(COL_ID) for v in crudo.iloc[i].tolist()):
            fila_header = i
            break
    if fila_header is None:
        print("   [!] No encontré 'N. Personal' en las primeras 30 filas; uso la fila 1 como encabezado.")
        fila_header = 0

    df = crudo.iloc[fila_header + 1:].copy()
    df.columns = [limpiar_encabezado(c) for c in crudo.iloc[fila_header].tolist()]
    df = df.reset_index(drop=True)
    print(f"[Datos] {len(df):,} filas x {df.shape[1]} columnas (encabezado en fila {fila_header + 1})")
    return df


# =============================================================================
# 4. PREPARACIÓN Y VALIDACIÓN
# =============================================================================
def mapear_columnas(df: pd.DataFrame, meta: dict):
    """Renombra las columnas de datos al nombre oficial del diccionario y aplica los alias.

    Acepta que la columna venga por Concepto ('Mutual de Seguros') o por código ('1D15').
    Luego aplica DEFINIR: 'Días Trab.'->dias, 'Líquido pagado'->pago, 'Tope Imponible'->RI.
    Retorna (df_renombrado, alertas).
    """
    alertas, renombres = [], {}
    for col in df.columns:
        n = normalizar(col)
        if n in meta["concepto_por_norm"]:
            renombres[col] = meta["concepto_por_norm"][n]
        elif str(col).strip().upper() in meta["codigo_a_concepto"]:
            renombres[col] = meta["codigo_a_concepto"][str(col).strip().upper()]
    df = df.rename(columns=renombres).rename(columns=ALIAS_COLUMNAS)

    dup = df.columns[df.columns.duplicated()].tolist()
    if dup:
        alertas.append(f"Columnas duplicadas tras normalizar (se conserva la primera): {sorted(set(dup))}")
        df = df.loc[:, ~df.columns.duplicated()]

    no_reconocidas = [c for c in df.columns if normalizar(c) not in meta["concepto_por_norm"]
                      and c not in (COL_DIAS, COL_PAGO, COL_RI)]
    if no_reconocidas:
        alertas.append(f"{len(no_reconocidas)} columnas no están en el diccionario (se conservan sin usar): "
                       f"{no_reconocidas[:15]}{'...' if len(no_reconocidas) > 15 else ''}")

    faltan = [c for c in COLS_OBLIGATORIAS if c not in df.columns]
    if faltan:
        raise ValueError(f"Faltan columnas obligatorias {faltan}. Columnas disponibles: {list(df.columns)[:40]}")

    for c in COLS_MONTO_OPCIONALES:
        if c not in df.columns:
            df[c] = 0.0
            alertas.append(f"Columna '{c}' no viene en los datos; se asume 0 para todos.")
    return df, alertas


def simplificar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    """Elimina las columnas que no se necesitan en este reporte (lista COLS_A_ELIMINAR).

    Se hace temprano para trabajar con menos datos y para no arrastrar datos personales
    (RUT, banco, cuenta, etc.) al Excel de salida. Nunca se elimina una columna obligatoria.
    """
    objetivo = {normalizar(c) for c in COLS_A_ELIMINAR}
    quitar = [c for c in df.columns if normalizar(c) in objetivo and c not in COLS_OBLIGATORIAS]
    print(f"[Simplificación] {len(quitar)} columnas eliminadas; quedan {df.shape[1] - len(quitar)}")
    return df.drop(columns=quitar)


def limpiar_y_tipar(df: pd.DataFrame, meta: dict):
    """Fija tipos: ID como texto, montos y días como float, nulos numéricos como 0.

    Informa tipos ANTES/DESPUÉS y cuántos valores no se pudieron convertir.
    Retorna (df_tipado, alertas).
    """
    alertas = []
    nombres_grupos = _nombres_en_grupos()
    cols_num = [c for c in df.columns if c in meta["conceptos_con_codigo"]
                or c in COLS_NUMERICAS_EXTRA or normalizar(c) in nombres_grupos]

    print(f"[Tipos ANTES]  object={int((df.dtypes == object).sum())} | "
          f"numéricas={int(df.dtypes.map(pd.api.types.is_numeric_dtype).sum())}")

    df[COL_ID] = df[COL_ID].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    df.loc[df[COL_ID].isin(["nan", "None", ""]), COL_ID] = np.nan

    coaccionados = 0
    for c in cols_num:
        original = df[c]
        convertido = a_numero(original)
        malos = int((convertido.isna() & original.notna() & (original.astype(str).str.strip() != "")).sum())
        if malos:
            coaccionados += malos
            alertas.append(f"'{c}': {malos} valores no numéricos convertidos a 0")
        df[c] = convertido.fillna(0.0)

    print(f"[Tipos DESPUÉS] numéricas tipadas={len(cols_num)} | valores no convertibles->0: {coaccionados}")
    return df, alertas


def validar_datos(df: pd.DataFrame, ctrl: dict):
    """Descarta filas sin ID, elimina duplicados exactos y alerta si hay IDs repetidos.

    - Filas sin 'N. Personal' (p. ej. línea de totales) se descartan.
    - Duplicados EXACTOS (fila completa idéntica) se eliminan.
    - ID repetido con datos distintos NO se elimina: se informa en alertas y en Control_Filas
      (ya no se agrega una columna al Resultado).
    Retorna (df_validado, alertas) y actualiza 'ctrl' con los conteos.
    """
    alertas = []
    n0 = len(df)
    sin_id = df[COL_ID].isna()
    ctrl["descartados_sin_id"] = int(sin_id.sum())
    df = df[~sin_id].copy()

    n_antes = len(df)
    df = df.drop_duplicates()
    ctrl["duplicados_exactos_eliminados"] = n_antes - len(df)

    rep = df.loc[df.duplicated(subset=[COL_ID], keep=False), COL_ID]
    ctrl["ids_repetidos"] = int(len(rep))
    if len(rep):
        alertas.append(f"{rep.nunique()} N. Personal con más de una fila (no se eliminan; revisar): "
                       f"{sorted(rep.unique())[:10]}")

    print(f"[Validación] filas={n0:,} | sin ID descartadas={ctrl['descartados_sin_id']} | "
          f"duplicados exactos eliminados={ctrl['duplicados_exactos_eliminados']} | "
          f"filas con ID repetido={ctrl['ids_repetidos']}")
    if df.empty:
        raise ValueError("No quedaron filas válidas tras la validación; revisa el archivo de datos.")
    return df.reset_index(drop=True), alertas


# =============================================================================
# 5. COLUMNAS CALCULADAS
# =============================================================================
def calcular_columnas(df: pd.DataFrame):
    """Crea las columnas nuevas definidas en REGLAS.

    Max_permitido = 15% * RI (mínimo 0) | Dentro_15% | Dscto_voluntario_total | IT45 |
    Otros_Obligatorio, más las sumas auxiliares por grupo (AUX_GRUPOS) que usa el motor.
    Retorna (df, alertas); la alerta lista, en una sola línea, los conceptos ausentes.
    """
    faltantes = set()

    def sg(nombres):
        s, f = suma_grupo(df, nombres)
        faltantes.update(f)
        return s

    df["Max_permitido"] = np.maximum(redondear_clp(PORCENTAJE_TOPE * df[COL_RI]), 0)
    df["Dentro_15%"] = sg(G_DENTRO_15)
    df["Dscto_voluntario_total"] = sg(G_TOTAL_VOL)      # = Total_Descto_Vol
    df["IT45"] = sg(G_VOL_IT45)
    df["Otros_Obligatorio"] = sg(G_OTROS_OBLIG)
    for col, grupo in AUX_GRUPOS.items():
        df[col] = sg(grupo)

    # Montos clave: magnitud (ignora signo) para comparaciones coherentes
    if DESCUENTOS_EN_ABSOLUTO:
        for c in (COL_S01, COL_S02, COL_1D51, COL_2H43):
            df[c] = df[c].abs()

    alertas = []
    if faltantes:
        alertas.append(f"{len(faltantes)} conceptos de las reglas no vienen en los datos "
                       f"(se asumen 0): {sorted(faltantes)}")
    return df, alertas


# =============================================================================
# 6. MOTOR DE REGLAS
#    Decide, por trabajador: acciones, 2H43 a crear, Nuevo_1D51 (mes actual) y
#    PROX_MES_1D51 (saldo a traspasar). Resumen de casos:
#      CASO 1  días = 0 con descuentos        -> eliminar o traspasar
#      CASO 2  días > 0 y líquido = 0         -> revisar descuentos
#      CASO 3  días < 0 o RI < 0              -> eliminar atrasos
#      CASO 4  días = 0 con /S02 y RI > 0     -> compensar 2H43; PROX = 1D51 + /S02
#      CASO 5A días > 0, solo /S02            \
#      CASO 7/8 días > 0, /S02 + 1D51          > misma lógica del 15% (_ajuste_con_s02)
#      CASO 11 / 11B días > 0, /S01 + /S02     /
#      CASO 5B días > 0, solo 1D51            \  lógica del 15% sobre 1D51
#      CASO 9  días > 0, /S01 + 1D51          /  (_ajuste_solo_1d51)
#      CASO 6  solo /S01                      -> eliminar/traspasar o 'Sobrepago OK'
#      CASO 12 2H43 ya existente              -> validar contra /S02
# =============================================================================
def _ajuste_con_s02(s02, d51, mx, cap):
    """Lógica del 15% cuando hay /S02 (CASOS 5A, 7, 8, 11 y 11B). Solo para días > 0.

    cap = IT45 + descuentos voluntarios (lo que ya consume del tope).
    * /S02 + cap > Max  (CASO 8): 2H43 = /S02; Nuevo_1D51 = Max - cap (mín. 0);
      PROX_MES_1D51 = (/S02 - Nuevo_1D51) + 1D51.
    * /S02 + cap <= Max (CASO 7): /S02 se mantiene dentro del tope. Si /S02 + 1D51 + cap > Max:
      Nuevo_1D51 = Max - /S02 - cap y PROX = 1D51 - Nuevo_1D51; si no, OK.
    Retorna (comp_2H43, nuevo_1D51, prox_1D51, lista_acciones, es_caso8).
    """
    if s02 + cap > mx + TOLERANCIA_CLP:                       # CASO 8
        nuevo = max(mx - cap, 0.0)
        return s02, nuevo, (s02 - nuevo) + d51, [A_COMPENSAR, A_AJUSTAR_1D51, A_TRASP_1D51], True
    if s02 + d51 + cap > mx + TOLERANCIA_CLP:                 # CASO 7
        nuevo = mx - s02 - cap
        return np.nan, nuevo, d51 - nuevo, [A_AJUSTAR_1D51, A_TRASP_1D51], False
    return np.nan, np.nan, np.nan, [A_OK], False


def _ajuste_solo_1d51(d51, mx, cap):
    """Lógica del 15% cuando solo hay 1D51 (CASOS 5B y 9). Solo para días > 0.

    * cap > Max: no cabe nada -> Nuevo_1D51 = 0 y todo el 1D51 pasa al próximo mes.
    * 1D51 + cap <= Max: OK.
    * Si no: Nuevo_1D51 = Max - cap y PROX_MES_1D51 = 1D51 - Nuevo_1D51.
    Retorna (nuevo_1d51, prox_1d51, lista_acciones).
    """
    if cap > mx + TOLERANCIA_CLP:
        return 0.0, d51, [A_AJUSTAR_1D51, A_TRASP_1D51]
    if d51 + cap <= mx + TOLERANCIA_CLP:
        return np.nan, np.nan, [A_OK]
    nuevo = mx - cap
    return nuevo, d51 - nuevo, [A_AJUSTAR_1D51, A_TRASP_1D51]


# Columnas que necesita el motor (se evalúa sobre un DataFrame reducido: más rápido).
COLS_MOTOR = [COL_DIAS, COL_PAGO, COL_RI, COL_S01, COL_S02, COL_1D51, COL_2H43,
              "Max_permitido", "IT45", "Dscto_voluntario_total", "Otros_Obligatorio",
              "Vol_IT14", "Vol_IT15", "Oblig_IT14", "Oblig_IT15", "Oblig_IT3213",
              "Descto_eliminar_total", "Descto_No_Transf_total"]


def evaluar_fila(r: pd.Series) -> pd.Series:
    """Evalúa las reglas para UN trabajador y devuelve acciones, montos y controles.

    Un trabajador puede caer en varios casos (p. ej. CASO 2 y CASO 9); se acumulan las
    acciones. Los montos (2H43, Nuevo_1D51, PROX_MES_1D51) vienen del primer caso numérico
    aplicable (los casos numéricos son excluyentes entre sí).
    Retorna Serie con: Casos_Aplicados, Acciones, Detalle_Descuentos, Estado,
    Comp_2H43_Calculado, Nuevo_1D51, PROX_MES_1D51, Dentro_15_Post_Ajuste, Excede_Tope_Post.
    """
    d, pago, ri = r[COL_DIAS], r[COL_PAGO], r[COL_RI]
    s01, s02, d51, h43 = r[COL_S01], r[COL_S02], r[COL_1D51], r[COL_2H43]
    mx, it45, vol = r["Max_permitido"], r["IT45"], r["Dscto_voluntario_total"]
    elim, no_tr, ob15 = r["Descto_eliminar_total"], r["Descto_No_Transf_total"], r["Oblig_IT15"]
    cap = it45 + vol                       # consumo del tope por voluntarios + IT45
    a, b, c = s01 > 0, s02 > 0, d51 > 0    # hay /S01, /S02, 1D51

    casos, acc, detalle = [], [], ""
    comp = nuevo = prox = np.nan
    asignado = False

    def set_num(cmp_=np.nan, nvo=np.nan, prx=np.nan):
        """Guarda los montos del primer caso numérico aplicable."""
        nonlocal comp, nuevo, prox, asignado
        if not asignado:
            comp, nuevo, prox, asignado = cmp_, nvo, prx, True

    # ---- CASO 3: días negativos o RI negativo -> eliminar atrasos ----------
    if d < 0 or ri < 0:
        casos.append("CASO 3"); acc.append(A_ELIM_ATRASOS)

    # ---- CASO 2: días > 0 y líquido = 0 -> revisar descuentos ---------------
    if d > 0 and pago == 0:
        casos.append("CASO 2"); acc.append(A_REV_DESC)

    # ======================= DÍAS = 0 =======================================
    if d == 0:
        # Qué se traspasa (deuda que se mantiene) y qué se elimina.
        traspasar, eliminar = [], []
        if c:
            traspasar.append("1D51")
        if r["Vol_IT14"] > 0:
            traspasar.append("Vol. IT-14")
        if r["Oblig_IT14"] > 0:
            traspasar.append("Oblig. IT-14")
        if r["Oblig_IT3213"] > 0:
            traspasar.append("Oblig. IT-3213")
        if r["Vol_IT15"] > 0 or r["Oblig_IT15"] > 0:
            eliminar.append("IT-15")
        if it45 > 0:
            eliminar.append("IT-45")

        compensa = b and ri > 0
        otros = [x for x in traspasar if x != "1D51"] + eliminar     # descuentos distintos del 1D51
        # Si se compensa, el 1D51 ya quedó cubierto: solo cuentan los demás descuentos.
        hay_mov = bool(otros) if compensa else bool(traspasar or eliminar)
        if compensa:                       # CASO 4: 2H43 = /S02; 1D51 próximo mes = 1D51 + /S02
            casos.append("CASO 4")
            acc += [A_COMPENSAR, A_TRASP_1D51]
            set_num(s02, 0.0, d51 + s02)
        else:
            if b:                          # CASO 10 con RI = 0: no se compensa; se eliminan/traspasan
                casos.append("CASO 4")     # los descuentos. Solo si no hay ninguno, la regla no existe.
                if not hay_mov:
                    acc.append(A_REV_S02_DIAS0_SIN_RI)
            if c:                          # el 1D51 completo pasa al próximo mes
                set_num(nvo=0.0, prx=d51)
        if hay_mov:                        # CASO 1: eliminar o traspasar descuentos
            casos.append("CASO 1"); acc.append(A_ELIM_TRASP)
            partes = []
            if traspasar:
                partes.append("Traspasar próx. mes: " + ", ".join(traspasar))
            if eliminar:
                partes.append("Eliminar: " + ", ".join(eliminar))
            detalle = " | ".join(partes)
        elif a and not b and not c:        # CASO 6: solo /S01 y sin descuentos que mover
            casos.append("CASO 6"); acc.append(A_SOBREPAGO_OK)

    # ======================= DÍAS > 0 =======================================
    elif d > 0:
        if b:                              # hay /S02
            if a and ri <= 0:              # /S01 + /S02 sin RI: sin regla -> revisar
                casos.append("CASO 11"); acc.append(A_REV_S01S02_SIN_RI)
            else:
                cmp_, nvo, prx, acc_, es8 = _ajuste_con_s02(s02, d51, mx, cap)
                if not a:
                    casos.append("CASO 5A" if not c else ("CASO 8" if es8 else "CASO 7"))
                else:
                    casos.append("CASO 11" if c else "CASO 11B")
                acc += acc_
                set_num(cmp_, nvo, prx)
        elif c:                            # solo 1D51 (con o sin /S01)
            nvo, prx, acc_ = _ajuste_solo_1d51(d51, mx, cap)
            casos.append("CASO 9" if a else "CASO 5B")
            acc += acc_
            set_num(nvo=nvo, prx=prx)
        elif a:                            # CASO 6: solo /S01
            casos.append("CASO 6")
            acc.append(A_ELIM_TRASP if elim > 0 else A_SOBREPAGO_OK)

        # Validación final: con /S01 y obligatorios IT-15, se eliminan/traspasan
        # descuentos para equilibrar el pago y no generar sobrepagos.
        if a and ob15 > 0:
            acc.append(A_ELIM_TRASP)

    # ---- CASO 12: validación de un 2H43 ya existente (modificación manual) --
    if h43 > 0:
        casos.append("CASO 12")
        if abs(h43 - s02) <= TOLERANCIA_CLP:
            acc.append(A_OK if mx + TOLERANCIA_CLP >= d51 + cap else A_REV_DESC)
        else:
            acc.append(A_REV_2H43)

    # ---- Cierre: limpiar acciones y calcular estado ------------------------
    acc = list(dict.fromkeys(acc))                     # quita repetidas conservando el orden
    casos = list(dict.fromkeys(casos))
    reales = [x for x in acc if x not in ACCIONES_OK]
    acc = reales or (acc[:1] if acc else [A_OK])       # si hay acciones reales, se omiten los 'OK'
    if any(x.startswith("Revisar") for x in acc):
        estado = "REVISAR"
    elif all(x in ACCIONES_OK for x in acc):
        estado = "OK"
    else:
        estado = "ACCIÓN"

    # ---- Chequeo post-ajuste del tope 15% (solo días > 0) ------------------
    # El tope efectivo es el mayor entre el 15% y los descuentos NO transferibles
    # (esos se mantienen aunque superen el 15%).
    post = excede = np.nan
    if d > 0:
        s02_post = 0.0 if not np.isnan(comp) else s02
        d51_post = nuevo if not np.isnan(nuevo) else d51
        post = s02_post + cap + d51_post
        excede = "Sí" if post > max(mx, no_tr) + TOLERANCIA_CLP else "No"

    return pd.Series({
        "Casos_Aplicados": "; ".join(casos), "Acciones": " | ".join(acc),
        "Detalle_Descuentos": detalle, "Estado": estado,
        "Comp_2H43_Calculado": comp, "Nuevo_1D51": nuevo, "PROX_MES_1D51": prox,
        "Dentro_15_Post_Ajuste": post, "Excede_Tope_Post": excede,
    })


def aplicar_reglas(df: pd.DataFrame) -> pd.DataFrame:
    """Aplica evaluar_fila a cada trabajador y agrega el resultado como columnas nuevas.

    Row-wise por claridad y trazabilidad (miles de filas: tiempo despreciable).
    """
    res = df[COLS_MOTOR].apply(evaluar_fila, axis=1)
    for col in ("Comp_2H43_Calculado", "Nuevo_1D51", "PROX_MES_1D51", "Dentro_15_Post_Ajuste"):
        res[col] = pd.to_numeric(res[col], errors="coerce")
    return pd.concat([df, res], axis=1)


# =============================================================================
# 7. MÉTRICAS
# =============================================================================
def generar_metricas(df: pd.DataFrame, ctrl: dict) -> dict:
    """Genera las tablas de control y gestión. Retorna dict {nombre_hoja: DataFrame}.

    Resumen_Ajustes  : cuántos trabajadores requieren ajuste, por tipo, y montos.
    Ajustes_por_Dias : concentración de ajustes según días trabajados.
    Motivos_Revision : por qué quedaron en revisión manual.
    Top10_Traspaso   : los 10 mayores saldos que pasan al próximo mes.
    Control_Filas    : conciliación de filas del archivo.
    Nota: un trabajador puede estar en más de un tipo de ajuste (las filas por tipo no suman al total).
    """
    n = len(df)
    ctrl["filas_finales"] = n
    acc = df["Acciones"]
    sin_saldos = df["Casos_Aplicados"] == ""
    req = df["Estado"].isin(["ACCIÓN", "REVISAR"])
    contiene = lambda txt: int(acc.str.contains(txt, regex=False).sum())
    comp, prox = df["Comp_2H43_Calculado"], df["PROX_MES_1D51"]

    filas = [
        ("1. Cobertura", "Trabajadores analizados", n, np.nan),
        ("1. Cobertura", "Sin saldos ni descuentos que mover (nada que validar)", int(sin_saldos.sum()), np.nan),
        ("1. Cobertura", "Con saldos validados OK (sin ajuste)",
         int(((df["Estado"] == "OK") & ~sin_saldos).sum()), np.nan),
        ("1. Cobertura", "REQUIEREN AJUSTE EN SISTEMA (modificar y simular liquidación)", int(req.sum()), np.nan),
        ("1. Cobertura", "   · con acción de ajuste", int((df["Estado"] == "ACCIÓN").sum()), np.nan),
        ("1. Cobertura", "   · con revisión manual", int((df["Estado"] == "REVISAR").sum()), np.nan),
        ("2. Ajustes por tipo", "Compensar /S02 con 2H43 (monto = Σ 2H43 a crear)",
         int((comp > 0).sum()), comp.sum()),
        ("2. Ajustes por tipo", "Ajustar 1D51 del mes (monto = Σ Nuevo_1D51)",
         contiene(A_AJUSTAR_1D51), df["Nuevo_1D51"].sum()),
        ("2. Ajustes por tipo", "Traspasar 1D51 a próximo mes (monto = Σ PROX_MES_1D51)",
         int((prox > 0).sum()), prox.sum()),
        ("2. Ajustes por tipo", "Eliminar o traspasar descuentos", contiene(A_ELIM_TRASP), np.nan),
        ("2. Ajustes por tipo", "Eliminar atrasos", contiene(A_ELIM_ATRASOS), np.nan),
        ("2. Ajustes por tipo", "Revisión manual", contiene("Revisar"), np.nan),
        ("3. Montos y controles", "Sobrepago generado en el mes (/S01)",
         int((df[COL_S01] > 0).sum()), df[COL_S01].sum()),
        ("3. Montos y controles", "Sobrepago de meses anteriores (/S02)",
         int((df[COL_S02] > 0).sum()), df[COL_S02].sum()),
        ("3. Montos y controles", "Exceden el tope tras el ajuste (sin contar no transferibles)",
         int((df["Excede_Tope_Post"] == "Sí").sum()), np.nan),
    ]
    resumen = pd.DataFrame(filas, columns=["Bloque", "Indicador", "Trabajadores", "Monto_CLP"])
    resumen.insert(3, "% del total", (100 * resumen["Trabajadores"] / n).round(1) if n else 0.0)

    tramos = pd.cut(df[COL_DIAS], bins=[-np.inf, -0.001, 0.001, 14.999, 29.999, np.inf],
                    labels=["Días < 0", "Días = 0", "1 a 14 días", "15 a 29 días", "30 o más días"])
    por_dias = (pd.DataFrame({"Tramo": tramos.astype(str), "req": req})
                .groupby("Tramo", sort=False).agg(Trabajadores=("req", "size"), Requieren_ajuste=("req", "sum"))
                .reindex(["Días < 0", "Días = 0", "1 a 14 días", "15 a 29 días", "30 o más días"])
                .fillna(0).astype(int).reset_index())
    por_dias["% con ajuste"] = (100 * por_dias["Requieren_ajuste"]
                                / por_dias["Trabajadores"].replace(0, np.nan)).round(1)

    motivos = (acc.str.split(" | ", regex=False).explode()
               .loc[lambda s: s.str.startswith("Revisar")]
               .value_counts().rename_axis("Motivo").reset_index(name="Trabajadores"))
    if motivos.empty:
        motivos = pd.DataFrame({"Motivo": ["Sin casos en revisión manual"], "Trabajadores": [0]})

    cols_top = [c for c in [COL_ID, "Nombre", COL_DIAS, COL_RI, COL_1D51, "PROX_MES_1D51", "Acciones"]
                if c in df.columns]
    top10 = df.loc[prox > 0].nlargest(10, "PROX_MES_1D51")[cols_top]

    control = pd.DataFrame({
        "Métrica": ["Filas iniciales", "Descartadas (sin N. Personal)", "Duplicados exactos eliminados",
                    "Filas finales", "Filas con N. Personal repetido (no eliminadas)"],
        "Valor": [ctrl["filas_iniciales"], ctrl["descartados_sin_id"], ctrl["duplicados_exactos_eliminados"],
                  ctrl["filas_finales"], ctrl["ids_repetidos"]],
    })

    metricas = {"Resumen_Ajustes": resumen, "Ajustes_por_Dias": por_dias, "Motivos_Revision": motivos,
                "Top10_Traspaso": top10, "Control_Filas": control}
    print("\n===== RESUMEN DE AJUSTES =====")
    display(resumen)
    print("\n===== CONTROL DE FILAS =====")
    display(control)
    return metricas


def construir_lista_ajustes(df: pd.DataFrame) -> pd.DataFrame:
    """Lista de trabajo: solo quienes requieren ajuste (ACCIÓN o REVISAR).

    Incluye columnas vacías para registrar la modificación en el sistema y la simulación
    de la liquidación uno a uno (checklist de QA).
    """
    cols = [COL_ID, "Nombre"] + ORDEN_COLUMNAS          # mismo orden que el Resultado
    lista = df.loc[df["Estado"].isin(["ACCIÓN", "REVISAR"]), [c for c in cols if c in df.columns]].copy()
    for extra in ("Modificado_en_sistema", "Liquidacion_simulada_OK", "Observaciones_QA"):
        lista[extra] = ""
    return lista.reset_index(drop=True)


# =============================================================================
# 8. SALIDA Y EXPORTACIÓN
# =============================================================================
def ordenar_columnas_resultado(df: pd.DataFrame):
    """Ordena el Resultado según ORDEN_COLUMNAS y oculta las columnas internas.

    Orden: N. Personal, Nombre -> ORDEN_COLUMNAS (días, RI, pago, saldos, montos calculados,
    controles y acciones) -> resto de conceptos del archivo. Si EXPORTAR_AUXILIARES o
    EXPORTAR_COLUMNAS_QA están activos, esas columnas van tras Detalle_Descuentos.
    Retorna (df_ordenado, columnas_creadas, columnas_saldos) para colorear los encabezados.
    """
    saldos = [COL_S01, COL_2H43, COL_S02, COL_1D51]
    extras = (AUX_COLS if EXPORTAR_AUXILIARES else []) + (QA_COLS if EXPORTAR_COLUMNAS_QA else [])
    ocultas = (set() if EXPORTAR_AUXILIARES else set(AUX_COLS)) | (set() if EXPORTAR_COLUMNAS_QA else set(QA_COLS))
    base = [c for c in [COL_ID, "Nombre"] + ORDEN_COLUMNAS + extras if c in df.columns]
    resto = [c for c in df.columns if c not in set(base) and c not in ocultas]
    originales = {COL_DIAS, COL_RI, COL_PAGO}                  # vienen del archivo: sin color
    creadas = [c for c in ORDEN_COLUMNAS + extras
               if c in df.columns and c not in originales and c not in saldos]
    return df[base + resto], creadas, [c for c in saldos if c in df.columns]


def _formatear_columnas(ws, columnas, color_enc=None):
    """Colorea SOLO el encabezado de las columnas indicadas y da formato de miles a sus montos."""
    from openpyxl.styles import PatternFill
    relleno_enc = PatternFill("solid", fgColor=color_enc) if color_enc else None
    posicion = {celda.value: celda.column for celda in ws[1]}
    for nombre in columnas:
        j = posicion.get(nombre)
        if j is None:
            continue
        if relleno_enc:
            ws.cell(row=1, column=j).fill = relleno_enc
        for fila in ws.iter_rows(min_row=2, min_col=j, max_col=j):
            celda = fila[0]
            if isinstance(celda.value, (int, float)):
                celda.number_format = FORMATO_MONTO


def _tabla_leyenda(ruta_datos, n_filas):
    """Construye la hoja 'Leyenda': parámetros, colores, vocabulario de acciones y descripción de hojas."""
    filas = [
        ("Versión del script", VERSION_SCRIPT),
        ("Fecha de ejecución", f"{datetime.now():%Y-%m-%d %H:%M}"),
        ("Archivo de datos", os.path.basename(str(ruta_datos))),
        ("Trabajadores analizados", str(n_filas)),
        ("Parámetro: tope", f"{PORCENTAJE_TOPE:.0%} de RI (Max_permitido, mínimo 0)"),
        ("Parámetro: tolerancia", f"${TOLERANCIA_CLP} al comparar montos"),
        ("Encabezado azul", "Columnas CREADAS por el script"),
        ("Encabezado amarillo", "Saldos para comparar: /S01, 2H43, /S02 y 1D51"),
        (A_COMPENSAR, "Crear 2H43 por el mismo monto de /S02 (Comp_2H43_Calculado)"),
        (A_AJUSTAR_1D51, "Crear/modificar el 1D51 del mes con el valor de Nuevo_1D51"),
        (A_TRASP_1D51, "Dejar para el próximo mes el saldo indicado en PROX_MES_1D51"),
        (A_ELIM_TRASP, "Revisar la columna Detalle_Descuentos (qué se traspasa y qué se elimina)"),
        (A_ELIM_ATRASOS, "Eliminar los atrasos (días o RI negativos)"),
        ("Revisar ...", "Caso sin regla definida o con inconsistencia: decisión manual"),
        ("Comp_2H43_Calculado", "2H43 que se debe crear"),
        ("Nuevo_1D51", "1D51 que debe quedar este mes"),
        ("PROX_MES_1D51", "1D51 que se debe cargar el próximo mes"),
        ("Dentro_15_Post_Ajuste", "Suma /S02 + IT45 + voluntarios + 1D51 después del ajuste (días > 0)"),
        ("Excede_Tope_Post", "Sí si, tras el ajuste, supera el mayor entre el 15% y los descuentos no transferibles"),
        ("Hoja Lista_Ajustes", "Trabajadores a modificar en el sistema, con columnas de checklist de simulación"),
        ("Hojas de resumen", "Resumen_Ajustes, Ajustes_por_Dias, Motivos_Revision, Top10_Traspaso, Control_Filas"),
    ]
    return pd.DataFrame(filas, columns=["Elemento", "Descripción"])


def exportar_resultado(df_res, cols_creadas, cols_saldos, lista, metricas, alertas,
                       ruta_datos="", ruta_salida=None) -> str:
    """Exporta el Excel (Resultado, Lista_Ajustes, resúmenes, Alertas, Leyenda) y lo descarga en Colab.

    Aplica formato: encabezado en negrita, filtros, paneles congelados, anchos de columna y
    colores SOLO en el encabezado (azul = columnas creadas, amarillo = saldos). Retorna la ruta del archivo.
    """
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    ruta_salida = ruta_salida or f"Reporte_Saldos_{datetime.now():%Y%m%d_%H%M}.xlsx"
    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as w:
        df_res.to_excel(w, sheet_name="Resultado", index=False)
        lista.to_excel(w, sheet_name="Lista_Ajustes", index=False)
        for nombre, t in metricas.items():
            t.to_excel(w, sheet_name=nombre, index=False)
        pd.DataFrame({"Alertas": alertas or ["Sin alertas"]}).to_excel(w, sheet_name="Alertas", index=False)
        _tabla_leyenda(ruta_datos, len(df_res)).to_excel(w, sheet_name="Leyenda", index=False)

        for ws in w.book.worksheets:                       # formato común a todas las hojas
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for j, celda in enumerate(ws[1], start=1):
                celda.font = Font(bold=True)
                ws.column_dimensions[get_column_letter(j)].width = min(max(len(str(celda.value or "")) + 2, 10), 38)
        for nombre_hoja in ("Resultado", "Lista_Ajustes"):  # colores de comparación
            ws = w.book[nombre_hoja]
            _formatear_columnas(ws, [COL_DIAS, COL_RI, COL_PAGO])            # solo formato de miles
            _formatear_columnas(ws, cols_creadas, COLOR_CREADAS_ENC)
            _formatear_columnas(ws, cols_saldos, COLOR_SALDOS_ENC)
        ws = w.book["Leyenda"]                              # ancho y color de muestra en la leyenda
        ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 30, 105
        for fila in ws.iter_rows(min_row=2, max_col=1):
            if fila[0].value == "Encabezado azul":
                fila[0].fill = PatternFill("solid", fgColor=COLOR_CREADAS_ENC)
            elif fila[0].value == "Encabezado amarillo":
                fila[0].fill = PatternFill("solid", fgColor=COLOR_SALDOS_ENC)
        w.book["Alertas"].column_dimensions["A"].width = 140
        for fila in w.book["Alertas"].iter_rows(min_row=2, max_col=1):
            fila[0].alignment = Alignment(wrap_text=True, vertical="top")

    print(f"\n[Exportación] Archivo generado: {ruta_salida}")
    try:
        from google.colab import files
        files.download(ruta_salida)
    except ImportError:
        print("   (Fuera de Colab: el archivo quedó en el directorio actual.)")
    return ruta_salida


# =============================================================================
# 9. PROGRAMA PRINCIPAL
# =============================================================================
def main():
    """Orquesta el flujo: cargar -> normalizar -> simplificar -> tipar -> validar ->
    calcular -> reglas -> métricas -> exportar. Retorna el DataFrame completo (útil para QA)."""
    ctrl, alertas = {}, []

    ruta_datos = obtener_ruta_datos()
    meta = cargar_metadatos()
    alertas += validar_reglas_vs_diccionario(meta)

    df = cargar_datos(ruta_datos)
    ctrl["filas_iniciales"] = len(df)

    df, a = mapear_columnas(df, meta);          alertas += a
    df = simplificar_columnas(df)
    df, a = limpiar_y_tipar(df, meta);          alertas += a
    df, a = validar_datos(df, ctrl);            alertas += a
    df, a = calcular_columnas(df);              alertas += a
    df = aplicar_reglas(df)

    metricas = generar_metricas(df, ctrl)
    lista = construir_lista_ajustes(df)
    df_res, cols_creadas, cols_saldos = ordenar_columnas_resultado(df)

    print("\n===== ALERTAS =====")
    for a in alertas:
        print("  -", a)

    print("\n===== VISTA PREVIA =====")
    vista = [COL_ID, COL_DIAS, COL_RI, "Max_permitido", COL_S01, COL_S02, COL_1D51,
             "Acciones", "Comp_2H43_Calculado", "Nuevo_1D51", "PROX_MES_1D51"]
    display(df_res[[c for c in vista if c in df_res.columns]].head(10))

    exportar_resultado(df_res, cols_creadas, cols_saldos, lista, metricas, alertas, ruta_datos)
    return df


if __name__ == "__main__":
    resultado = main()
