"""
validacion.py
=============
COMPROBACIÓN DE TODA LA CALCULADORA CONTRA EL EXCEL DEL PROFESOR

¿Qué hace este archivo?
    Corre nuestros cálculos con casos que capturamos en el simulador de Excel y
    compara los resultados. Tiene dos partes:

    PARTE 1. TABLA DE AMORTIZACIÓN (amortizacion.py)
        11 escenarios: meses que dura el crédito, mensualidad del mes 1, saldo al
        final del mes 12, total de intereses y total pagado. Incluyen prepagos
        (reduciendo plazo y pago), penalizaciones y aportaciones del Infonavit.

    PARTE 2. CÁLCULO DEL CRÉDITO (credito.py + amortizacion.py + indicadores.py)
        18 casos que parten de lo que captura el usuario: crédito máximo, crédito
        de la tabla, gastos, enganche, inversión, ingreso mínimo, mensualidad,
        CAT y avisos. Cubren los 7 productos, los 3 modos de cálculo (valor,
        ingresos y monto), los dos esquemas y las campañas.

¿Por qué existe?
    El criterio de "replicabilidad" pide demostrar que la calculadora da los
    mismos resultados que el Excel. Este archivo es esa evidencia, y se vuelve a
    correr cada vez que cambiemos el código para asegurarnos de no romper nada.

¿De dónde salen los valores "excel"?
    Se capturaron los datos de cada caso en el simulador con una fecha fija (el
    Excel usa la fecha de hoy y eso cambia los días de cada mes) y se copiaron
    sus resultados.

Tolerancia:
    Hasta 1 centavo (0.01) en montos y 0.0000001 % en el CAT: solo diferencias
    de redondeo de la computadora. Cualquier diferencia mayor es un error.

Cómo correrlo:
    python validacion.py
"""

from datetime import date

import amortizacion
import credito
import indicadores

TOLERANCIA = 0.01            # un centavo
TOLERANCIA_CAT = 1e-9        # el CAT es un porcentaje: se pide más precisión


# =============================================================================
# PARTE 1. TABLA DE AMORTIZACIÓN
# =============================================================================
# -----------------------------------------------------------------------------
# ESCENARIOS
# Cada escenario es un diccionario con:
#   "nombre": descripción
#   "datos":  lo que recibe generar_tabla (más el sueldo para la aportación Infonavit)
#   "excel":  los resultados que dio el simulador de Excel
# -----------------------------------------------------------------------------
ESCENARIOS_TABLA = [
    {
        "nombre": "Escenario 1: Tradicional, Pagos Oportunos 20 años",
        "datos": {
            "credito": 1_900_000,
            "tasa_inicial": 0.13,
            "pago_al_millar": 11.86,
            "esquema": "Pagos Oportunos",
            "fecha_calculo": date(2026, 11, 18),
            "valor_seguro_danos": 2_000_000,
            "incremento_anual": 0.0,
            "piso_valora": 0.12,
            "seguros_externos": False,
            "prepagos": {},
            "cobra_penalizacion": False,
            "tasa_penalizacion": 0.03,
            "reducir": "plazo",
            "sueldo": 0,
            "producto": 1,
        },
        "excel": {"meses": 212, "mensualidad": 24294, "saldo_mes_12": 1878747.81, "intereses": 2857262.94, "total_pagado": 5052255.72},
    },
    {
        "nombre": "Escenario 2: Tradicional, Valora 20 años, prepago 50,000 (mes 26) con penalización",
        "datos": {
            "credito": 1_900_000,
            "tasa_inicial": 0.12,
            "pago_al_millar": 10.36,
            "esquema": "Valora",
            "fecha_calculo": date(2026, 11, 18),
            "valor_seguro_danos": 2_000_000,
            "incremento_anual": 0.02,
            "piso_valora": 0.12,
            "seguros_externos": False,
            "prepagos": {26: 50_000},
            "cobra_penalizacion": True,
            "tasa_penalizacion": 0.03,
            "reducir": "plazo",
            "sueldo": 0,
            "producto": 1,
        },
        "excel": {"meses": 198, "mensualidad": 21444, "saldo_mes_12": 1894649.62, "intereses": 2694399.5, "total_pagado": 4878239.71},
    },
    {
        "nombre": "Valora nivel 3 (tasa baja), 15 años, prepago 100,000 (mes 30) reduciendo el pago, corte día 3",
        "datos": {
            "credito": 2_700_000,
            "tasa_inicial": 0.13,
            "pago_al_millar": 12.22,
            "esquema": "Valora",
            "fecha_calculo": date(2026, 11, 5),
            "valor_seguro_danos": 3_000_000,
            "incremento_anual": 0.014,
            "piso_valora": 0.12,
            "seguros_externos": False,
            "prepagos": {30: 100_000},
            "cobra_penalizacion": False,
            "tasa_penalizacion": 0.03,
            "reducir": "pago",
            "sueldo": 0,
            "producto": 1,
        },
        "excel": {"meses": 162, "mensualidad": 35544, "saldo_mes_12": 2657403.52, "intereses": 3042815.29, "total_pagado": 6066623.72},
    },
    {
        "nombre": "Caso A: Apoyo Infonavit, sueldo 30,000",
        "datos": {
            "credito": 1_900_000,
            "tasa_inicial": 0.13,
            "pago_al_millar": 11.86,
            "esquema": "Pagos Oportunos",
            "fecha_calculo": date(2026, 11, 18),
            "valor_seguro_danos": 2_000_000,
            "incremento_anual": 0.0,
            "piso_valora": 0.12,
            "seguros_externos": False,
            "prepagos": {},
            "cobra_penalizacion": False,
            "tasa_penalizacion": 0,
            "reducir": "plazo",
            "sueldo": 30_000,
            "producto": 2,
        },
        "excel": {"meses": 169, "mensualidad": 24294, "saldo_mes_12": 1859723.57, "intereses": 2200070.57, "total_pagado": 4022805.86},
    },
    {
        "nombre": "Caso B: Pasivo con Apoyo Infonavit, aportación con tope",
        "datos": {
            "credito": 2_375_000,
            "tasa_inicial": 0.13,
            "pago_al_millar": 11.86,
            "esquema": "Pagos Oportunos",
            "fecha_calculo": date(2026, 11, 18),
            "valor_seguro_danos": 2_500_000,
            "incremento_anual": 0.0,
            "piso_valora": 0.12,
            "seguros_externos": False,
            "prepagos": {},
            "cobra_penalizacion": False,
            "tasa_penalizacion": 0.03,
            "reducir": "plazo",
            "sueldo": 100_000,
            "producto": 5,
        },
        "excel": {"meses": 136, "mensualidad": 30367.5, "saldo_mes_12": 2291897.5, "intereses": 2149364.82, "total_pagado": 4054427.24},
    },
    {
        "nombre": "Caso C: Cofinavit",
        "datos": {
            "credito": 1_600_000,
            "tasa_inicial": 0.13,
            "pago_al_millar": 11.86,
            "esquema": "Pagos Oportunos",
            "fecha_calculo": date(2026, 11, 18),
            "valor_seguro_danos": 2_000_000,
            "incremento_anual": 0.0,
            "piso_valora": 0.12,
            "seguros_externos": False,
            "prepagos": {},
            "cobra_penalizacion": False,
            "tasa_penalizacion": 0,
            "reducir": "plazo",
            "sueldo": 0,
            "producto": 3,
        },
        "excel": {"meses": 212, "mensualidad": 20556, "saldo_mes_12": 1582103.42, "intereses": 2406116.16, "total_pagado": 4275284.82},
    },
    {
        "nombre": "Caso E: CI+CB Cofinavit",
        "datos": {
            "credito": 1_260_000,
            "tasa_inicial": 0.13,
            "pago_al_millar": 11.86,
            "esquema": "Pagos Oportunos",
            "fecha_calculo": date(2026, 11, 18),
            "valor_seguro_danos": 2_000_000,
            "incremento_anual": 0.0,
            "piso_valora": 0.12,
            "seguros_externos": False,
            "prepagos": {},
            "cobra_penalizacion": False,
            "tasa_penalizacion": 0,
            "reducir": "plazo",
            "sueldo": 0,
            "producto": 7,
        },
        "excel": {"meses": 212, "mensualidad": 16319.6, "saldo_mes_12": 1245906.44, "intereses": 1894816.48, "total_pagado": 3394717.79},
    },
    {
        "nombre": "Caso G: Pasivo con gastos financiados",
        "datos": {
            "credito": 1_590_000,
            "tasa_inicial": 0.13,
            "pago_al_millar": 11.86,
            "esquema": "Pagos Oportunos",
            "fecha_calculo": date(2026, 11, 18),
            "valor_seguro_danos": 2_000_000,
            "incremento_anual": 0.0,
            "piso_valora": 0.12,
            "seguros_externos": False,
            "prepagos": {},
            "cobra_penalizacion": False,
            "tasa_penalizacion": 0.03,
            "reducir": "plazo",
            "sueldo": 0,
            "producto": 4,
        },
        "excel": {"meses": 212, "mensualidad": 20431.4, "saldo_mes_12": 1572215.27, "intereses": 2391077.94, "total_pagado": 4249385.78},
    },
    {
        "nombre": "CI+CB, 10 años, campaña WM, seguros externos, prepago 80,000 reduciendo el pago",
        "datos": {
            "credito": 1_260_000,
            "tasa_inicial": 0.095,
            "pago_al_millar": 13.02,
            "esquema": "Pagos Oportunos",
            "fecha_calculo": date(2026, 11, 18),
            "valor_seguro_danos": 2_000_000,
            "incremento_anual": 0.0,
            "piso_valora": 0.095,
            "seguros_externos": True,
            "prepagos": {12: 80_000},
            "cobra_penalizacion": True,
            "tasa_penalizacion": 0,
            "reducir": "pago",
            "sueldo": 0,
            "producto": 7,
        },
        "excel": {"meses": 117, "mensualidad": 16405.2, "saldo_mes_12": 1101064.6, "intereses": 616280.12, "total_pagado": 1876280.12},
    },
    {
        "nombre": "Liquidez, 7 años, Vinculación, valor destructible, prepago 200,000 con penalización",
        "datos": {
            "credito": 1_900_000,
            "tasa_inicial": 0.1075,
            "pago_al_millar": 17.1,
            "esquema": "Pagos Oportunos",
            "fecha_calculo": date(2027, 2, 2),
            "valor_seguro_danos": 1_200_000,
            "incremento_anual": 0.0,
            "piso_valora": 0.1075,
            "seguros_externos": False,
            "prepagos": {16: 200_000},
            "cobra_penalizacion": True,
            "tasa_penalizacion": 0.03,
            "reducir": "plazo",
            "sueldo": 0,
            "producto": 6,
        },
        "excel": {"meses": 73, "mensualidad": 34002, "saldo_mes_12": 1707623.09, "intereses": 647351.55, "total_pagado": 2625497.6},
    },
    {
        "nombre": "Apoyo Infonavit, Valora nivel 2, sueldo 60,000, prepago 150,000 reduciendo el pago",
        "datos": {
            "credito": 2_375_000,
            "tasa_inicial": 0.125,
            "pago_al_millar": 10.84,
            "esquema": "Valora",
            "fecha_calculo": date(2026, 11, 18),
            "valor_seguro_danos": 2_500_000,
            "incremento_anual": 0.017,
            "piso_valora": 0.12,
            "seguros_externos": False,
            "prepagos": {36: 150_000},
            "cobra_penalizacion": False,
            "tasa_penalizacion": 0,
            "reducir": "pago",
            "sueldo": 60_000,
            "producto": 2,
        },
        "excel": {"meses": 162, "mensualidad": 27945, "saldo_mes_12": 2328588.44, "intereses": 2537612.45, "total_pagado": 4605996.61},
    },
    {
        "nombre": "Pasivo con apoyo (producto 5), sueldo 100,000, reducir pago sin prepagos",
        "datos": {
            "credito": 2_375_000, "tasa_inicial": 0.13, "pago_al_millar": 11.86,
            "esquema": "Pagos Oportunos", "fecha_calculo": date(2026, 11, 21),
            "valor_seguro_danos": 2_500_000, "reducir": "pago", "producto": 5,
            "sueldo": 100_000,
        },
        "excel": {"meses": 204, "mensualidad": 30367.5, "saldo_mes_12": 2295196.564878,
                  "intereses": 2830251.493218, "total_pagado": 4372641.551331},
    },
]


def resultados_tabla(datos):
    """Corre nuestra tabla y regresa los mismos 5 resultados que guardamos del Excel."""
    datos = dict(datos)                     # copia, para no modificar el original
    sueldo = datos.pop("sueldo")            # .pop() saca el sueldo del diccionario
    aportacion = amortizacion.aportacion_bimestral_inicial(sueldo) if sueldo else 0.0

    # **datos "desempaca" el diccionario: cada llave se manda como argumento con su nombre.
    tabla = amortizacion.generar_tabla(**datos, aportacion_bimestral=aportacion)

    return {
        "meses": len(tabla),
        "mensualidad": tabla["Pago total"].iloc[0],
        "saldo_mes_12": tabla["Saldo final"].iloc[11],
        "intereses": tabla["Interés"].sum(),
        "total_pagado": tabla["Pago total"].sum(),
    }


# =============================================================================
# PARTE 2. CÁLCULO DEL CRÉDITO (desde lo que captura el usuario)
# "datos" = lo que el usuario captura; "excel" = lo que dio el Excel.
# =============================================================================
# -----------------------------------------------------------------------------
# Cada caso: "datos" = lo que el usuario captura; "excel" = lo que dio el Excel.
# -----------------------------------------------------------------------------
ESCENARIOS_CREDITO = [
    {
        "nombre": 'Escenario 1: Tradicional, PO 20 años',
        "datos": {'producto': 1, 'monto': 2000000, 'modo': 'valor'},
        "fecha": '2026-11-18',
        "excel": {
            'credito_maximo': 1900000,
            'credito_para_tabla': 1900000,
            'gastos_originacion': 129500,
            'enganche': 100000,
            'inversion_cliente': 229500,
            'ingreso_minimo': 40970.909091,
            'mensualidad': 24294,
            'cat': 0.151295024057,
        },
        "avisos": ['Línea máxima a otorgar puede ser menor por perfil del cliente'],
    },
    {
        "nombre": 'Escenario 2: Tradicional, Valora 20 años, pago anticipado',
        "datos": {'producto': 1, 'monto': 2000000, 'modo': 'valor', 'esquema': 'Valora', 'tipo_comision': 'Pago anticipado'},
        "fecha": '2026-11-18',
        "excel": {
            'credito_maximo': 1900000,
            'credito_para_tabla': 1900000,
            'gastos_originacion': 105750,
            'enganche': 100000,
            'inversion_cliente': 205750,
            'ingreso_minimo': 35789.090909,
            'mensualidad': 21444,
            'cat': 0.142481782323,
        },
        "avisos": [],
    },
    {
        "nombre": 'Caso A: Apoyo Infonavit, sueldo 30,000',
        "datos": {'producto': 2, 'monto': 2000000, 'modo': 'valor', 'sueldo_cotizacion': 30000},
        "fecha": '2026-11-18',
        "excel": {
            'credito_maximo': 1900000,
            'credito_para_tabla': 1900000,
            'gastos_originacion': 120000,
            'enganche': 100000,
            'inversion_cliente': 220000,
            'ingreso_minimo': 34667.692308,
            'mensualidad': 24294,
            'cat': 0.150234711861,
        },
        "avisos": [],
    },
    {
        "nombre": 'Caso B: Pasivo con apoyo, sueldo al tope',
        "datos": {'producto': 5, 'monto': 2500000, 'modo': 'valor', 'sueldo_cotizacion': 100000},
        "fecha": '2026-11-18',
        "excel": {
            'credito_maximo': 2375000,
            'credito_para_tabla': 2375000,
            'gastos_originacion': 125000,
            'enganche': 0,
            'inversion_cliente': 125000,
            'ingreso_minimo': 43334.615385,
            'mensualidad': 30367.5,
            'cat': 0.148033997802,
        },
        "avisos": [],
    },
    {
        "nombre": 'Caso C: Cofinavit (titular)',
        "datos": {'producto': 3, 'monto': 2000000, 'modo': 'valor', 'ssv_titular': 100000, 'credito_infonavit_titular': 300000},
        "fecha": '2026-11-18',
        "excel": {
            'credito_maximo': 1600000,
            'credito_para_tabla': 1600000,
            'gastos_originacion': 126750,
            'enganche': 0,
            'inversion_cliente': 126750,
            'ingreso_minimo': 29193.846154,
            'mensualidad': 20556,
            'cat': 0.151362519184,
        },
        "avisos": [],
    },
    {
        "nombre": 'Caso D: Cofinavit (titular y cónyuge)',
        "datos": {'producto': 3, 'monto': 3000000, 'modo': 'valor', 'ssv_titular': 80000, 'ssv_conyuge': 60000, 'credito_infonavit_titular': 250000, 'credito_infonavit_conyuge': 150000},
        "fecha": '2026-11-18',
        "excel": {
            'credito_maximo': 2460000,
            'credito_para_tabla': 2460000,
            'gastos_originacion': 188650,
            'enganche': 0,
            'inversion_cliente': 188650,
            'ingreso_minimo': 44885.538462,
            'mensualidad': 31581.6,
            'cat': 0.151151683129,
        },
        "avisos": [],
    },
    {
        "nombre": 'Caso E: CI+CB',
        "datos": {'producto': 7, 'monto': 2000000, 'modo': 'valor', 'actividad': 'Independiente', 'ssv_titular': 100000, 'ssv_conyuge': 50000},
        "fecha": '2026-11-18',
        "excel": {
            'credito_maximo': 1260000,
            'credito_para_tabla': 1260000,
            'gastos_originacion': 137700,
            'enganche': 50000,
            'inversion_cliente': 187700,
            'ingreso_minimo': 27170.181818,
            'mensualidad': 16319.6,
            'cat': 0.154364833334,
        },
        "avisos": [],
    },
    {
        "nombre": 'Caso F: CI+CB con tope Infonavit',
        "datos": {'producto': 7, 'monto': 4000000, 'modo': 'valor', 'actividad': 'Independiente', 'ssv_titular': 200000},
        "fecha": '2026-11-18',
        "excel": {
            'credito_maximo': 2714554.4,
            'credito_para_tabla': 2714554.4,
            'gastos_originacion': 271495.298,
            'enganche': 200000,
            'inversion_cliente': 471495.298,
            'ingreso_minimo': 58535.663971,
            'mensualidad': 35063.347824,
            'cat': 0.153671920493,
        },
        "avisos": [],
    },
    {
        "nombre": 'Caso G: Pasivo con gastos financiados',
        "datos": {'producto': 4, 'monto': 2000000, 'modo': 'valor', 'credito_solicitado': 1500000},
        "fecha": '2026-11-18',
        "excel": {
            'credito_maximo': 1900000,
            'credito_para_tabla': 1590000,
            'gastos_originacion': 10000,
            'enganche': 0,
            'inversion_cliente': 10000,
            'ingreso_minimo': 27369.230769,
            'mensualidad': 20431.4,
            'cat': 0.149072690881,
        },
        "avisos": ['Línea máxima a otorgar puede ser menor por perfil del cliente'],
    },
    {
        "nombre": 'Liquidez, Vinculación, 7 años, valor destructible',
        "datos": {'producto': 6, 'monto': 2000000, 'modo': 'valor', 'campana': 'Vinculación', 'plazo': 7, 'tipo_comision': 'Pago anticipado', 'valor_destructible': 1200000},
        "fecha": '2027-02-02',
        "excel": {
            'credito_maximo': 1900000,
            'credito_para_tabla': 1900000,
            'gastos_originacion': 105750,
            'enganche': 100000,
            'inversion_cliente': 105750,
            'ingreso_minimo': 59072.727273,
            'mensualidad': 34002,
            'cat': 0.124318350543,
        },
        "avisos": ['Línea máxima a otorgar puede ser menor por perfil del cliente'],
    },
    {
        "nombre": 'Modo ingresos: 50,000 al mes',
        "datos": {'producto': 1, 'monto': 50000, 'modo': 'ingresos'},
        "fecha": '2026-11-20',
        "excel": {
            'credito_maximo': 2318718.381113,
            'credito_para_tabla': 2318718.381113,
            'gastos_originacion': 158021.789296,
            'enganche': 122037.809532,
            'inversion_cliente': 280059.598828,
            'ingreso_minimo': 50000,
            'mensualidad': 29647.865448,
            'cat': 0.151293446126,
        },
        "avisos": ['Línea máxima a otorgar puede ser menor por perfil del cliente'],
    },
    {
        "nombre": 'Modo monto: Valora nivel 2, 15 años',
        "datos": {'producto': 1, 'monto': 1500000, 'modo': 'monto', 'esquema': 'Valora', 'plazo': 15, 'nivel': '2'},
        "fecha": '2026-11-20',
        "excel": {
            'credito_maximo': 1500000,
            'credito_para_tabla': 1500000,
            'gastos_originacion': 103447.368421,
            'enganche': 78947.368421,
            'inversion_cliente': 182394.736842,
            'ingreso_minimo': 32127.272727,
            'mensualidad': 19059.473684,
            'cat': 0.149168026287,
        },
        "avisos": [],
    },
    {
        "nombre": 'Modo ingresos: Valora nivel R (regla 25 veces el ingreso)',
        "datos": {'producto': 1, 'monto': 60000, 'modo': 'ingresos', 'esquema': 'Valora', 'nivel': 'R'},
        "fecha": '2026-11-20',
        "excel": {
            'credito_maximo': 1500000,
            'credito_para_tabla': 1500000,
            'gastos_originacion': 189781.036322,
            'enganche': 1700620.726444,
            'inversion_cliente': 1890401.762766,
            'ingreso_minimo': 33081.818182,
            'mensualidad': 20087.192425,
            'cat': 0.168051585166,
        },
        "avisos": ['Línea máxima a otorgar puede ser menor por perfil del cliente'],
    },
    {
        "nombre": 'Pasivo en modo ingresos con gastos financiados',
        "datos": {'producto': 4, 'monto': 50000, 'modo': 'ingresos'},
        "fecha": '2026-11-20',
        "excel": {
            'credito_maximo': 2740303.541315,
            'credito_para_tabla': 2904721.753794,
            'gastos_originacion': -20191.710304,
            'enganche': 0,
            'inversion_cliente': -20191.710304,
            'ingreso_minimo': 50000,
            'mensualidad': 37087.037366,
            'cat': 0.147732241705,
        },
        "avisos": ['Línea máxima a otorgar puede ser menor por perfil del cliente'],
    },
    {
        "nombre": 'Avisos: casa y crédito debajo del mínimo',
        "datos": {'producto': 1, 'monto': 300000, 'modo': 'valor', 'credito_solicitado': 200000},
        "fecha": '2026-11-20',
        "excel": {
            'credito_maximo': 285000,
            'credito_para_tabla': 200000,
            'gastos_originacion': 19400,
            'enganche': 100000,
            'inversion_cliente': 119400,
            'ingreso_minimo': 10000,
            'mensualidad': 2585,
            'cat': 0.154986593674,
        },
        "avisos": ['Valor de vivienda debajo del mínimo', 'Crédito debajo del mínimo', 'Línea máxima a otorgar puede ser menor por perfil del cliente'],
    },
    {
        "nombre": 'Aviso: monto arriba de 15 millones',
        "datos": {'producto': 1, 'monto': 16000000, 'modo': 'monto'},
        "fecha": '2026-11-20',
        "excel": {
            'credito_maximo': 16000000,
            'credito_para_tabla': 16000000,
            'gastos_originacion': 1079105.263158,
            'enganche': 842105.263158,
            'inversion_cliente': 1921210.526316,
            'ingreso_minimo': 345018.181818,
            'mensualidad': 204581.052632,
            'cat': 0.151143051286,
        },
        "avisos": ['El crédito solicitado debe ser igual al máximo'],
    },
    {
        "nombre": 'Cofinavit con % de gastos y gastos exactos',
        "datos": {'producto': 3, 'monto': 2000000, 'modo': 'valor', 'ssv_titular': 100000, 'credito_infonavit_titular': 300000, 'porcentaje_impuestos': 0.04, 'porcentaje_honorarios': 0.015, 'gastos_exactos_infonavit': 50000},
        "fecha": '2026-11-20',
        "excel": {
            'credito_maximo': 1600000,
            'credito_para_tabla': 1600000,
            'gastos_originacion': 97750,
            'enganche': 0,
            'inversion_cliente': 97750,
            'ingreso_minimo': 29193.846154,
            'mensualidad': 20556,
            'cat': 0.151362519184,
        },
        "avisos": [],
    },
    {
        "nombre": 'CI+CB con subcuenta grande',
        "datos": {'producto': 7, 'monto': 600000, 'modo': 'valor', 'actividad': 'Independiente', 'ssv_titular': 150000, 'ssv_conyuge': 20000},
        "fecha": '2026-11-20',
        "excel": {
            'credito_maximo': 301000,
            'credito_para_tabla': 301000,
            'gastos_originacion': 40932.5,
            'enganche': 0,
            'inversion_cliente': 40932.5,
            'ingreso_minimo': 10000,
            'mensualidad': 3936.46,
            'cat': 0.157844027452,
        },
        "avisos": [],
    },
]


def resultados_credito(caso):
    """Corre la calculadora completa con los datos del caso."""
    r = credito.calcular_credito(**caso["datos"])
    fecha = date.fromisoformat(caso["fecha"])           # "2026-11-18" -> fecha de Python
    tabla = amortizacion.generar_tabla(fecha_calculo=fecha, **r["datos_tabla"])
    return {
        "credito_maximo": r["credito_maximo"],
        "credito_para_tabla": r["credito_para_tabla"],
        "gastos_originacion": r["gastos_originacion"],
        "enganche": r["enganche"],
        "inversion_cliente": r["inversion_cliente"],
        "ingreso_minimo": r["ingreso_minimo"],
        "mensualidad": tabla.loc[0, "Pago total"],
        "cat": indicadores.calcular_cat(r, fecha)["cat"],
    }, r["avisos"]


# =============================================================================
# FUNCIÓN QUE COMPARA (se usa en las dos partes)
# =============================================================================
def comparar(nuestros, del_excel):
    """Regresa la lista de diferencias mayores a la tolerancia (vacía si todo coincide)."""
    diferencias = []
    for concepto, valor_excel in del_excel.items():
        tolerancia = TOLERANCIA_CAT if concepto == "cat" else TOLERANCIA
        if abs(nuestros[concepto] - valor_excel) > tolerancia:
            diferencias.append(f"{concepto}: Excel {valor_excel:,.4f} vs Python {nuestros[concepto]:,.4f}")
    return diferencias


def imprimir(numero, nombre, diferencias):
    """Imprime [OK] o [FALLA] con el detalle."""
    if diferencias:
        print(f"[FALLA] {numero:2d}. {nombre}")
        for d in diferencias:
            print(f"           {d}")
    else:
        print(f"[OK]    {numero:2d}. {nombre}")


if __name__ == "__main__":
    # ----- Parte 1 -----
    print("PARTE 1. TABLA DE AMORTIZACIÓN")
    ok_tabla = 0
    for numero, escenario in enumerate(ESCENARIOS_TABLA, start=1):
        diferencias = comparar(resultados_tabla(escenario["datos"]), escenario["excel"])
        imprimir(numero, escenario["nombre"], diferencias)
        if not diferencias:
            ok_tabla += 1

    # ----- Parte 2 -----
    print("\nPARTE 2. CÁLCULO DEL CRÉDITO")
    ok_credito = 0
    for numero, caso in enumerate(ESCENARIOS_CREDITO, start=1):
        nuestros, avisos = resultados_credito(caso)
        diferencias = comparar(nuestros, caso["excel"])
        for aviso in caso["avisos"]:
            # any(...) es True si el texto del Excel aparece en alguno de nuestros avisos
            if not any(aviso in a for a in avisos):
                diferencias.append(f"falta el aviso: {aviso}")
        imprimir(numero, caso["nombre"], diferencias)
        if not diferencias:
            ok_credito += 1

    # ----- Resumen -----
    print("\n" + "=" * 60)
    print(f"Tabla de amortización: {ok_tabla} de {len(ESCENARIOS_TABLA)} coinciden con el Excel")
    print(f"Cálculo del crédito:   {ok_credito} de {len(ESCENARIOS_CREDITO)} coinciden con el Excel")
    print("=" * 60)
