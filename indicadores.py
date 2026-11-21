"""
indicadores.py
==============
INDICADORES DEL CRÉDITO: CAT, plazo real y totales

¿Qué hace este archivo?
    Con el crédito ya calculado (credito.py) y su tabla (amortizacion.py) obtiene
    los números de resumen que muestra el simulador:
        - CAT (Costo Anual Total), igual que la hoja "Tabla de amortizacion"
        - en cuánto tiempo se liquida el crédito ("17 años 8 meses")
        - totales: intereses, seguros, prepagos, penalizaciones, total pagado

Idea central: el CAT es una TASA INTERNA DE RETORNO (TIR)
    Es la tasa que hace que lo que el banco te da HOY sea igual al valor presente
    de todo lo que pagarás después. Es una ECUACIÓN DE VALOR con fecha focal en 0:

        * crédito neto = suma de  pago_k  x  v^k        con  v = 1 / (1 + i)
        *   (crédito neto = crédito - comisión - avalúo)
        *   (pago_k = pago del mes k + seguros)

    La i que cumple la igualdad es la TIR mensual, y el CAT la vuelve anual:

        * CAT = (1 + i)^12 - 1          (tasa efectiva anual equivalente)

    No hay una fórmula para despejar i, así que la buscamos por BISECCIÓN
    (el método numérico del curso): probamos una i, vemos si nos pasamos o nos
    falta, y partimos el intervalo a la mitad hasta encontrarla.

Etiquetas de comentarios (Better Comments):
    # !  HALLAZGO   # ?  NOTA   # *  FÓRMULA
"""

# -----------------------------------------------------------------------------
# IMPORTACIONES
# -----------------------------------------------------------------------------
import amortizacion             # (nuestro archivo) para armar la "corrida para CAT"


# -----------------------------------------------------------------------------
# CONSTANTES
# -----------------------------------------------------------------------------
ITERACIONES_BISECCION = 200     # cada vuelta parte el intervalo a la mitad;
                                # con 200 la precisión sobra (más que el Excel)


# =============================================================================
# 1. VALOR PRESENTE Y TIR
# =============================================================================

def valor_presente(flujos, i):
    """
    Suma de cada flujo traído al mes 0:  flujo_k x v^k,  con v = 1/(1+i).
    flujos[0] es el mes 0, flujos[1] el mes 1, etc.
    """
    v = 1 / (1 + i)
    # ? enumerate(flujos) da parejas (posición, valor): (0, f0), (1, f1), ...
    return sum(flujo * v ** k for k, flujo in enumerate(flujos))


def tir_mensual(flujos, minimo=0.0, maximo=1.0):
    """
    Busca la tasa mensual i que hace valor_presente(flujos, i) = 0, por bisección.

    ¿Por qué funciona?
        El primer flujo es negativo (el crédito que recibes, visto como salida
        del banco) y los demás positivos (tus pagos). Si i es muy chica, el valor
        presente de los pagos es grande -> resultado positivo. Si i es muy grande,
        los pagos valen poco hoy -> resultado negativo. En medio hay un cero.
    """
    for _ in range(ITERACIONES_BISECCION):      # "_" = no necesitamos el contador
        medio = (minimo + maximo) / 2
        if valor_presente(flujos, medio) > 0:
            minimo = medio      # todavía sobra valor presente: la tasa es más alta
        else:
            maximo = medio      # nos pasamos: la tasa es más baja
    return (minimo + maximo) / 2


# =============================================================================
# 2. CAT
# =============================================================================

def calcular_cat(resultado_credito, fecha_calculo):
    """
    Calcula el CAT igual que el Excel (hoja Tabla, columnas AT a BB, celda BB9).

    El Excel lo saca de una "Corrida para CAT": una segunda tabla del mismo
    crédito pero SIN prepagos y SIN aportaciones del Infonavit, es decir,
    pagando puntualmente lo pactado. Así el CAT no cambia si el usuario simula
    abonos extra.

    Devuelve un diccionario con: tir_mensual, tasa_anual_nominal y cat.
    """
    # ? dict(...) hace una COPIA del diccionario para no modificar el original.
    datos = dict(resultado_credito["datos_tabla"])
    datos["aportacion_bimestral"] = 0.0         # sin aportaciones Infonavit
    corrida_cat = amortizacion.generar_tabla(fecha_calculo=fecha_calculo, **datos)  # sin prepagos

    # Mes 0 (celda BB13):  -(crédito - comisión - avalúo)
    # * Lo que el cliente realmente recibe es el crédito menos lo que paga al firmar
    # * por comisión y avalúo; por eso el CAT es más alto que la tasa.
    credito_neto = (resultado_credito["credito_para_tabla"]
                    - resultado_credito["comision_contratacion"]
                    - resultado_credito["avaluo"])

    # Meses 1, 2, ... (columna BB): pago + seguros = columna "Pago total".
    # ? list(...) convierte la columna de la tabla en una lista normal de Python.
    # ? [a] + lista  pega el mes 0 al principio de la lista de pagos.
    flujos = [-credito_neto] + list(corrida_cat["Pago total"])

    i = tir_mensual(flujos)                                   # celda BB12
    return {
        "tir_mensual": i,
        "tasa_anual_nominal": i * 12,                         # celda BB8
        # * CAT = (1 + i)^12 - 1                              # celda BB9
        "cat": (1 + i) ** 12 - 1,
    }


# =============================================================================
# 3. PLAZO REAL Y TOTALES
# =============================================================================

def plazo_real(tabla):
    """
    ¿En cuánto tiempo se liquida?  Celdas Y10 (años) y Y11 (meses).
    El Excel lleva un contador por renglón (columnas AC y AD): en el mes 12
    marca 1 año 0 meses, en el mes 212 marca 17 años 8 meses.
    * años = meses // 12      (división entera)
    * meses = meses % 12      (residuo)
    """
    meses_totales = len(tabla)          # un renglón por mes pagado
    return {
        "meses_totales": meses_totales,
        "anios": meses_totales // 12,
        "meses": meses_totales % 12,
        "fecha_ultimo_pago": tabla.iloc[-1]["Fecha de pago"],   # .iloc[-1] = último renglón
    }


def totales(tabla):
    """Suma las columnas de la tabla (lo que se paga en toda la vida del crédito)."""
    return {
        "total_intereses": tabla["Interés"].sum(),
        "total_amortizacion": tabla["Amortización"].sum(),
        "total_prepagos": tabla["Prepago"].sum(),
        "total_penalizaciones": tabla["Penalización"].sum(),
        "total_aportaciones_infonavit": tabla["Aportación Infonavit"].sum(),
        "total_seguro_vida": tabla["Seguro de vida"].sum(),
        "total_seguro_danos": tabla["Seguro de daños"].sum(),
        "total_pagado": tabla["Pago total"].sum(),
    }


# =============================================================================
# 4. TODO JUNTO
# =============================================================================

def calcular_indicadores(resultado_credito, tabla, fecha_calculo):
    """
    Reúne en un solo diccionario todo lo que muestra el resumen.
    tabla = la tabla del usuario (con sus prepagos, si los puso).
    """
    indicadores = {
        # Mensualidad (celda I41): pago + seguros del mes 1.
        "mensualidad_mes_1": tabla.loc[0, "Pago total"],
    }
    # ? .update(otro) agrega al diccionario todas las llaves del otro diccionario.
    indicadores.update(calcular_cat(resultado_credito, fecha_calculo))
    indicadores.update(plazo_real(tabla))
    indicadores.update(totales(tabla))
    return indicadores


# =============================================================================
# PRUEBA RÁPIDA:  python indicadores.py
# Escenario 1: Tradicional, PO 20 años, casa de 2,000,000.
# =============================================================================
if __name__ == "__main__":
    from datetime import date
    import credito

    fecha = date(2026, 11, 18)
    r = credito.calcular_credito(producto=1, monto=2_000_000)
    tabla = amortizacion.generar_tabla(fecha_calculo=fecha, **r["datos_tabla"])
    ind = calcular_indicadores(r, tabla, fecha)

    print("Mensualidad del mes 1: {:,.2f}      (Excel: 24,294.00)".format(ind["mensualidad_mes_1"]))
    print("CAT:                   {:.2%}          (Excel: 15.13%)".format(ind["cat"]))
    print("Se liquida en:         {} años {} meses (Excel: 17 años 8 meses)".format(ind["anios"], ind["meses"]))
    print("Último pago:           {}".format(ind["fecha_ultimo_pago"]))
    print("Total de intereses:    {:,.2f}".format(ind["total_intereses"]))
    print("Total pagado:          {:,.2f}".format(ind["total_pagado"]))
