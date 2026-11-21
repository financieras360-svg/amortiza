"""
amortizacion.py
===============
MOTOR DE LA TABLA DE AMORTIZACIÓN

¿Qué hace este archivo?
    Construye la tabla de amortización del crédito: un renglón por mes, desde el
    mes 1 hasta que la deuda llega a cero. En cada mes calcula el interés, cuánto
    baja la deuda (amortización), los prepagos, la penalización, la aportación del
    Infonavit, los seguros y el saldo que queda.

¿Qué necesita?
    Los datos del crédito ya resueltos (monto, tasa, pago al millar, etc.). Esos
    datos los obtendrá credito.py a partir de lo que elija el usuario; este archivo
    solo se encarga de la tabla. Las constantes del banco las toma de parametros.py.

¿Qué entrega?
    Una tabla de pandas (DataFrame): una especie de hoja de cálculo dentro de
    Python, con columnas y renglones, que después se mostrará en la página.

Idea central (presentación 3 del curso, Amortización):
    Cada mes, del pago primero se cubre el interés y lo que sobra baja la deuda:
        interés       = saldo x tasa / 360 x días del mes         (F9)
        amortización  = pago - interés                            (F10)
        saldo final   = saldo inicial - amortización              (F11)
    Es el ENFOQUE RETROSPECTIVO del saldo, aplicado un mes a la vez.

Todas las fórmulas siguen exactamente al simulador del profesor (hoja
"Tabla de amortizacion"); al lado de cada cálculo se indica su columna.
"""

# -----------------------------------------------------------------------------
# IMPORTACIONES
# Traemos herramientas que ya existen para no programarlas desde cero.
# -----------------------------------------------------------------------------
import calendar                 # (de Python) para saber cuántos días tiene un mes
from datetime import date       # (de Python) para manejar fechas

import pandas as pd             # (librería) para armar la tabla final

import parametros               # (nuestro archivo) los datos del banco

# Al restar números con muchos decimales, la computadora puede dejar residuos
# diminutos (por ejemplo 0.0000000002) cuando el resultado debería ser cero.
# Excel los muestra como cero; nosotros tratamos como cero todo lo menor a esto.
CERO = 1e-6


# =============================================================================
# FUNCIÓN 1. FECHA DEL PRIMER PAGO
# Columnas BD, BJ y BK del Excel.
# -----------------------------------------------------------------------------
# Regla del banco:
#   - Si el crédito se calcula entre el día 3 y el 16 del mes, los pagos vencen
#     el día 3; en cualquier otro día, vencen el día 17.
#   - Si ese día de corte ya pasó en el mes del cálculo, el primer pago es el
#     mes siguiente; si no, es en el mismo mes.
# Ej.: cálculo el 18/nov/2026 -> corte día 17 -> el 17 de noviembre ya pasó
#      -> primer pago: 17/dic/2026.
# =============================================================================
def fecha_primer_pago(fecha_calculo):
    if parametros.DIA_CORTE_TEMPRANO <= fecha_calculo.day <= 16:
        dia_corte = parametros.DIA_CORTE_TEMPRANO
    else:
        dia_corte = parametros.DIA_CORTE_TARDIO

    # ¿Ya pasó el día de corte en este mes? Entonces el primer pago es el mes siguiente.
    if fecha_calculo.day > dia_corte:
        return sumar_un_mes(date(fecha_calculo.year, fecha_calculo.month, dia_corte))
    return date(fecha_calculo.year, fecha_calculo.month, dia_corte)


def sumar_un_mes(fecha):
    """Regresa la misma fecha un mes después (diciembre pasa a enero del año siguiente).
    El día siempre es 3 o 17, así que existe en todos los meses."""
    if fecha.month == 12:
        return date(fecha.year + 1, 1, fecha.day)
    return date(fecha.year, fecha.month + 1, fecha.day)


def dias_del_mes(fecha):
    """Días que tiene el mes de esa fecha (28, 29, 30 o 31). Columna BE del Excel.
    calendar.monthrange regresa dos números; el segundo [1] es el número de días."""
    return calendar.monthrange(fecha.year, fecha.month)[1]


# =============================================================================
# FUNCIÓN 2. TASA DEL MES (fórmula F5)
# Columnas AJ (Valora) y AL:AO (Pagos Oportunos) del Excel.
# -----------------------------------------------------------------------------
# La tasa empieza en tasa_inicial y baja por pagar puntual:
#   Pagos Oportunos: baja 0.20% en los meses 13, 25 y 37, y 0.15% en el 49.
#   Valora:          baja 0.25% cada 12 meses desde el mes 37 (37, 49, 61...),
#                    sin quedar por debajo del piso.
# Esta función recibe la tasa del mes anterior y regresa la de este mes.
# =============================================================================
def tasa_del_mes(mes, tasa_anterior, esquema, piso_valora):
    if esquema == "Pagos Oportunos":
        # .get(mes, 0) busca el mes en el diccionario; si no está, regresa 0 (no baja).
        return tasa_anterior - parametros.BAJADAS_PAGOS_OPORTUNOS.get(mes, 0)

    # Valora: ¿es un mes de bajada? (37, 49, 61...: mes >= 37 y separado de 37 por años completos)
    es_mes_de_bajada = (mes >= parametros.PRIMER_MES_BAJADA_VALORA and
                        (mes - parametros.PRIMER_MES_BAJADA_VALORA) % parametros.MESES_POR_ANIO == 0)
    if not es_mes_de_bajada:
        return tasa_anterior
    # max() se queda con el mayor de los dos: así la tasa nunca pasa del piso hacia abajo.
    return max(tasa_anterior - parametros.BAJADA_ANUAL_VALORA, piso_valora)


# =============================================================================
# FUNCIÓN 3. APORTACIÓN DEL INFONAVIT (productos 2 y 5, "con Apoyo")
# Columna N del Excel y celdas AF70:AF77 del Simulador.
# -----------------------------------------------------------------------------
# El patrón aporta 5% del sueldo; cada bimestre (2 meses) llega:
#   aportación = sueldo x 5% x 2      con tope = UMA x 25 x 5% x 60.8 = 8,915.56
# Ej.: sueldo de 30,000 -> 30,000 x 0.05 x 2 = 3,000 cada bimestre.
# =============================================================================
def aportacion_bimestral_inicial(sueldo_mensual):
    tope = (parametros.UMA_DIARIA * parametros.TOPE_VECES_UMA *
            parametros.APORTACION_PATRONAL_VIVIENDA * parametros.DIAS_BIMESTRE)
    aportacion = sueldo_mensual * parametros.APORTACION_PATRONAL_VIVIENDA * 2
    return min(aportacion, tope)          # min() se queda con el menor: nunca pasa del tope


def es_mes_de_aportacion(mes):
    """La aportación llega en los meses 2, 4, 6, 8... (meses pares)."""
    return mes >= parametros.MES_PRIMERA_APORTACION and mes % parametros.MESES_ENTRE_APORTACIONES == 0


def es_mes_de_aumento_aportacion(mes):
    """La aportación sube 3% en los meses 14, 26, 38... (cada 12 meses desde el 14)."""
    return (mes >= parametros.MES_PRIMER_INCREMENTO_APORTACION and
            (mes - parametros.MES_PRIMER_INCREMENTO_APORTACION) % parametros.MESES_POR_ANIO == 0)


def es_mes_de_aumento_pago(mes):
    """En Valora el pago sube en los meses 13, 25, 37... (inicio de cada año del crédito).
    Columna AA del Excel (llega hasta el mes 229)."""
    return 13 <= mes <= 229 and (mes - 13) % parametros.MESES_POR_ANIO == 0


# =============================================================================
# FUNCIÓN PRINCIPAL. GENERAR LA TABLA DE AMORTIZACIÓN
# -----------------------------------------------------------------------------
# Argumentos (los datos que recibe):
#   credito              Monto del crédito: saldo inicial del mes 1.
#   tasa_inicial         Tasa anual con la que empieza (ej. 0.13).
#   pago_al_millar       Pago mensual por cada 1,000 de crédito (ej. 11.86).
#   esquema              "Pagos Oportunos" o "Valora".
#   fecha_calculo        Fecha en que se hace el cálculo (define las fechas de pago).
#   valor_seguro_danos   Valor de la casa (o valor destructible) para el seguro de daños.
#   incremento_anual     g de Valora: cuánto sube el pago cada año (0 en Pagos Oportunos).
#   piso_valora          Tasa mínima de Valora (no se usa en Pagos Oportunos).
#   seguros_externos     True si el cliente contrató sus seguros por fuera (seguros = 0).
#   prepagos             Diccionario {mes: monto} con los abonos extra. Ej.: {26: 50_000}.
#   cobra_penalizacion   True si el tipo de comisión es "Pago anticipado".
#   tasa_penalizacion    % de penalización del producto (ej. 0.03).
#   reducir              Qué hacer tras un prepago: "plazo" (mismo pago, termina antes)
#                        o "pago" (paga menos cada mes). Celda AE3 del Excel.
#   aportacion_bimestral Aportación inicial del Infonavit (0 si el producto no tiene apoyo).
#   producto             Número de producto (los productos 2 y 7 tienen una regla
#                        especial al reducir el pago; ver más abajo).
#
# Los argumentos con "=" tienen un valor por defecto: si no se mandan, se usa ese.
# =============================================================================
def generar_tabla(credito, tasa_inicial, pago_al_millar, esquema, fecha_calculo,
                  valor_seguro_danos, incremento_anual=0.0, piso_valora=0.0,
                  seguros_externos=False, prepagos=None, cobra_penalizacion=False,
                  tasa_penalizacion=0.0, reducir="plazo", aportacion_bimestral=0.0,
                  producto=1):

    prepagos = prepagos or {}           # si no hay prepagos, usamos un diccionario vacío

    # Para "reducir el pago" el Excel usa una tabla de referencia SIN prepagos ni
    # aportaciones (su "Corrida para CAT"). De ella toma, cada mes, la proporción
    # pago / saldo. La generamos llamando a esta misma función (sin repetir código).
    # ? Ojo: el Excel aplica esa proporción SIEMPRE que se elige "reducir pago",
    # ? aunque no haya prepagos. Importa en el producto 5: sus aportaciones del
    # ? Infonavit bajan el saldo y, con "reducir pago", también bajan la mensualidad.
    referencia = None
    if reducir == "pago":
        referencia = generar_tabla(credito, tasa_inicial, pago_al_millar, esquema,
                                   fecha_calculo, valor_seguro_danos, incremento_anual,
                                   piso_valora, seguros_externos)

    filas = []                          # aquí se guardará un renglón (diccionario) por mes

    # Valores que van "pasando" de un mes al siguiente:
    saldo = credito                     # saldo inicial del mes 1 = crédito          (col. B)
    tasa = tasa_inicial                 # tasa anual vigente                         (col. P)
    pago = 0.0                          # pago de capital + interés del mes anterior (col. D)
    prepago_anterior = 0.0              # prepago del mes anterior (regla de productos 2 y 7)
    aportacion = aportacion_bimestral   # aportación Infonavit vigente               (col. N)
    fecha = fecha_primer_pago(fecha_calculo)

    # range(1, 241) produce los números 1, 2, ..., 240: hasta 20 años de meses.
    for mes in range(1, 241):

        # --- 1. Tasa y días del mes (F5) -------------------------------------
        tasa = tasa_del_mes(mes, tasa, esquema, piso_valora)
        dias = dias_del_mes(fecha)

        # --- 2. Interés del mes (F9): saldo x tasa / 360 x días ------  col. F
        interes = tasa / parametros.DIAS_ANIO_BANCARIO * dias * saldo

        # --- 3. Pago de capital + interés (F6, F7, F12, F14) ---------  col. D
        if mes == 1:
            pago = saldo / 1000 * pago_al_millar                     # F6: pago al millar
        else:
            aumento = incremento_anual if es_mes_de_aumento_pago(mes) else 0.0
            if saldo + interes - pago <= 0:
                pago = saldo + interes                               # F12: último pago
            elif producto in (2, 7) and prepago_anterior == 0:
                pago = pago * (1 + aumento)                          # regla especial 2 y 7
            elif reducir == "pago" and referencia is not None:
                proporcion = 0.0                                     # si la referencia ya terminó
                if mes <= len(referencia):
                    ref = referencia.iloc[mes - 1]                   # mismo mes en la referencia
                    proporcion = ref["Pago"] / ref["Saldo inicial"]
                pago = saldo * proporcion                            # F14: reducir el pago
            else:
                pago = pago * (1 + aumento)                          # F7: Valora sube cada año

        # --- 4. Amortización (F10): lo que sobra del pago --------------  col. H
        amortizacion = pago - interes

        # --- 5. Prepago y penalización (F13) ---------------------------  col. J, L
        prepago = prepagos.get(mes, 0.0)
        if cobra_penalizacion:
            penalizacion = prepago * tasa_penalizacion * (1 + parametros.IVA)
        else:
            penalizacion = 0.0

        # --- 6. Aportación del Infonavit (productos 2 y 5) -------------  col. N
        if aportacion_bimestral and es_mes_de_aportacion(mes):
            if es_mes_de_aumento_aportacion(mes):
                aportacion = aportacion * (1 + parametros.INCREMENTO_ANUAL_APORTACION)
            abono_infonavit = 0.0 if abs(saldo - amortizacion) < CERO else aportacion
        else:
            abono_infonavit = 0.0

        # --- 7. Saldo final (F11) ---------------------------------------  col. R
        # Como en el Excel, la penalización se SUMA a la deuda.
        saldo_final = saldo - amortizacion - prepago + penalizacion - abono_infonavit

        # --- 8. Seguros y pago total (F15) -----------------------------  col. T, V, Y
        if seguros_externos:
            seguro_vida = seguro_danos = 0.0       # asignación doble: ambos valen 0
        else:
            seguro_vida = saldo * parametros.SEGURO_VIDA
            seguro_danos = valor_seguro_danos * parametros.SEGURO_DANOS
        pago_total = pago + prepago + seguro_vida + seguro_danos

        # --- 9. Guardar el renglón del mes -------------------------------------
        filas.append({
            "Mes": mes,
            "Fecha de pago": fecha,
            "Días": dias,
            "Saldo inicial": saldo,
            "Tasa anual": tasa,
            "Interés": interes,
            "Pago": pago,
            "Amortización": amortizacion,
            "Prepago": prepago,
            "Penalización": penalizacion,
            "Aportación Infonavit": abono_infonavit,
            "Saldo final": saldo_final,
            "Seguro de vida": seguro_vida,
            "Seguro de daños": seguro_danos,
            "Pago total": pago_total,
        })

        # --- 10. Preparar el mes siguiente -------------------------------------
        if saldo_final <= CERO:       # la deuda se terminó: se detiene el ciclo
            break
        saldo = saldo_final
        prepago_anterior = prepago
        fecha = sumar_un_mes(fecha)

    # Convertimos la lista de renglones en una tabla de pandas.
    return pd.DataFrame(filas)


# -----------------------------------------------------------------------------
# PRUEBA RÁPIDA: solo se ejecuta con  python amortizacion.py
# Escenario del Excel: casa 2,000,000, crédito 1,900,000, Pagos Oportunos,
# 20 años, sin campaña, cálculo el 18/nov/2026.
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    tabla = generar_tabla(
        credito=1_900_000,
        tasa_inicial=0.13,
        pago_al_millar=11.86,
        esquema="Pagos Oportunos",
        fecha_calculo=date(2026, 11, 18),
        valor_seguro_danos=2_000_000,
    )
    pd.set_option("display.width", 200)     # para que la tabla quepa a lo ancho
    columnas = ["Mes", "Fecha de pago", "Días", "Saldo inicial", "Tasa anual",
                "Interés", "Amortización", "Saldo final", "Pago total"]
    print(tabla[columnas].head(3).round(2).to_string(index=False))
    print("...")
    print(tabla[columnas].tail(2).round(2).to_string(index=False))

    meses = len(tabla)                       # número de renglones = meses que dura
    print(f"\nMensualidad del mes 1: {tabla['Pago total'].iloc[0]:,.2f}   (Excel: 24,294.00)")
    print(f"Se liquida en: {meses // 12} años {meses % 12} meses   (Excel: 17 años 8 meses)")
