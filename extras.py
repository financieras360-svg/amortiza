"""
extras.py
=========
FUNCIONES ADICIONALES (capas de análisis sobre el núcleo validado)

¿Qué hace este archivo?
    Todo lo que la calculadora ofrece ADEMÁS de replicar el Excel:
        - mensualidad deseada  -> ¿cuánto me prestan si quiero pagar X al mes?
        - meta de liquidación   -> ¿cuánto abonar para terminar en N años?
        - pagar menos un tiempo -> ecuación de valor con opciones A, B y C
        - línea del tiempo      -> fechas clave del crédito
        - "¿por qué?"           -> explicación de cada número de un mes de la tabla
        - zona                  -> valor mínimo según la ciudad (solo informativo)
        - resumen anual de la tabla y datos de la carátula (RECA)

Regla de oro:
    Nada de aquí cambia la cotización. Todas las funciones USAN el motor
    (credito.py, amortizacion.py, indicadores.py) pero no lo modifican; por eso
    la validación contra el Excel sigue siendo válida.
"""

import pandas as pd

import amortizacion
import credito
import parametros


# =============================================================================
# 1. MENSUALIDAD DEL MES 1 Y "MENSUALIDAD DESEADA"
# =============================================================================

def mensualidad_mes_1(resultado_credito):
    """
    Pago total del mes 1 (pago + seguros) sin armar la tabla completa.
    * M1 = C / 1000 x pago al millar  +  0.0006 x C  +  0.00031 x V
    """
    d = resultado_credito["datos_tabla"]
    pago = d["credito"] / 1000 * d["pago_al_millar"]
    if d["seguros_externos"]:
        return pago
    return (pago + d["credito"] * parametros.SEGURO_VIDA
            + d["valor_seguro_danos"] * parametros.SEGURO_DANOS)


def credito_para_mensualidad(mensualidad_deseada, **datos):
    """
    Función nueva (no existe en el Excel): busca por BISECCIÓN el crédito cuya
    mensualidad del mes 1 es igual a la deseada, calculando en modo "monto".
    datos = los mismos argumentos de credito.calcular_credito (sin monto ni modo).
    """
    minimo, maximo = 0.0, 100_000_000.0
    for _ in range(100):
        medio = (minimo + maximo) / 2
        r = credito.calcular_credito(monto=medio, modo=credito.MODO_MONTO, **datos)
        if mensualidad_mes_1(r) < mensualidad_deseada:
            minimo = medio          # todavía alcanza para más crédito
        else:
            maximo = medio
    return round((minimo + maximo) / 2, 2)


# =============================================================================
# 2. ZONA (solo informativo, H9)
# =============================================================================
# Hoja "Pol", sección A.7. El simulador SIEMPRE usa la columna de CDMX, GDL y MTY.
MINIMO_POR_ZONA = {
    # producto: (CDMX, Guadalajara y Monterrey, resto del país)
    1: (400_000, None),          # la política no da cifra para el resto del país
    2: (400_000, None),
    3: (300_000, 275_000),
    4: (None, None),             # Pasivos: no aplica
    5: (None, None),
    6: (1_000_000, 750_000),
    7: (300_000, 275_000),
}


def minimo_por_zona(producto, zona):
    """zona = 'metro' (CDMX, GDL, MTY) o 'resto'. Devuelve el mínimo o None."""
    metro, resto = MINIMO_POR_ZONA[producto]
    if zona == "metro":
        return metro
    return resto if resto is not None else metro


# =============================================================================
# 3. META DE LIQUIDACIÓN ("Terminar antes")
# =============================================================================

def meta_de_liquidacion(resultado_credito, fecha_calculo, anios_meta,
                        modalidad="unico", mes_pago=24, prepagos_base=None):
    """
    ¿Cuánto hay que abonar para terminar en 'anios_meta' años?

    modalidad = "unico"   -> un solo pago extra en el mes 'mes_pago'
              = "mensual" -> la misma cantidad extra cada mes hasta la meta

    Se busca por BISECCIÓN el abono MÍNIMO con el que la tabla (el mismo motor
    del Excel) termina en el mes meta. Si el producto cobra penalización, el
    motor ya la suma al saldo, así que el monto que devuelve es el que paga el
    cliente (bruto).

    * Teoría (ecuación de valor, fecha focal 0):  P = R a_n + X v^k
    *   ->  X = R v^(n-k) a_(N-n): el valor, en el mes k, de los pagos que te ahorras.
    """
    datos = resultado_credito["datos_tabla"]
    prepagos_base = dict(prepagos_base or {})
    base = amortizacion.generar_tabla(fecha_calculo=fecha_calculo,
                                      prepagos=prepagos_base, **datos)
    meses_meta = int(anios_meta * 12)
    if meses_meta >= len(base):
        return None                                   # ya terminas antes de la meta

    def tabla_con(monto):
        p = dict(prepagos_base)
        meses = [mes_pago] if modalidad == "unico" else range(1, meses_meta + 1)
        for m in meses:
            p[m] = p.get(m, 0) + monto
        return amortizacion.generar_tabla(fecha_calculo=fecha_calculo, prepagos=p, **datos)

    minimo, maximo = 0.0, datos["credito"] * 1.2
    for _ in range(60):
        medio = (minimo + maximo) / 2
        if len(tabla_con(medio)) <= meses_meta:
            maximo = medio
        else:
            minimo = medio
    abono = maximo                                    # el menor abono que cumple la meta
    meta = tabla_con(abono)

    return {
        "abono": abono,
        "modalidad": modalidad,
        "mes_pago": mes_pago,
        "meses_base": len(base),
        "meses_meta": len(meta),
        "meses_ahorrados": len(base) - len(meta),
        "fecha_fin_base": base.iloc[-1]["Fecha de pago"],
        "fecha_fin_meta": meta.iloc[-1]["Fecha de pago"],
        "intereses_base": base["Interés"].sum(),
        "intereses_meta": meta["Interés"].sum(),
        "ahorro_intereses": base["Interés"].sum() - meta["Interés"].sum(),
        "penalizacion_total": meta["Penalización"].sum(),
        "nueva_mensualidad": (meta.loc[0, "Pago total"] if modalidad == "mensual"
                              else base.loc[0, "Pago total"]),
        "tabla_base": base,
        "tabla_meta": meta,
    }


# =============================================================================
# 4. PAGAR MENOS POR UN TIEMPO (ecuación de valor, fórmulas del curso)
# =============================================================================

def pagar_menos(D, desde, meses, reparto, tasa_anual, con_penalizacion=False,
                factor_penalizacion=0.03 * 1.16):
    """
    Pago D menos durante 'meses' meses, desde el mes 'desde' (a) hasta b = a+n-1.
    ¿Cómo lo compenso? Tres opciones que valen lo mismo (ecuación de valor):

    * A. Abonar antes, en el mes a-1:            X_A = D a_n
    * B. Pagar después, en el mes b:             X_B = D s_n   (= X_A (1+i)^n)
    * C. Repartir en m meses desde b+1:          E   = D s_n / a_m

    Se usa i = tasa anual / 12 (tasa nominal mensual del curso). El motor del
    banco usa días reales / 360 y bajadas de tasa, por eso puede variar unos pesos.
    Si los abonos pagan penalización, el abono bruto es X / (1 - 0.03 x 1.16).
    """
    i = tasa_anual / 12
    v = 1 / (1 + i)
    a_n = (1 - v ** meses) / i
    s_n = ((1 + i) ** meses - 1) / i
    a_m = (1 - v ** reparto) / i

    xa, xb, e = D * a_n, D * s_n, D * s_n / a_m
    if con_penalizacion:
        xa, xb, e = (x / (1 - factor_penalizacion) for x in (xa, xb, e))
    b = desde + meses - 1
    return {
        "i": i, "a_n": a_n, "s_n": s_n, "a_m": a_m,
        "A": xa, "mes_A": desde - 1,
        "B": xb, "mes_B": b,
        "C": e, "desde_C": b + 1, "hasta_C": b + reparto,
    }


# =============================================================================
# 5. LÍNEA DEL TIEMPO
# =============================================================================

def linea_del_tiempo(tabla, resultado_credito, fecha_calculo):
    """Lista de eventos clave (diccionarios) en orden de tiempo."""
    eventos = [{"mes": 0, "fecha": fecha_calculo, "tipo": "inicio",
                "titulo": "Cotizas tu crédito",
                "detalle": "Tu día de pago será el {} de cada mes.".format(
                    tabla.iloc[0]["Fecha de pago"].day)}]

    primero = tabla.iloc[0]
    eventos.append({"mes": 1, "fecha": primero["Fecha de pago"], "tipo": "pago",
                    "titulo": "Primer pago: ${:,.2f}".format(primero["Pago total"]),
                    "detalle": "De tu pago de ${:,.2f}, ${:,.2f} son interés y ${:,.2f} bajan tu deuda.".format(
                        primero["Pago"], primero["Interés"], primero["Amortización"])})

    # Cambios de tasa
    for k in range(1, len(tabla)):
        antes, ahora = tabla.iloc[k - 1]["Tasa anual"], tabla.iloc[k]["Tasa anual"]
        if ahora < antes - 1e-9:
            eventos.append({"mes": k + 1, "fecha": tabla.iloc[k]["Fecha de pago"], "tipo": "tasa",
                            "titulo": "Tu tasa baja a {:.2%}".format(ahora),
                            "detalle": "Por pagar puntualmente (antes {:.2%}).".format(antes)})

    # Valora: la mensualidad sube cada año
    g = resultado_credito["incremento_anual"]
    if g > 0 and len(tabla) >= 13:
        eventos.append({"mes": 13, "fecha": tabla.iloc[12]["Fecha de pago"], "tipo": "sube",
                        "titulo": "Tu mensualidad empieza a subir {:.1%} cada año".format(g),
                        "detalle": "Así funciona Valora: empiezas pagando menos."})

    # Primera aportación del Infonavit
    aport = tabla[tabla["Aportación Infonavit"] > 0]
    if len(aport):
        f = aport.iloc[0]
        eventos.append({"mes": int(f["Mes"]), "fecha": f["Fecha de pago"], "tipo": "infonavit",
                        "titulo": "Primera aportación del Infonavit: ${:,.2f}".format(f["Aportación Infonavit"]),
                        "detalle": "Cada dos meses tu patrón abona a tu deuda."})

    # Más capital que interés
    cruce = tabla[tabla["Amortización"] > tabla["Interés"]]
    if len(cruce):
        f = cruce.iloc[0]
        eventos.append({"mes": int(f["Mes"]), "fecha": f["Fecha de pago"], "tipo": "cruce",
                        "titulo": "Ya pagas más a tu deuda que de interés",
                        "detalle": "Desde aquí la amortización supera al interés."})

    # Mitad de la deuda
    mitad = tabla[tabla["Saldo final"] <= tabla.iloc[0]["Saldo inicial"] / 2]
    if len(mitad):
        f = mitad.iloc[0]
        eventos.append({"mes": int(f["Mes"]), "fecha": f["Fecha de pago"], "tipo": "mitad",
                        "titulo": "Debes menos de la mitad",
                        "detalle": "Tu saldo baja de ${:,.0f}.".format(tabla.iloc[0]["Saldo inicial"] / 2)})

    ultimo = tabla.iloc[-1]
    plazo_meses = resultado_credito["plazo"] * 12
    antes = plazo_meses - len(tabla)
    eventos.append({"mes": len(tabla), "fecha": ultimo["Fecha de pago"], "tipo": "fin",
                    "titulo": "¡Tu casa queda libre de deuda!",
                    "detalle": "Último pago de ${:,.2f}".format(ultimo["Pago total"])
                    + (" · {} meses antes del plazo de {} años.".format(antes, resultado_credito["plazo"]) if antes > 0 else ".")})

    eventos.sort(key=lambda e: e["mes"])
    return eventos


# =============================================================================
# 6. "¿POR QUÉ?" DE UN MES DE LA TABLA
# =============================================================================

def explicar_mes(tabla, resultado_credito, mes):
    """
    Explica con los números del usuario cada cálculo del mes elegido.
    Devuelve una lista de diccionarios: concepto, valor, cuenta (texto),
    latex (fórmula del curso) y columna del Excel.
    """
    f = tabla.iloc[mes - 1]
    d = resultado_credito["datos_tabla"]
    salida = []

    salida.append({
        "concepto": "Interés", "valor": f["Interés"], "columna": "F",
        "explicacion": "Lo que cobra el banco este mes por lo que todavía debes. "
                       "Usa año de 360 días y los días reales del mes.",
        "cuenta": "{:,.2f} × {:.4%} ÷ 360 × {} = {:,.2f}".format(
            f["Saldo inicial"], f["Tasa anual"], int(f["Días"]), f["Interés"]),
        "latex": r"I_k = S_{k-1}\cdot \frac{r_k}{360}\cdot d_k",
    })

    if mes == 1:
        cuenta_pago = "{:,.2f} ÷ 1000 × {} = {:,.2f}".format(d["credito"], d["pago_al_millar"], f["Pago"])
        expl = "El primer pago sale del pago al millar: cuánto pagas por cada $1,000."
        latex = r"R_1 = C\cdot\frac{p}{1000}"
    else:
        cuenta_pago = "{:,.2f}".format(f["Pago"])
        expl = ("En Pagos Oportunos el pago se queda igual; en Valora sube cada 12 meses; "
                "si abonaste y elegiste «reducir pago», se recalcula.")
        latex = r"R_k = R_{k-1}\,(1+g)\ \text{en los meses } 13, 25, \dots"
    salida.append({"concepto": "Pago (erogación)", "valor": f["Pago"], "columna": "D",
                   "explicacion": expl, "cuenta": cuenta_pago, "latex": latex})

    salida.append({
        "concepto": "Amortización", "valor": f["Amortización"], "columna": "H",
        "explicacion": "La parte de tu pago que sí baja tu deuda.",
        "cuenta": "{:,.2f} − {:,.2f} = {:,.2f}".format(f["Pago"], f["Interés"], f["Amortización"]),
        "latex": r"A_k = R_k - I_k",
    })

    salida.append({
        "concepto": "Saldo final", "valor": f["Saldo final"], "columna": "R",
        "explicacion": "Lo que debes después del pago (enfoque retrospectivo, mes a mes).",
        "cuenta": "{:,.2f} − {:,.2f} − {:,.2f} + {:,.2f} − {:,.2f} = {:,.2f}".format(
            f["Saldo inicial"], f["Amortización"], f["Prepago"], f["Penalización"],
            f["Aportación Infonavit"], f["Saldo final"]),
        "latex": r"S_k = S_{k-1} - A_k - P_k + \text{pen}_k - \text{ap}_k",
    })

    salida.append({
        "concepto": "Seguro de vida", "valor": f["Seguro de vida"], "columna": "T",
        "explicacion": "0.06 % de lo que debes al inicio del mes.",
        "cuenta": "{:,.2f} × 0.0006 = {:,.2f}".format(f["Saldo inicial"], f["Seguro de vida"]),
        "latex": r"SV_k = 0.0006\, S_{k-1}",
    })
    salida.append({
        "concepto": "Seguro de daños", "valor": f["Seguro de daños"], "columna": "V",
        "explicacion": "0.031 % del valor de la casa (o del valor destructible); es fijo.",
        "cuenta": "{:,.2f} × 0.00031 = {:,.2f}".format(d["valor_seguro_danos"], f["Seguro de daños"]),
        "latex": r"SD = 0.00031\, V",
    })
    salida.append({
        "concepto": "Pago total del mes", "valor": f["Pago total"], "columna": "Y",
        "explicacion": "Lo que sale de tu bolsillo: pago + abono extra + seguros.",
        "cuenta": "{:,.2f} + {:,.2f} + {:,.2f} + {:,.2f} = {:,.2f}".format(
            f["Pago"], f["Prepago"], f["Seguro de vida"], f["Seguro de daños"], f["Pago total"]),
        "latex": r"\text{Pago}_k = R_k + P_k + SV_k + SD",
    })
    if f["Penalización"] > 0:
        salida.append({
            "concepto": "Penalización", "valor": f["Penalización"], "columna": "L",
            "explicacion": "Por abonar antes con comisión «Pago anticipado»: 3 % + IVA del abono. "
                           "El Excel la suma al saldo.",
            "cuenta": "{:,.2f} × 0.03 × 1.16 = {:,.2f}".format(f["Prepago"], f["Penalización"]),
            "latex": r"\text{pen} = P\cdot 0.03\cdot 1.16",
        })
    return salida


# =============================================================================
# 7. RESUMEN ANUAL Y CARÁTULA
# =============================================================================

def resumen_anual(tabla):
    """Agrupa la tabla por año (12 meses) y suma las columnas de dinero."""
    t = tabla.copy()
    t["Año"] = (t["Mes"] - 1) // 12 + 1          # meses 1-12 -> año 1, etc.
    columnas = ["Interés", "Pago", "Amortización", "Prepago", "Penalización",
                "Aportación Infonavit", "Seguro de vida", "Seguro de daños", "Pago total"]
    resumen = t.groupby("Año")[columnas].sum()
    resumen["Saldo al final del año"] = t.groupby("Año")["Saldo final"].last()
    return resumen.reset_index()


# Registro de contratos (RECA) por clave producto-esquema-cónyuge
# (esquema 1 = Valora, 2 = Pagos Oportunos; cónyuge 2 = con datos del cónyuge).
RECAS = {
    "1-2-1": ("Adquisición Pagos Oportunos", "0319-138-038229/06-03318-1225"),
    "2-2-1": ("Adquisición Con Apoyo Pagos Oportunos", "0319-138-038230/04-03319-1225"),
    "1-1-1": ("Adquisición Valora", "0319-138-038232/05-03320-1225"),
    "2-1-1": ("Adquisición Con Apoyo Valora", "0319-138-038235/04-03321-1225"),
    "4-2-1": ("Pago de Pasivos Pagos Oportunos", "0319-138-007480/26-03330-1225"),
    "4-1-1": ("Pago de Pasivos Valora", "0319-138-007481/26-03331-1225"),
    "6-2-1": ("Liquidez Vivienda Pagos Oportunos", "0319-138-007482/27-03334-1225"),
    "6-1-1": ("Liquidez Vivienda Valora", "0319-138-007483/27-03335-1225"),
    "3-2-1": ("Cofinavit Pagos Oportunos", "0319-138-038237/04-03336-1225"),
    "3-2-2": ("Cofinavit Conyugales Pagos Oportunos", "0319-138-038238/04-03337-1225"),
    "3-1-1": ("Cofinavit Valora", "0319-138-038239/04-03338-1225"),
    "3-1-2": ("Cofinavit Conyugales Valora", "0319-138-038240/04-03339-1225"),
    "5-2-1": ("Pago de Pasivos Pagos Oportunos con Apoyo Infonavit", "0319-138-024405/17-03332-1225"),
    "5-1-1": ("Pago de Pasivos Valora con Apoyo Infonavit", "0319-138-024404/17-03333-1225"),
    "7-1-1": ("Cofinavit CI+CB Valora", "0319-138-035429/08-03354-1225"),
    "7-1-2": ("Cofinavit CI+CB Conyugal Valora", "0319-138-035430/08-03355-1225"),
    "7-2-1": ("Cofinavit CI+CB Pagos Oportunos", "0319-138-035431/08-03356-1225"),
    "7-2-2": ("Cofinavit CI+CB Conyugal Pagos Oportunos", "0319-138-035432/08-03357-1225"),
}

PROGRAMA = {1: "CREDIRESIDENCIAL", 2: "CREDIRESIDENCIAL", 3: "CREDIRESIDENCIAL",
            4: "CREDIRESIDENCIAL", 5: "CREDIRESIDENCIAL", 6: "CREDILIQUIDEZ",
            7: "CREDIRESIDENCIAL"}


def reca(producto, esquema, con_conyuge):
    clave = "{}-{}-{}".format(producto, 1 if esquema == "Valora" else 2, 2 if con_conyuge else 1)
    return clave, RECAS.get(clave)


# =============================================================================
# 8. "¿Y SI LA TASA FUERA...?" (simulador con la fórmula del curso)
# =============================================================================
# ? Esto NO es el cálculo del Excel: el banco usa su pago al millar, días reales
# ? / 360 y bajadas de tasa. Aquí el usuario escribe CUALQUIER tasa y se usa la
# ? anualidad vencida del curso, con i = tasa anual / 12. La cotización no cambia.
#
# * R = C · i / (1 - (1+i)^-n)      (pago fijo de una anualidad vencida)
# * I_k = S_(k-1) · i     A_k = R - I_k     S_k = S_(k-1) - A_k

def simular_tasa(credito_monto, tasa_anual, plazo_meses):
    """Tabla con pago fijo para una tasa elegida por el usuario."""
    i = tasa_anual / 12
    n = int(plazo_meses)
    if i == 0:
        pago = credito_monto / n          # sin interés: solo se reparte el crédito
    else:
        pago = credito_monto * i / (1 - (1 + i) ** -n)

    filas = []
    saldo = credito_monto
    for mes in range(1, n + 1):
        interes = saldo * i
        amortizacion_mes = pago - interes
        saldo_final = saldo - amortizacion_mes
        if mes == n:                      # el último mes elimina residuos de redondeo
            saldo_final = 0.0
        filas.append({"Mes": mes, "Saldo inicial": saldo, "Interés": interes,
                      "Pago": pago, "Amortización": amortizacion_mes, "Saldo final": saldo_final})
        saldo = saldo_final

    tabla = pd.DataFrame(filas)
    return {
        "i": i,
        "n": n,
        "pago": pago,
        "pago_al_millar": pago / credito_monto * 1000,
        "intereses": tabla["Interés"].sum(),
        "total_pagado": pago * n,
        "tabla": tabla,
    }


# =============================================================================
# PRUEBA RÁPIDA:  python extras.py
# =============================================================================
if __name__ == "__main__":
    from datetime import date
    fecha = date(2026, 11, 20)
    r = credito.calcular_credito(producto=1, monto=2_000_000)           # contratación
    u = meta_de_liquidacion(r, fecha, 15, "unico", 24)
    m = meta_de_liquidacion(r, fecha, 15, "mensual")
    print("Meta 15 años, pago único mes 24: {:,.2f}  (esperado 119,023.03)".format(u["abono"]))
    print("Meta 15 años, extra mensual:     {:,.2f}  (esperado 1,143.04)".format(m["abono"]))
    print("Ahorro de intereses (único):     {:,.2f}  (esperado 582,119.91)".format(u["ahorro_intereses"]))
    r2 = credito.calcular_credito(producto=1, monto=2_000_000, tipo_comision="Pago anticipado")
    print("Con pago anticipado:             {:,.2f}  (esperado 123,314.37)".format(
        meta_de_liquidacion(r2, fecha, 15, "unico", 24)["abono"]))
    pm = pagar_menos(3000, 13, 24, 24, 0.13)
    print("Pagar menos A/B/C: {:,.2f} / {:,.2f} / {:,.2f}  (esperado 63,102.33 / 81,724.97 / 3,885.35)".format(
        pm["A"], pm["B"], pm["C"]))
    c = credito_para_mensualidad(24_294, producto=1)
    print("Crédito para pagar 24,294 al mes: {:,.2f}  (esperado 1,900,000.00)".format(c))
