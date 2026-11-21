"""
credito.py
==========
CÁLCULO DEL CRÉDITO (todo lo que pasa ANTES de la tabla de amortización)

¿Qué hace este archivo?
    Recibe lo que el usuario captura en la calculadora ("mi casa vale 2,000,000,
    quiero Pagos Oportunos a 20 años, soy empleado...") y lo convierte en los
    números del crédito:
        - condiciones: tasa, pago al millar, incremento anual (Valora)
        - aforo y crédito máximo
        - valor de la vivienda
        - crédito Infonavit y subcuenta (productos 3 y 7)
        - gastos financiados (productos 4 y 5)
        - gastos de originación, enganche e inversión del cliente
        - ingreso mínimo
        - avisos (los mensajes en rojo del Excel)
    Al final arma los datos que necesita amortizacion.generar_tabla().

¿De dónde sale cada cálculo?
    De la hoja "Simulador" del Excel del profesor. Junto a cada paso se indica la
    celda. Todo se comprobó contra el Excel con los casos de validacion.py.

Etiquetas de comentarios (extensión Better Comments):
    # !  HALLAZGO  -> algo raro del Excel que replicamos TAL CUAL (indicación del profe)
    # ?  NOTA      -> explicación o decisión nuestra
    # *  FÓRMULA   -> la fórmula del cálculo
"""

# -----------------------------------------------------------------------------
# IMPORTACIONES
# -----------------------------------------------------------------------------
import parametros               # (nuestro archivo) los datos del banco
import amortizacion             # (nuestro archivo) para la aportación del Infonavit


# -----------------------------------------------------------------------------
# CONSTANTES DE ESTE ARCHIVO
# -----------------------------------------------------------------------------
# Los tres modos de cálculo del Excel (celda Y41: 1, 2 y 3).
MODO_VALOR = "valor"            # el usuario escribe el valor de la casa
MODO_INGRESOS = "ingresos"      # el usuario escribe su ingreso mensual
MODO_MONTO = "monto"            # el usuario escribe el crédito que quiere

# Monto máximo "de política" que revisa el Excel en modo monto (celda AD44).
MONTO_MAXIMO_POLITICA = 15_000_000
MONTO_MAXIMO_POLITICA_NIVEL_R = 7_000_000

# En modo ingresos, con Valora nivel 4 o R, el crédito máximo es 25 veces el
# ingreso (celda AC86, regla "LTI": loan to income, crédito contra ingreso).
VECES_INGRESO_LTI = 25

# Un número "casi cero" para comparar decimales sin errores de redondeo.
CERO = 1e-6


# =============================================================================
# FUNCIONES PEQUEÑAS (cada una resuelve una sola pregunta)
# =============================================================================

def condiciones_de_tasa(esquema, campana, nivel, plazo):
    """
    ¿Con qué tasa, pago al millar e incremento anual se calcula el crédito?
    Celdas T83 (tasa), T84 (pago al millar) y T85 (incremento anual).

    Devuelve un diccionario con: tasa, pago_al_millar, incremento_anual, piso.
    """
    if esquema == "Pagos Oportunos":
        # En PO la tasa no depende del nivel, solo de la campaña.
        datos = parametros.PAGOS_OPORTUNOS[campana]
        return {
            "tasa": datos["tasa"],
            "pago_al_millar": datos["pago_al_millar"][plazo],
            "incremento_anual": 0.0,     # en PO el pago no crece (T85 = 0)
            "piso": 0.0,                 # en PO no hay piso (baja por meses fijos)
        }

    # Esquema Valora
    datos_campana = parametros.VALORA[campana]
    niveles = datos_campana["niveles"]

    # ! HALLAZGO H4: en el Excel, Valora con campaña y nivel mayor a 1 deja la
    # ! tasa en blanco (la tabla de la campaña solo llena el renglón 1).
    # ! Decisión: a todos los niveles les aplicamos el único renglón de la campaña.
    # ? niveles.get(nivel, niveles.get("todos")) significa: "busca el nivel; si no
    # ? existe esa llave, usa la llave 'todos'". .get() evita un error si no está.
    datos_nivel = niveles.get(nivel, niveles.get("todos"))

    return {
        "tasa": datos_nivel["tasa"],
        "pago_al_millar": datos_nivel["pago_al_millar"][plazo],
        "incremento_anual": datos_nivel["incremento_anual"],
        "piso": datos_campana["piso"],
    }


def aforo_del_producto(producto, actividad, nivel):
    """
    ¿Qué porcentaje del valor de la casa presta el banco?  Celda X39.
    (Columnas AH, AI y AJ de las filas 48 a 52.)
    """
    # Cofinanciamientos (3 y 7): aforo fijo, celdas X34 y X38.
    if producto in parametros.AFORO_FIJO:
        return parametros.AFORO_FIJO[producto]

    # Productos con Apoyo Infonavit (2 y 5): columna "Emp. con apoyo" (AH).
    if parametros.PRODUCTOS[producto]["tipo_infonavit"] == "apoyo":
        return parametros.AFORO_EMPLEADO_CON_APOYO[nivel]

    # Producto 4 con independiente: columna AJ con su excepción para Pasivo.
    if producto == 4 and actividad == "Independiente":
        return parametros.AFORO_INDEPENDIENTE_PASIVO[nivel]

    # ! HALLAZGO H1: las fórmulas de aforo (AH48:AJ52) tienen una segunda parte
    # ! "por montos altos" que revisa la celda I21, que está vacía; por eso
    # ! siempre usan la primera parte (la de créditos hasta 10 millones).
    # Productos 1, 4 y 6: depende de la actividad y el nivel.
    return parametros.AFORO[actividad][nivel]


def costo_avaluo(valor_vivienda):
    """
    ¿Cuánto cuesta el avalúo?  Celda O30 (BUSCARV aproximado en S91:U112).
    Recorre la tabla y se queda con el último renglón cuyo "desde" no pase del
    valor de la casa.
    """
    costo = 0
    # ? "for desde, precio in ..." desarma cada tupla (desde, precio) en dos variables.
    for desde, precio in parametros.AVALUOS:
        if valor_vivienda >= desde:
            costo = precio
        else:
            break                # la lista está en orden: ya no hay que seguir
    # ! HALLAZGO H2: el último renglón del Excel dice "30,000,001 a 30,000,000";
    # ! es solo texto, el BUSCARV nunca lee esa columna, así que no afecta.
    return costo


def credito_cofinavit(valor_vivienda, ssv_total, credito_infonavit_total):
    """
    Producto 3 (Cofinavit): ¿cuánto presta el banco?  Celdas AE4 a AE20.

    * FÓRMULA
      banco = 95 % del valor  -  ajuste
      ajuste = lo que se pase de 100 % al sumar 95 % + (Infonavit + SSV) / valor
    """
    if valor_vivienda <= 0:
        return 0.0
    porcentaje_banco = parametros.AFORO_FIJO[3]                          # AE17
    porcentaje_infonavit = (credito_infonavit_total + ssv_total) / valor_vivienda  # AE16
    exceso = max(0.0, porcentaje_banco + porcentaje_infonavit - 1)      # AE19
    return porcentaje_banco * valor_vivienda - exceso * valor_vivienda   # AE20


def credito_ci_cb(valor_vivienda, ssv_total):
    """
    Producto 7 (CI+CB Cofinavit): reparte el crédito entre Infonavit y banco.
    Celdas AE24 a AE28.  Devuelve (total otorgado, crédito Infonavit, crédito banco).
    """
    monto_autorizado = valor_vivienda * parametros.AFORO_FIJO[7]         # AE24 (90 %)
    # AE26: si la subcuenta "cabe", se presta el 90 %; si no, valor - subcuenta.
    if valor_vivienda > monto_autorizado + ssv_total:
        total = monto_autorizado
    else:
        total = valor_vivienda - ssv_total
    # AE27: el Infonavit pone el 30 %, sin pasar de su monto máximo.
    infonavit = min(parametros.MONTO_MAXIMO_CREDITO_INFONAVIT,
                    total * parametros.PORCENTAJE_INFONAVIT_CI_CB)
    banco = total - infonavit                                            # AE28
    return total, infonavit, banco


# =============================================================================
# FUNCIÓN PRINCIPAL
# =============================================================================

def calcular_credito(
    producto,
    monto,                          # D21: valor de la casa, ingreso o crédito (según modo)
    modo=MODO_VALOR,                # Y41
    campana="Ninguna",              # W62
    actividad="Empleado",           # X29
    esquema="Pagos Oportunos",      # T68
    plazo=20,                       # T62 (en años)
    nivel="1",                      # T72 (solo cuenta en Valora)
    tipo_comision="Contratación",   # T41: "Contratación" o "Pago anticipado"
    comision_manual=None,           # D33 (porcentaje, ej. 0.01); None = la del producto
    seguros_externos=False,         # U29
    valor_destructible=0,           # D22 (para el seguro de daños)
    sueldo_cotizacion=0,            # D24 (productos 2 y 5)
    credito_solicitado=None,        # I25; None = el máximo (como el Excel)
    financiar_gastos=True,          # I27 (productos 4 y 5)
    ssv_titular=0,                  # D35
    ssv_conyuge=0,                  # D42 (producto 3) o D38 (producto 7)
    credito_infonavit_titular=0,    # D37 (producto 3)
    credito_infonavit_conyuge=0,    # D44 (producto 3)
    porcentaje_impuestos=None,      # M27; None = 3 %
    porcentaje_honorarios=None,     # M28; None = 2 %
    gastos_exactos_infonavit=0,     # L47 (productos 3 y 7)
):
    """
    Calcula todo el crédito igual que la hoja "Simulador".
    Devuelve un diccionario con los resultados (ver el final de la función).
    """
    datos = parametros.PRODUCTOS[producto]
    avisos = []        # lista de mensajes; se van agregando con .append()

    # -------------------------------------------------------------------------
    # PASO 1. Ajustes de captura (lo que el Excel no deja elegir)
    # -------------------------------------------------------------------------
    # ! HALLAZGO H5: en Pagos Oportunos la lista de "Tasa de interés" (T72)
    # ! solo trae una opción, así que el nivel siempre es 1.
    if esquema == "Pagos Oportunos":
        nivel = "1"

    # El Excel solo ofrece "Empleado" en los productos 2, 3 y 5 (celda W30).
    if producto in parametros.PRODUCTOS_SOLO_EMPLEADO:
        actividad = "Empleado"

    # ! HALLAZGO H6: el producto 7 (CI+CB) solo ofrece la actividad "Independiente".
    if producto == 7:
        actividad = "Independiente"

    # "Pago anticipado" no existe en los productos 2, 3 y 5 (celda S43).
    if tipo_comision == "Pago anticipado" and producto in (2, 3, 5):
        tipo_comision = "Contratación"
        avisos.append("Este producto no ofrece comisión por pago anticipado; "
                      "se usa comisión por contratación.")

    # Hipoteca Green solo existe en el producto 1 (celda V67).
    if campana in parametros.CAMPANAS_SOLO_PRODUCTO_1 and producto != 1:
        campana = "Ninguna"
        avisos.append("La campaña Hipoteca Green solo aplica a Adquisición Tradicional.")

    # -------------------------------------------------------------------------
    # PASO 2. Condiciones del crédito (T83, T84, T85) y aforo (X39)
    # -------------------------------------------------------------------------
    condiciones = condiciones_de_tasa(esquema, campana, nivel, plazo)
    pago_al_millar = condiciones["pago_al_millar"]
    aforo = aforo_del_producto(producto, actividad, nivel)

    # -------------------------------------------------------------------------
    # PASO 3. Crédito y valor de la vivienda según el modo (filas 82 a 86)
    # -------------------------------------------------------------------------
    if modo == MODO_VALOR:
        # * crédito = valor x aforo                              (W82)
        credito_base = monto * aforo
        valor_vivienda = monto                                 # AA82
    elif modo == MODO_MONTO:
        # * valor = crédito / aforo                              (AA82)
        credito_base = monto
        valor_vivienda = monto / aforo
    else:  # MODO_INGRESOS
        # * pago máximo = ingreso x PTI                          (Y83)
        # * crédito = pago máximo / (pago al millar / 1000)      (W83)
        pago_maximo = monto * datos["pti"]
        credito_base = pago_maximo / (pago_al_millar / 1000)
        valor_vivienda = credito_base / aforo                  # AA83

    # -------------------------------------------------------------------------
    # PASO 4. Crédito máximo posible (I24 = AD39, o la regla LTI de AC86)
    # -------------------------------------------------------------------------
    ssv_total = ssv_titular + ssv_conyuge
    credito_infonavit = 0.0         # I30
    porcentaje_ci_cb = None

    if producto == 3:
        credito_maximo = credito_cofinavit(
            valor_vivienda, ssv_total,
            credito_infonavit_titular + credito_infonavit_conyuge)        # AD34 = AE20
        credito_infonavit = credito_infonavit_titular + credito_infonavit_conyuge
    elif producto == 7:
        total_ci_cb, credito_infonavit, credito_maximo = credito_ci_cb(
            valor_vivienda, ssv_total)                                   # AD38 = AE28
        porcentaje_ci_cb = total_ci_cb / valor_vivienda if valor_vivienda else 0  # AF26
    else:
        credito_maximo = credito_base                                    # W86

    # Regla LTI (AC86): modo ingresos + Valora + nivel 4 o R  ->  25 x ingreso.
    if modo == MODO_INGRESOS and esquema == "Valora" and nivel in ("4", "R"):
        credito_maximo = monto * VECES_INGRESO_LTI

    # -------------------------------------------------------------------------
    # PASO 5. Crédito solicitado (I25) y gastos financiados (I28, O25)
    # -------------------------------------------------------------------------
    if credito_solicitado is None:
        credito_solicitado = credito_maximo          # en el Excel, I25 = I24

    gastos_financiados = 0.0
    if datos["financia_gastos"] and financiar_gastos:
        tasa_gf = parametros.PORCENTAJE_GASTOS_FINANCIADOS
        # ? round(..., 2) compara a centavos; así evitamos que 1,899,999.9999999
        # ? y 1,900,000 se vean distintos por los decimales de la computadora.
        if modo == MODO_VALOR and round(valor_vivienda * aforo, 2) == round(credito_solicitado, 2):
            gastos_financiados = 0.0                 # pidió el máximo: no hay espacio
        elif modo == MODO_VALOR and credito_solicitado * (1 + tasa_gf) > credito_maximo:
            gastos_financiados = credito_maximo - credito_solicitado   # lo que alcance
        else:
            gastos_financiados = credito_solicitado * tasa_gf          # 6 %
        # ! HALLAZGO H11: el tope (no pasar del crédito máximo) solo se revisa en
        # ! modo valor. En modo ingresos o monto siempre se suma el 6 %, aunque el
        # ! total quede arriba del máximo (comprobado en el Excel).

    # Crédito con el que se arma la tabla (B15 de la hoja Tabla).
    credito_para_tabla = credito_solicitado + gastos_financiados          # O25

    # -------------------------------------------------------------------------
    # PASO 6. Gastos de originación (O26 a O33)
    # -------------------------------------------------------------------------
    es_cofinanciamiento = datos["tipo_infonavit"] == "cofinanciamiento"

    # O26: comisión del Infonavit (3 % de su crédito), solo productos 3 y 7.
    comision_infonavit = credito_infonavit * parametros.COMISION_INFONAVIT if es_cofinanciamiento else 0.0

    # O27: impuestos y derechos.
    if es_cofinanciamiento and gastos_exactos_infonavit > 0:
        impuestos = gastos_exactos_infonavit - comision_infonavit
    else:
        pct = porcentaje_impuestos if porcentaje_impuestos is not None else parametros.IMPUESTOS_Y_DERECHOS
        impuestos = valor_vivienda * pct

    # O28: honorarios notariales.
    pct = porcentaje_honorarios if porcentaje_honorarios is not None else parametros.HONORARIOS_NOTARIALES
    honorarios = valor_vivienda * pct

    # O30: avalúo. Gratis con campaña WM o Premium y en los productos 4 y 5.
    if campana in parametros.CAMPANAS_SIN_AVALUO or not datos["cobra_avaluo"]:
        avaluo = 0
    else:
        avaluo = costo_avaluo(valor_vivienda)
    if valor_vivienda > parametros.AVALUO_LIMITE_COTIZAR:
        avisos.append("Avalúo: 'Favor de cotizar'. Valor aproximado.")     # M30

    # O31: comisión por contratación (sobre el crédito solicitado, I25).
    tasa_comision = comision_manual if comision_manual is not None else datos["comision_apertura"]  # V39
    if campana in parametros.CAMPANAS_SIN_COMISION or tipo_comision == "Pago anticipado":
        comision = 0.0
    else:
        comision = credito_solicitado * tasa_comision

    # O33: total de gastos, menos los que se financian.
    gastos_originacion = (comision_infonavit + impuestos + honorarios + avaluo + comision
                          - gastos_financiados)

    # -------------------------------------------------------------------------
    # PASO 7. Recursos, enganche e inversión del cliente (I33, L35, L37, L38)
    # -------------------------------------------------------------------------
    subcuenta = ssv_total if es_cofinanciamiento else 0.0                  # I29
    recursos_totales = credito_solicitado + subcuenta + (credito_infonavit if es_cofinanciamiento else 0.0)  # I33

    if producto in (4, 5):
        enganche = 0.0                                                      # L35
        etiqueta_enganche = "Desembolso"
    else:
        enganche = valor_vivienda - recursos_totales
        etiqueta_enganche = "Monto Remanente" if recursos_totales - valor_vivienda > 0 else "Enganche para la compra"
    if producto == 6:
        etiqueta_enganche = ""          # el Excel deja la etiqueta vacía en Liquidez (I35)

    if producto == 6:
        inversion_cliente = gastos_originacion                              # L38
    else:
        inversion_cliente = enganche + gastos_originacion

    # % de financiamiento (M20)
    if valor_vivienda == 0:
        porcentaje_financiamiento = 0.0
    elif producto == 7:
        porcentaje_financiamiento = porcentaje_ci_cb
    else:
        porcentaje_financiamiento = credito_para_tabla / valor_vivienda

    # -------------------------------------------------------------------------
    # PASO 8. Ingreso mínimo (I45, filas 59 a 73)
    # * ingreso mínimo = máx( ingreso mínimo del producto,
    # *                       crédito solicitado / 1000 x pago al millar / PTI )
    # -------------------------------------------------------------------------
    ingreso_minimo = max(datos["ingreso_minimo"],
                         credito_solicitado / 1000 * pago_al_millar / datos["pti"])

    # -------------------------------------------------------------------------
    # PASO 9. Avisos (los mensajes del Excel)
    # -------------------------------------------------------------------------
    # C23
    if monto > 0:
        if modo == MODO_VALOR and monto < datos["valor_minimo_vivienda"]:
            avisos.append("Valor de vivienda debajo del mínimo")
            # ! HALLAZGO H9: la política pide menos en el resto del país, pero el
            # ! simulador siempre usa el mínimo de CDMX, Guadalajara y Monterrey.
        elif monto < datos["ingreso_minimo"]:
            avisos.append("Ingresos menores al mínimo")

    # G26 / I41
    fuera_de_rango = credito_solicitado > credito_maximo + CERO
    if credito_solicitado > 0:
        if fuera_de_rango:
            avisos.append("Crédito fuera de rango")
        elif credito_solicitado < datos["credito_minimo"]:
            avisos.append("Crédito debajo del mínimo (mínimo {:,.2f})".format(datos["credito_minimo"]))

    # G31
    # ! HALLAZGO H7: el monto máximo del producto (15 M; 2.5 M en Liquidez) no se
    # ! revisa en ninguna fórmula. Solo en modo monto se compara contra 15 M
    # ! (7 M con nivel R) y el mensaje dice "debe ser igual al máximo".
    # ! HALLAZGO H10: la hoja de políticas dice que el máximo es 8 M.
    tope_politica = MONTO_MAXIMO_POLITICA_NIVEL_R if nivel == "R" else MONTO_MAXIMO_POLITICA
    if modo == MODO_MONTO and credito_maximo > tope_politica:
        avisos.append("El crédito solicitado debe ser igual al máximo")
    elif producto in (1, 4, 6) and (esquema == "Pagos Oportunos" or nivel in ("4", "R")):
        avisos.append("Línea máxima a otorgar puede ser menor por perfil del cliente")

    if enganche < 0:
        avisos.append("Negativo significa remanente para el cliente")      # M36

    # -------------------------------------------------------------------------
    # PASO 10. Datos que necesita la tabla de amortización
    # -------------------------------------------------------------------------
    # ? Seguro de daños: sobre el valor destructible si se capturó; si no, sobre
    # ? el valor de la vivienda (columna V de la hoja Tabla).
    valor_seguro_danos = valor_destructible if valor_destructible > 0 else valor_vivienda

    # Aportación del Infonavit (columna N): solo productos con apoyo (2 y 5).
    if datos["tipo_infonavit"] == "apoyo":
        aportacion = amortizacion.aportacion_bimestral_inicial(sueldo_cotizacion)
    else:
        aportacion = 0.0

    datos_tabla = {
        "credito": credito_para_tabla,
        "tasa_inicial": condiciones["tasa"],
        "pago_al_millar": pago_al_millar,
        "esquema": esquema,
        "valor_seguro_danos": valor_seguro_danos,
        "incremento_anual": condiciones["incremento_anual"],
        "piso_valora": condiciones["piso"],
        "seguros_externos": seguros_externos,
        # La penalización solo se cobra con comisión "Pago anticipado" (columna L).
        "cobra_penalizacion": tipo_comision == "Pago anticipado",
        "tasa_penalizacion": datos["penalizacion_prepago"],
        "aportacion_bimestral": aportacion,
        "producto": producto,
    }

    # -------------------------------------------------------------------------
    # RESULTADO: un diccionario con todo lo que mostrará la calculadora.
    # -------------------------------------------------------------------------
    return {
        # Detalle
        "producto": producto,
        "nombre_producto": datos["nombre"],
        "esquema": esquema,
        "plazo": plazo,
        "nivel": nivel,
        "campana": campana,
        "actividad": actividad,
        "porcentaje_financiamiento": porcentaje_financiamiento,
        # Recursos
        "aforo": aforo,
        "valor_vivienda": valor_vivienda,
        "credito_maximo": credito_maximo,
        "credito_solicitado": credito_solicitado,
        "gastos_financiados": gastos_financiados,
        "credito_para_tabla": credito_para_tabla,
        "subcuenta_vivienda": subcuenta,
        "credito_infonavit": credito_infonavit,
        "recursos_totales": recursos_totales,
        "enganche": enganche,
        "etiqueta_enganche": etiqueta_enganche,
        "inversion_cliente": inversion_cliente,
        # Gastos de originación (desglose)
        "comision_infonavit": comision_infonavit,
        "impuestos_y_derechos": impuestos,
        "honorarios_notariales": honorarios,
        "avaluo": avaluo,
        "comision_contratacion": comision,
        "gastos_originacion": gastos_originacion,
        # Condiciones
        "tasa": condiciones["tasa"],
        "pago_al_millar": pago_al_millar,
        "incremento_anual": condiciones["incremento_anual"],
        "ingreso_minimo": ingreso_minimo,
        # Parámetros del producto
        "credito_minimo": datos["credito_minimo"],
        "valor_minimo_vivienda": datos["valor_minimo_vivienda"],
        "enganche_minimo": 0 if aforo == 0 else 1 - aforo,                 # O44
        # Avisos
        "fuera_de_rango": fuera_de_rango,
        "avisos": avisos,
        # Para la tabla de amortización
        "datos_tabla": datos_tabla,
    }


# =============================================================================
# PRUEBA RÁPIDA: se ejecuta solo con  python credito.py
# Caso del escenario 1: casa de 2,000,000, Adquisición Tradicional, PO 20 años.
# =============================================================================
if __name__ == "__main__":
    r = calcular_credito(producto=1, monto=2_000_000)
    print("Producto:", r["nombre_producto"], "|", r["esquema"], r["plazo"], "años")
    print("Crédito máximo:        {:>15,.2f}   (Excel: 1,900,000.00)".format(r["credito_maximo"]))
    print("Enganche:              {:>15,.2f}   (Excel:   100,000.00)".format(r["enganche"]))
    print("Gastos de originación: {:>15,.2f}   (Excel:   129,500.00)".format(r["gastos_originacion"]))
    print("Inversión del cliente: {:>15,.2f}   (Excel:   229,500.00)".format(r["inversion_cliente"]))
    print("Ingreso mínimo:        {:>15,.2f}   (Excel:    40,970.91)".format(r["ingreso_minimo"]))
    print("Avisos:", r["avisos"])

    # Con estos datos se arma la tabla, igual que antes:
    from datetime import date
    tabla = amortizacion.generar_tabla(fecha_calculo=date(2026, 11, 18), **r["datos_tabla"])
    print("Mensualidad del mes 1: {:>15,.2f}   (Excel:    24,294.00)".format(tabla.loc[0, "Pago total"]))
