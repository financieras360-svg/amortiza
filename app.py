"""
app.py
======
LA PÁGINA WEB DE LA CALCULADORA (Streamlit)

¿Qué hace este archivo?
    Dibuja la calculadora "amortiza" con el diseño que aprobó el equipo.
    NO calcula nada por su cuenta: todo número sale de nuestros archivos
    validados contra el Excel:

        credito.py      -> crédito, gastos, enganche, avisos
        amortizacion.py -> tabla mes a mes
        indicadores.py  -> CAT, plazo real, totales
        extras.py       -> meta, pagar menos, línea del tiempo, "¿por qué?"
        estilos.py      -> colores, tarjetas y gráficas (solo apariencia)

Cómo correrla (en la terminal, dentro de la carpeta del proyecto):
    streamlit run app.py

Cómo funciona Streamlit (lo único que hay que saber):
    Cada vez que el usuario cambia algo, Streamlit vuelve a correr este archivo
    de arriba a abajo. Por eso no hay botón "Calcular": los resultados siempre
    están al día. Lo que debe recordarse entre corridas (por ejemplo, la lista
    de prepagos) se guarda en st.session_state, que es un diccionario.
"""

from datetime import date

import streamlit as st

import amortizacion
import credito
import extras
import indicadores
import parametros
from estilos import (CSS, anios_meses, aviso, fecha_corta, fecha_dmy, fila, money,
                     num, paso, pct, svg_barra_tiempo, svg_planes, svg_saldos,
                     tarjeta, tex_num)


# =============================================================================
# 0. CONFIGURACIÓN Y CATÁLOGOS DE LA PANTALLA
# =============================================================================
st.set_page_config(page_title="amortiza · Simulador de crédito hipotecario", layout="wide")
st.markdown(CSS, unsafe_allow_html=True)        # aplica el diseño (estilos.py)

# Paso 1: desde dónde calcula el usuario (valor interno -> texto en pantalla)
MODOS = {
    "ingresos": "Ingreso mensual",
    "valor": "Valor del inmueble",
    "monto": "Crédito solicitado",
    "mensualidad": "Mensualidad deseada",
}
PREGUNTA_PRINCIPAL = {
    "ingresos": "¿Cuánto ganas al mes?",
    "valor": "¿Cuánto cuesta la vivienda?",
    "monto": "¿Cuánto necesitas que te presten?",
    "mensualidad": "¿Cuánto quieres pagar al mes?",
}
VALOR_INICIAL = {"ingresos": 50_000.0, "valor": 2_000_000.0, "monto": 1_900_000.0, "mensualidad": 24_294.0}

# Paso 2: destino + modalidad = producto del simulador (1 a 7)
DESTINOS = {
    "compra": "Comprar vivienda",
    "pasivos": "Pagar mi hipoteca de otro banco",
    "liquidez": "Liquidez con mi vivienda",
}
MODALIDADES = {
    "normal": "Normal",
    "apoyo": "Apoyo Infonavit",
    "cofinavit": "Cofinavit",
    "cicb": "CI+CB Cofinavit",
}
PRODUCTO_DE = {
    "compra": {"normal": 1, "apoyo": 2, "cofinavit": 3, "cicb": 7},
    "pasivos": {"normal": 4, "apoyo": 5},
    "liquidez": {"normal": 6},
}


# =============================================================================
# 1. FUNCIONES DE APOYO PARA LA PANTALLA
# =============================================================================

def html(contenido):
    """Inserta HTML del diseño (st.html no interpreta $ como fórmula)."""
    st.html(contenido)


def llave(nombre):
    """
    Llave de cada control con la "versión" actual. Al restablecer, la versión
    sube y todos los controles nacen de nuevo con sus valores iniciales
    (borrar las llaves no basta: algunos controles recuerdan su opción).
    """
    return "{}_v{}".format(nombre, st.session_state.get("version", 0))


def restablecer():
    """Vuelve todo a los valores iniciales del Excel."""
    st.session_state.version = st.session_state.get("version", 0) + 1
    st.session_state.prepagos = {}


def agregar_prepago(llave_mes, llave_monto):
    mes, monto = int(st.session_state[llave_mes]), float(st.session_state[llave_monto])
    if monto > 0:
        st.session_state.prepagos[mes] = st.session_state.prepagos.get(mes, 0) + monto


def quitar_prepago(mes):
    st.session_state.prepagos.pop(mes, None)


def clasificar(texto):
    """Nivel visual de cada aviso del Excel."""
    if texto.startswith(("Valor de vivienda", "Ingresos menores", "Crédito fuera", "Crédito debajo")):
        return "err"
    if texto.startswith("Negativo"):
        return "ok"
    if texto.startswith(("Este producto", "La campaña")):
        return "info"
    return "warn"


EXPLICACION_AVISO = {
    "Valor de vivienda debajo del mínimo": "Este crédito pide que la casa valga al menos {}.",
    "Ingresos menores al mínimo": "El ingreso mínimo de este producto es {}.",
    "Crédito fuera de rango": "Pediste más que tu crédito máximo.",
    "Línea máxima a otorgar puede ser menor por perfil del cliente": "El banco puede prestarte menos según tu perfil. Tu cotización no cambia.",
    "El crédito solicitado debe ser igual al máximo": "El monto pasa del máximo de política del banco.",
}


# Lo que se guarda entre corridas
st.session_state.setdefault("prepagos", {})
st.session_state.setdefault("version", 0)


# =============================================================================
# 2. ENCABEZADO
# =============================================================================
html('<div class="am-header"><div class="am-logo"><b>amortiza</b>'
     '<span>Simulador de crédito hipotecario</span></div></div>')

vista_calc, vista_info = st.tabs(["Calculadora", "Más información"])

# #############################################################################
# VISTA 1: CALCULADORA
# #############################################################################
with vista_calc:
    html('<div class="am-hero"><h1>Simulador de crédito hipotecario</h1>'
         '<div class="pill">Tasa desde 9.50 % anual</div>'
         '<p>Calcula tu mensualidad, tu crédito máximo y tu tabla de pagos en 3 pasos.</p></div>')

    # ---------------------------------------------------------------- PASO 1
    html(paso(1, "¿Desde dónde quieres calcular?"))
    modo = st.segmented_control("Modo de cálculo", list(MODOS), format_func=MODOS.get,
                                default="valor", key=llave("modo"), label_visibility="collapsed") or "valor"
    if modo == "mensualidad":
        st.caption("Función nueva: no existe en el Excel. Busca el crédito cuya mensualidad del mes 1 sea la que escribas.")

    # ---------------------------------------------------------------- PASO 2
    html(paso(2, "¿Para qué es tu crédito?"))
    destino = st.segmented_control("Destino", list(DESTINOS), format_func=DESTINOS.get,
                                   default="compra", key=llave("destino"), label_visibility="collapsed") or "compra"
    opciones_mod = list(PRODUCTO_DE[destino])
    modalidad = st.segmented_control("Modalidad", opciones_mod, format_func=MODALIDADES.get,
                                     default="normal", key=llave("modalidad_" + destino)) or "normal"
    producto = PRODUCTO_DE[destino][modalidad]
    datos_prod = parametros.PRODUCTOS[producto]

    # ---------------------------------------------------------------- PASO 3 (izquierda) + RESULTADO (derecha)
    col_datos, col_resultado = st.columns([1.1, 1], gap="large")

    with col_datos:
        html(paso(3, "Completa tus datos"))

        principal = st.number_input(PREGUNTA_PRINCIPAL[modo], min_value=0.0, step=10_000.0 if modo != "mensualidad" else 500.0,
                                    value=VALOR_INICIAL[modo], format="%.2f", key=llave("principal_" + modo))
        st.caption("= " + money(principal))
        lugar_aviso_principal = st.empty()          # aquí se ponen los avisos de este dato

        esquema = st.segmented_control("¿Cómo quieres pagar?", ["Pagos Oportunos", "Valora"],
                                       default="Pagos Oportunos", key=llave("esquema")) or "Pagos Oportunos"
        st.caption("Pago fijo; la tasa baja si pagas a tiempo." if esquema == "Pagos Oportunos"
                   else "Empiezas pagando menos y la mensualidad sube cada año.")

        plazos = parametros.PLAZOS_PAGOS_OPORTUNOS if esquema == "Pagos Oportunos" else parametros.PLAZOS_VALORA
        plazo = st.segmented_control("¿Cuánto tiempo necesitas para pagar?", plazos, default=20,
                                     format_func=lambda p: "{} años".format(p), key=llave("plazo_" + esquema)) or 20

        nivel = "1"
        if esquema == "Valora":
            nivel = st.segmented_control("Nivel de tasa (tu perfil de riesgo)", parametros.NIVELES_DE_TASA,
                                         default="1", key=llave("nivel")) or "1"
            st.caption("1 es el mejor perfil y R el de mayor riesgo."
                       + (" Con nivel R el monto máximo de política es de 7 millones." if nivel == "R" else ""))

        if producto in (1, 4, 6):
            actividad = st.segmented_control("¿Cómo recibes tus ingresos?", ["Empleado", "Independiente"],
                                             default="Empleado", key=llave("actividad")) or "Empleado"
        else:
            actividad = "Independiente" if producto == 7 else "Empleado"
            st.caption("Actividad: **{}** (fija para este producto).".format(actividad))

        if producto in (2, 3, 5):
            tipo_comision = "Contratación"
            st.caption("Comisión: **por contratación** (este producto no ofrece pago anticipado).")
        else:
            pena = datos_prod["penalizacion_prepago"]
            tipo_comision = st.radio(
                "¿Cuándo pagas la comisión?", ["Pago anticipado", "Contratación"], key=llave("comision"), horizontal=True,
                format_func=lambda t: "Pago anticipado" if t == "Pago anticipado" else "Por contratación",
                captions=["No pagas comisión al firmar." + (" Si abonas antes, pagas 3 % + IVA." if pena else ""),
                          "Pagas {} al firmar. Si abonas antes, no hay penalización.".format(pct(datos_prod["comision_apertura"]))])

        # Campos que solo aparecen en ciertos productos
        sueldo = ssv_t = ssv_c = cinf_t = cinf_c = 0.0
        con_conyuge = False
        financiar = True
        lugar_infonavit_7 = None
        if producto in (2, 5):
            sueldo = st.number_input("Sueldo de cotización (Apoyo Infonavit)", min_value=0.0, value=30_000.0,
                                     step=1_000.0, format="%.2f", key=llave("sueldo"))
            st.caption("Tu patrón aporta 5 % bimestral (tope $8,915.56) y se abona a tu deuda cada 2 meses.")
        if producto == 3:
            ssv_t = st.number_input("Tu subcuenta de vivienda", min_value=0.0, value=0.0, step=10_000.0, key=llave("ssv_t"))
            cinf_t = st.number_input("Tu crédito Infonavit", min_value=0.0, value=0.0, step=10_000.0, key=llave("cinf_t"),
                                     help="Aquí debes capturar el Crédito Neto (Anexo B) cuando se tenga el dato.")
            con_conyuge = st.checkbox("Sumar a mi cónyuge", key=llave("conyuge3"))
            if con_conyuge:
                ssv_c = st.number_input("Subcuenta de vivienda del cónyuge", min_value=0.0, value=0.0, step=10_000.0, key=llave("ssv_c3"))
                cinf_c = st.number_input("Crédito Infonavit del cónyuge", min_value=0.0, value=0.0, step=10_000.0, key=llave("cinf_c3"))
        if producto == 7:
            ssv_t = st.number_input("Tu subcuenta de vivienda", min_value=0.0, value=0.0, step=10_000.0, key=llave("ssv_t7"))
            lugar_infonavit_7 = st.empty()          # el crédito Infonavit se calcula; se muestra después
            con_conyuge = st.checkbox("Sumar a mi cónyuge", key=llave("conyuge7"))
            if con_conyuge:
                ssv_c = st.number_input("Subcuenta de vivienda del cónyuge", min_value=0.0, value=0.0, step=10_000.0, key=llave("ssv_c7"))
            st.caption("* El monto del crédito aplica para el acreditado y conyuge, en su caso.")
        if producto in (4, 5):
            financiar = (st.segmented_control("¿Sumar los gastos notariales a tu crédito? (6 %)", ["Sí", "No"],
                                              default="Sí", key=llave("fin_gastos")) or "Sí") == "Sí"

        # Ajustes avanzados (opcionales)
        with st.expander("Ajustes avanzados (opcional): promoción, gastos, seguros, fecha"):
            st.caption("Ya usamos los valores más comunes. Cámbialos solo si tienes una promoción, otros gastos o seguros propios.")
            campanas = [c for c in parametros.CAMPANAS if producto == 1 or c not in parametros.CAMPANAS_SOLO_PRODUCTO_1]
            a1, a2 = st.columns(2)
            campana = a1.selectbox("Promoción (campaña)", campanas, key=llave("campana_{}".format(len(campanas))))
            fecha = a2.date_input("Fecha de cálculo", value=date.today(), format="DD/MM/YYYY", key=llave("fecha"),
                                  help="Entre el día 3 y el 16 pagas el día 3; si no, el día 17.")
            pct_imp = a1.number_input("% de impuestos y derechos", min_value=0.0, max_value=20.0, value=3.0, step=0.25, key=llave("pct_imp"))
            pct_hon = a2.number_input("% de honorarios notariales", min_value=0.0, max_value=20.0, value=2.0, step=0.25, key=llave("pct_hon"))
            com_manual = a1.number_input("Comisión manual (decimal, ej. 0.0100)", min_value=0.0, max_value=0.5, value=0.0,
                                         step=0.0005, format="%.4f", key=llave("com_manual"), help="Déjala en 0 para usar la del producto.")
            valor_dest = a2.number_input("Valor de la construcción (destructible)", min_value=0.0, value=0.0, step=50_000.0,
                                         key=llave("valor_dest"), help="Construcción + instalaciones especiales, sin terreno. Para el seguro de daños.")
            seg_ext = a1.toggle("Tengo mis propios seguros", key=llave("seg_ext"))
            gastos_exactos = 0.0
            if producto in (3, 7):
                gastos_exactos = a2.number_input("Gastos exactos del Infonavit", min_value=0.0, value=0.0, step=1_000.0, key=llave("gastos_exactos"))
            zona = a2.selectbox("¿Dónde está el inmueble? (informativo)", ["metro", "resto"], key=llave("zona"),
                                format_func=lambda z: "CDMX, Guadalajara o Monterrey" if z == "metro" else "Resto del país")
            lugar_credito_solicitado = st.container()

        st.button("Restablecer", on_click=restablecer, type="secondary")

    # =========================================================================
    # 3. CÁLCULO (todo con nuestro motor validado)
    # =========================================================================
    datos_calculo = dict(
        producto=producto, campana=campana, actividad=actividad, esquema=esquema, plazo=plazo,
        nivel=nivel, tipo_comision=tipo_comision, comision_manual=com_manual if com_manual > 0 else None,
        seguros_externos=seg_ext, valor_destructible=valor_dest, sueldo_cotizacion=sueldo,
        financiar_gastos=financiar, ssv_titular=ssv_t, ssv_conyuge=ssv_c,
        credito_infonavit_titular=cinf_t, credito_infonavit_conyuge=cinf_c,
        porcentaje_impuestos=pct_imp / 100, porcentaje_honorarios=pct_hon / 100,
        gastos_exactos_infonavit=gastos_exactos,
    )

    if modo == "mensualidad":
        monto, modo_motor = extras.credito_para_mensualidad(principal, **datos_calculo), credito.MODO_MONTO
    else:
        monto, modo_motor = principal, modo

    # Crédito solicitado: por omisión el máximo; el usuario puede pedir menos
    credito_solicitado = None
    if principal > 0 and modo in ("valor", "ingresos"):
        maximo = credito.calcular_credito(monto=monto, modo=modo_motor, **datos_calculo)["credito_maximo"]
        with lugar_credito_solicitado:
            if st.checkbox("Quiero pedir un crédito distinto al máximo", key=llave("pedir_menos")):
                credito_solicitado = st.number_input("Crédito solicitado", min_value=0.0, value=float(round(maximo, 2)),
                                                     step=10_000.0, format="%.2f", key=llave("cred_sol"),
                                                     help="Aquí puede capturar el monto del crédito que desee.")

    hay_calculo = principal > 0
    if hay_calculo:
        r = credito.calcular_credito(monto=monto, modo=modo_motor, credito_solicitado=credito_solicitado, **datos_calculo)
        reducir = st.session_state.get(llave("reducir")) or "plazo"
        tabla = amortizacion.generar_tabla(fecha_calculo=fecha, prepagos=st.session_state.prepagos,
                                           reducir=reducir, **r["datos_tabla"])
        ind = indicadores.calcular_indicadores(r, tabla, fecha) if r["credito_para_tabla"] > 0 else None
        hay_calculo = ind is not None

    # =========================================================================
    # 4. PANEL DE RESULTADO (derecha)
    # =========================================================================
    with col_resultado:
        if not hay_calculo:
            html('<div class="am-idle"><div style="font-size:22px;font-weight:700">Calcula tu crédito con los datos que ingresaste</div>'
                 '<p style="color:#6A6A6A">Escribe una cantidad en el paso 3.</p></div>')
        else:
            # Avisos junto al dato principal (rojos) y en el panel (amarillos)
            avisos_campo, avisos_panel = [], []
            for a in r["avisos"]:
                nivel_aviso = clasificar(a)
                extra = EXPLICACION_AVISO.get(a, "")
                if a.startswith("Valor de vivienda"):
                    extra = extra.format(money(datos_prod["valor_minimo_vivienda"], 0))
                elif a.startswith("Ingresos menores"):
                    extra = extra.format(money(datos_prod["ingreso_minimo"], 0))
                bloque = aviso(nivel_aviso, a, extra)
                (avisos_campo if nivel_aviso == "err" else avisos_panel).append(bloque)
            if avisos_campo:
                with lugar_aviso_principal:
                    html("".join(avisos_campo))

            if lugar_infonavit_7 is not None:
                lugar_infonavit_7.caption("Crédito Infonavit (calculado): **{}** · 30 % del monto otorgado, tope $885,445.60".format(
                    money(r["credito_infonavit"])))

            mes1 = tabla.iloc[0]
            invalida = any(a.startswith("Valor de vivienda") for a in r["avisos"])
            encabezado = "{} · {} · {} años".format(r["nombre_producto"], esquema, plazo)

            if r["fuera_de_rango"]:
                cifra = aviso("err", "Crédito solicitado fuera de rango",
                              "Pediste {} y tu máximo es {}. Baja el monto para ver tu mensualidad.".format(
                                  money(r["credito_solicitado"]), money(r["credito_maximo"])))
            elif modo == "ingresos":
                regla = ""
                if esquema == "Valora" and nivel in ("4", "R"):
                    regla = " Con Valora nivel {}: hasta 25 veces tu ingreso.".format(nivel)
                cifra = ('<div class="lbl">Crédito máximo posible</div><div class="big">{}</div>'
                         '<div class="lbl">Con tu ingreso de {} al mes.{}</div>'
                         '<p class="sentence">Tu mensualidad del mes 1 sería de <b>{}</b> (pago + seguros).</p>').format(
                    money(r["credito_maximo"], 0), money(principal, 0), regla, money(mes1["Pago total"]))
            else:
                frase = "Por un crédito de {} a {} años, pagarías {} al mes durante el primer año".format(
                    money(r["credito_para_tabla"], 0), plazo, money(mes1["Pago total"]))
                frase += "; después sube {} cada año.".format(pct(r["incremento_anual"], 1)) if esquema == "Valora" else "."
                cifra = ('<div class="lbl">Tu mensualidad · mes 1</div><div class="big">{}</div>'
                         '<div class="lbl">Pago + seguro de vida + seguro de daños</div><p class="sentence">{}</p>').format(
                    money(mes1["Pago total"]), frase)

            etiqueta_invalida = ('<div style="margin-top:6px"><span style="background:#FDECEE;color:#A50010;border:1px solid #A50010;'
                                 'border-radius:999px;padding:2px 10px;font-size:12px;font-weight:700">Cotización no válida</span></div>'
                                 if invalida else "")
            kpis = ('<div class="am-kpis"><div><b>{}</b><span>Crédito solicitado</span></div>'
                    '<div><b>{}</b><span>CAT sin IVA</span></div>'
                    '<div><b>{}</b><span>Inversión del cliente</span></div></div>').format(
                money(r["credito_para_tabla"], 0), pct(ind["cat"], 1), money(r["inversion_cliente"], 0))
            html('<div class="am-result" style="{}"><div class="eyebrow">{}</div>{}{}{}{}</div>'.format(
                "opacity:.6" if invalida else "", encabezado, etiqueta_invalida, cifra,
                kpis if not r["fuera_de_rango"] else "", "".join(avisos_panel)))

            if not r["fuera_de_rango"]:
                with st.expander("¿De dónde sale?"):
                    d = r["datos_tabla"]
                    st.latex(r"M_1 = C\cdot\frac{p}{1000} + 0.0006\,C + 0.00031\,V")
                    st.latex(r"M_1 = {}\cdot\frac{{{}}}{{1000}} + {} + {} = \${}".format(
                        tex_num(d["credito"]), d["pago_al_millar"], tex_num(mes1["Seguro de vida"]),
                        tex_num(mes1["Seguro de daños"]), tex_num(mes1["Pago total"])))
                    st.caption("C = crédito · p = pago al millar · V = valor de la vivienda (o valor destructible).")
            st.caption("CAT: costo anual total del crédito en %; sirve para comparar créditos entre sí. "
                       "Tus resultados se actualizan solos al cambiar cualquier dato.")

    # =========================================================================
    # 5. RESULTADOS DETALLADOS (pestañas)
    # =========================================================================
    if hay_calculo and not r["fuera_de_rango"]:
        html('<div class="am-banner">Tu cotización para una casa de {}</div>'.format(money(r["valor_vivienda"], 0)))
        # ¿Los abonos extra pagan penalización? (comisión «Pago anticipado» y producto con 3 %)
        cobra_pena = r["datos_tabla"]["cobra_penalizacion"] and r["datos_tabla"]["tasa_penalizacion"] > 0
        pes_resumen, pes_tabla, pes_tiempo, pes_planea, pes_caratula = st.tabs(
            ["Resumen y cotización", "Tabla de amortización", "Línea del tiempo", "Planea tus pagos", "Carátula"])

        # --------------------------------------------------------- RESUMEN
        with pes_resumen:
            es_cofin = producto in (3, 7)
            detalle = [fila("Programa", r["nombre_producto"]),
                       fila("Esquema · plazo", "{} · {} años".format(esquema, plazo)),
                       fila("Campaña", r["campana"]),
                       fila("Actividad económica", r["actividad"])]
            if esquema == "Valora":
                detalle.append(fila("Nivel de tasa", r["nivel"]))
            detalle.append(fila("% de financiamiento" + (" CI+CB" if producto == 7 else ""), pct(r["porcentaje_financiamiento"])))

            recursos = [fila("Crédito máximo posible", num(r["credito_maximo"])),
                        fila("Crédito solicitado", num(r["credito_solicitado"]))]
            if producto in (4, 5):
                recursos += [fila("Gastos notariales financiados", num(r["gastos_financiados"])),
                             fila("Monto máximo financiado", num(r["credito_para_tabla"]))]
            if es_cofin:
                recursos += [fila("Subcuenta de vivienda", num(r["subcuenta_vivienda"])),
                             fila("Crédito Infonavit", num(r["credito_infonavit"])),
                             fila("Recursos totales", num(r["recursos_totales"]))]
            if producto != 6:
                recursos.append(fila(r["etiqueta_enganche"] or "Enganche", num(r["enganche"])))
            recursos += [fila("Gastos de originación", num(r["gastos_originacion"])),
                         fila("Inversión del cliente", num(r["inversion_cliente"]), total=True)]

            gastos = [fila("Valor de vivienda", num(r["valor_vivienda"]))]
            if es_cofin:
                gastos.append(fila("Infonavit gastos de originación · 3 %", num(r["comision_infonavit"])))
            gastos += [fila("Impuestos y derechos · {}".format(pct(pct_imp / 100)), num(r["impuestos_y_derechos"])),
                       fila("Honorarios notariales · {}".format(pct(pct_hon / 100)), num(r["honorarios_notariales"])),
                       fila("Avalúo" + (" (favor de cotizar)" if r["valor_vivienda"] > parametros.AVALUO_LIMITE_COTIZAR else ""),
                            num(r["avaluo"]))]
            if tipo_comision == "Pago anticipado":
                gastos.append(fila("Comisión por contratación (pago anticipado)", num(0)))
            else:
                gastos.append(fila("Comisión por contratación", num(r["comision_contratacion"])))
            if producto in (4, 5):
                gastos.append(fila("Menos gastos financiados", num(-r["gastos_financiados"])))
            gastos.append(fila("Total", num(r["gastos_originacion"]), total=True))

            resultados = [fila("Tasa anual", pct(r["tasa"])),
                          fila("Pago al millar", num(r["pago_al_millar"])),
                          fila("Mensualidad sin seguros", num(mes1["Pago"])),
                          fila("Mensualidad mes 1 (con seguros)", num(mes1["Pago total"]))]
            if esquema == "Valora":
                resultados.append(fila("Incremento anual", pct(r["incremento_anual"], 1)))
            resultados += [fila("Ingreso mínimo requerido", num(r["ingreso_minimo"])),
                           fila("Penalización si abonas antes", "3 % + IVA" if cobra_pena else "0 %"),
                           fila("CAT sin IVA", pct(ind["cat"], 1), total=True)]

            parametros_prod = [fila("Crédito mínimo", num(r["credito_minimo"])),
                               fila("Valor mínimo del inmueble", num(r["valor_minimo_vivienda"])),
                               fila("Enganche mínimo", pct(r["enganche_minimo"], 0)),
                               fila("PTI (parte del ingreso para el pago)", pct(datos_prod["pti"], 0))]
            minimo_zona = extras.minimo_por_zona(producto, zona)
            if minimo_zona is not None and zona == "resto":
                parametros_prod.append(fila("Mínimo en el resto del país (informativo)", num(minimo_zona)))

            html('<div class="am-grid3">{}{}{}{}{}</div>'.format(
                tarjeta("Detalle", "Qué crédito estás cotizando.", detalle),
                tarjeta("Recursos", "De dónde sale el dinero y cuánto pones tú.", recursos),
                tarjeta("Gastos de originación", "Lo que pagas al firmar las escrituras.", gastos),
                tarjeta("Resultados", "Las condiciones de tu crédito.", resultados),
                tarjeta("Parámetros del producto", "Los mínimos que pide este producto.", parametros_prod)))

            with st.expander("Ver observaciones (texto del simulador)"):
                st.markdown(
                    "- Los impuestos y derechos están calculados al {} del valor del inmueble; pueden cambiar dependiendo de la entidad federativa.\n"
                    "- Los honorarios notariales están calculados al {} del valor del inmueble; pueden variar dependiendo de la entidad federativa.\n"
                    "- El crédito solicitado dependerá de la capacidad de pago de cada cliente.\n"
                    "- El banco se reserva el derecho de solicitar información adicional y/o restringir parámetros de enganche, monto, plazo o tasa.\n"
                    "- En caso de que aplique, la comisión de penalización por pago anticipado de crédito causa I.V.A.\n"
                    "- Esta cotización tiene carácter informativo y NO implica que el crédito haya sido autorizado.\n"
                    "- Para los esquemas Pagos Oportunos y Valora se deben cumplir las reglas del programa para la disminución de la tasa."
                    .format(pct(pct_imp / 100), pct(pct_hon / 100)))

        # --------------------------------------------------------- TABLA
        with pes_tabla:
            c1, c2, c3 = st.columns(3)
            c1.html('<div class="am-stat"><div class="t">ESQUEMA</div><div class="v" style="font-size:22px">{} {} AÑOS</div>'
                    '<div class="t" style="font-weight:400">Cotizado el {} · día de pago {}</div></div>'.format(
                        esquema.upper(), plazo, fecha_dmy(fecha), tabla.iloc[0]["Fecha de pago"].day))
            c2.html('<div class="am-stat ok"><div class="t">El crédito se liquida por pago puntual en:</div>'
                    '<div class="v">{} años {} meses</div><div class="t" style="font-weight:400">{} pagos mensuales</div></div>'.format(
                        ind["anios"], ind["meses"], ind["meses_totales"]))
            c3.html('<div class="am-stat"><div class="t">CAT SIN IVA</div><div class="v">{}</div>'
                    '<div class="t" style="font-weight:400">Corrida para CAT: sin prepagos ni aportaciones</div></div>'.format(pct(ind["cat"], 1)))

            html('<div class="am-grid4" style="margin:14px 0">'
                 '<div class="am-spec"><b>Interés:</b> lo que te cobra el banco ese mes por la deuda que aún tienes.</div>'
                 '<div class="am-spec"><b>Amortización:</b> la parte de tu pago que sí baja tu deuda.</div>'
                 '<div class="am-spec"><b>Erogación:</b> interés + amortización. Pago total = erogación + seguros.</div>'
                 '<div class="am-spec"><b>Tip:</b> al principio casi todo es interés; con los años la amortización crece.</div></div>')

            with st.expander("¿Vas a abonar extra? · Prepagos a capital", expanded=bool(st.session_state.prepagos)):
                if cobra_pena:
                    html(aviso("warn", "Con «Pago anticipado» este producto cobra 3 % + IVA sobre cada abono.",
                               "El simulador suma la penalización a tu saldo."))
                p1, p2, p3, p4 = st.columns([1, 1.4, 1, 1.6])
                llave_mes = llave("prep_mes_{}".format(plazo))         # la llave cambia con el plazo (cambia el máximo)
                p1.number_input("Mes", min_value=1, max_value=plazo * 12, value=24, step=1, key=llave_mes)
                p2.number_input("Monto extra", min_value=0.0, value=100_000.0, step=10_000.0, key=llave("prep_monto"))
                p3.button("+ Agregar", on_click=agregar_prepago, args=(llave_mes, llave("prep_monto")))
                p4.segmented_control("Al prepagar, reducir:", ["plazo", "pago"], default="plazo", key=llave("reducir"),
                                     format_func=lambda x: "Plazo" if x == "plazo" else "Pago")
                for mes_p, monto_p in sorted(st.session_state.prepagos.items()):
                    q1, q2 = st.columns([4, 1])
                    q1.write("Mes {} · {}".format(mes_p, money(monto_p)))
                    q2.button("Quitar", key=llave("quitar_{}".format(mes_p)), on_click=quitar_prepago, args=(mes_p,))

            v1, v2 = st.columns([1, 1])
            vista = v1.segmented_control("Vista", ["Mensual", "Resumen anual"], default="Mensual", key=llave("vista"),
                                         label_visibility="collapsed") or "Mensual"
            secundarias = v2.checkbox("Mostrar columnas secundarias (6)", key=llave("secundarias"))

            if vista == "Mensual":
                columnas = ["Mes", "Fecha de pago", "Saldo inicial", "Pago", "Interés", "Amortización", "Tasa anual",
                            "Saldo final", "Pago total"]
                if secundarias:
                    columnas[2:2] = ["Días"]
                    columnas += ["Prepago", "Penalización", "Aportación Infonavit", "Seguro de vida", "Seguro de daños"]
                vista_tabla = tabla[columnas].rename(columns={"Pago": "Erogación", "Tasa anual": "Tasa aplicada",
                                                              "Pago total": "Pago mensual"})
                baja = tabla["Tasa anual"].diff().fillna(0) < -1e-9      # renglones donde bajó la tasa

                def pintar(renglon):
                    if baja.iloc[renglon.name]:
                        return ["background-color:#FDF0F0"] * len(renglon)
                    if renglon.name == len(tabla) - 1:
                        return ["background-color:#E3F0EB"] * len(renglon)
                    return [""] * len(renglon)

                formato = {c: "{:,.2f}" for c in vista_tabla.columns if c not in ("Mes", "Fecha de pago", "Días", "Tasa aplicada")}
                formato["Tasa aplicada"] = "{:.2%}"
                formato["Fecha de pago"] = lambda f: f.strftime("%d/%m/%Y")
                st.dataframe(vista_tabla.style.format(formato).apply(pintar, axis=1), hide_index=True, height=460)
                st.caption("Rosa = mes en que baja la tasa · Verde = último pago (liquidado).")
            else:
                anual = extras.resumen_anual(tabla)
                st.dataframe(anual.style.format({c: "{:,.2f}" for c in anual.columns if c != "Año"}), hide_index=True)

            st.download_button("Descargar tabla (CSV para Excel)", tabla.to_csv(index=False).encode("utf-8-sig"),
                               file_name="tabla_amortizacion.csv", mime="text/csv")

            # ¿Por qué? de un mes
            html('<div class="am-step" style="font-size:18px">¿De dónde sale cada número?</div>')
            mes_q = st.number_input("Elige un mes de la tabla", min_value=1, max_value=len(tabla), value=1, step=1,
                                    key=llave("mes_porque_{}".format(len(tabla))))
            explicaciones = extras.explicar_mes(tabla, r, int(mes_q))
            e1, e2 = st.columns(2)
            for k, e in enumerate(explicaciones):
                with (e1 if k % 2 == 0 else e2):
                    html('<div class="am-why"><b>{} · {}</b><br><span style="font-size:13px;color:#6A6A6A">{}</span>'
                         '<br><code>{}</code><br><span style="font-size:12px;color:#6A6A6A">Columna {} del Excel</span></div>'.format(
                             e["concepto"], money(e["valor"]), e["explicacion"], e["cuenta"], e["columna"]))
                    st.latex(e["latex"])

        # --------------------------------------------------------- LÍNEA DEL TIEMPO
        with pes_tiempo:
            eventos = extras.linea_del_tiempo(tabla, r, fecha)
            colores = {"inicio": "#007EAB", "pago": "#EC111A", "tasa": "#F2A900", "sube": "#F2A900",
                       "infonavit": "#7A2E39", "cruce": "#333333", "mitad": "#333333", "fin": "#00664F"}
            marcas = [(e["mes"], colores[e["tipo"]], e["titulo"]) for e in eventos if e["tipo"] in ("tasa", "cruce", "mitad")]
            html('<div class="am-card">{}</div>'.format(svg_barra_tiempo(len(tabla), plazo * 12, marcas)))
            filas_tl = "".join(
                '<div class="ev"><span class="f">{}</span><span class="dot" style="background:{}"></span>'
                '<div><div class="tt">{}</div><div class="dd">{}</div></div><span class="m">{}</span></div>'.format(
                    fecha_corta(e["fecha"]), colores[e["tipo"]], e["titulo"], e["detalle"],
                    "Hoy" if e["mes"] == 0 else "Mes {}".format(e["mes"]))
                for e in eventos)
            html('<div class="am-tl" style="margin-top:14px">{}</div>'.format(filas_tl))
            st.caption("La línea del tiempo usa tu tabla de amortización, incluidos los abonos extra que hayas agregado.")

        # --------------------------------------------------------- PLANEA TUS PAGOS
        with pes_planea:
            herramienta = st.segmented_control("Herramienta", ["Terminar antes", "Pagar menos por un tiempo"],
                                               default="Terminar antes", key=llave("herramienta"),
                                               label_visibility="collapsed") or "Terminar antes"

            if herramienta == "Terminar antes":
                html(aviso("info", "Tu cotización no cambia",
                           "Hoy terminas de pagar en {}. Elige en cuántos años quieres terminar y cómo quieres abonar.".format(
                               anios_meses(len(tabla)))))
                opciones_anios = list(range(1, (len(tabla) - 1) // 12 + 1))
                if not opciones_anios:
                    st.info("Tu crédito ya termina en menos de un año.")
                else:
                    m1, m2 = st.columns([1, 1.3])
                    anios_meta = m1.select_slider("¿En cuántos años quieres terminar?", opciones_anios,
                                                  value=15 if 15 in opciones_anios else opciones_anios[-1], key=llave("meta_anios"))
                    forma = m2.radio("¿Cómo quieres lograrlo?", ["unico", "mensual"], horizontal=True, key=llave("meta_modo"),
                                     format_func=lambda f: "Un solo pago extra" if f == "unico" else "Un poco más cada mes")
                    mes_extra = 24
                    if forma == "unico":
                        tope_mes = anios_meta * 12 - 1
                        mes_extra = m1.number_input("¿En qué mes harías el pago extra?", min_value=1, max_value=tope_mes,
                                                    value=min(24, tope_mes), step=1, key=llave("meta_mes_{}".format(anios_meta)))
                        m1.caption("= al cumplir {}".format(anios_meses(mes_extra)))
                    if cobra_pena:
                        m2.html(aviso("warn", "Cada abono paga 3 % + IVA de penalización",
                                      "La cifra ya la incluye. Con comisión «por contratación» no pagarías penalización."))

                    meta = extras.meta_de_liquidacion(r, fecha, anios_meta, forma, int(mes_extra), st.session_state.prepagos)
                    if meta:
                        if forma == "unico":
                            frase = ("Si en el <b>mes {}</b> abonas <b>{}</b>, terminas de pagar en <b>{}</b> en lugar de {}, "
                                     "y te ahorras cerca de <b>{}</b> en intereses.").format(
                                mes_extra, money(meta["abono"]), anios_meses(meta["meses_meta"]),
                                anios_meses(meta["meses_base"]), money(meta["ahorro_intereses"], 0))
                            kpi1 = (money(meta["abono"]), "Pago extra (una vez)")
                        else:
                            frase = ("Si pagas <b>{} extra cada mes</b> (tu mensualidad del mes 1 sería <b>{}</b>), terminas en "
                                     "<b>{}</b> en lugar de {}, y te ahorras cerca de <b>{}</b> en intereses.").format(
                                money(meta["abono"]), money(meta["nueva_mensualidad"]), anios_meses(meta["meses_meta"]),
                                anios_meses(meta["meses_base"]), money(meta["ahorro_intereses"], 0))
                            kpi1 = (money(meta["abono"]), "Extra cada mes")
                        html('<div class="am-result"><div class="eyebrow">Tu resultado</div><p class="sentence" style="font-size:17px">{}</p>'
                             '<div class="am-kpis"><div><b>{}</b><span>{}</span></div><div><b>{} meses</b><span>Terminas antes</span></div>'
                             '<div><b>{}</b><span>Intereses que te ahorras</span></div></div></div>'.format(
                                 frase, kpi1[0], kpi1[1], meta["meses_ahorrados"], money(meta["ahorro_intereses"], 0)))

                        def saldos_por_anio(t):
                            return [t.iloc[min(12 * a, len(t)) - 1]["Saldo final"] for a in range(1, (len(tabla) + 11) // 12 + 1)]
                        html('<div class="am-card" style="margin-top:14px"><h4>Lo que te falta por pagar, año con año</h4>{}</div>'.format(
                            svg_saldos([("Sin meta", "#333333", saldos_por_anio(meta["tabla_base"])),
                                        ("Con tu meta", "#EC111A", saldos_por_anio(meta["tabla_meta"]))])))

                        lado = [("Terminas de pagar en", anios_meses(meta["meses_base"]), anios_meses(meta["meses_meta"])),
                                ("Último pago", fecha_corta(meta["fecha_fin_base"]), fecha_corta(meta["fecha_fin_meta"])),
                                ("Pago extra", "—", money(meta["abono"]) + (" cada mes" if forma == "mensual" else " en el mes {}".format(mes_extra))),
                                ("Penalización por abonar", "—", money(meta["penalizacion_total"])),
                                ("Intereses en total", money(meta["intereses_base"]), money(meta["intereses_meta"]))]
                        html('<div class="am-card" style="margin-top:14px"><h4>Lado a lado</h4>' +
                             '<div class="am-row" style="font-weight:700"><span>&nbsp;</span><span>Tu cotización · Con tu meta</span></div>' +
                             "".join('<div class="am-row"><span>{}</span><span>{} · <b style="color:#C70E17">{}</b></span></div>'.format(*f) for f in lado) +
                             "</div>")

                        with st.expander("¿Cómo se calcula? (ecuación de valor)"):
                            st.markdown("Con la fecha focal en hoy, lo que debes es igual al valor presente de lo que pagarás:")
                            st.latex(r"P = R\,a_{\overline{n}|} + X\,v^{k}")
                            st.markdown("Despejando con el enfoque prospectivo, el pago extra vale lo mismo que **las "
                                        "mensualidades que te ahorras**, traídas al mes del abono:")
                            st.latex(r"X = R\,v^{\,n-k}\,a_{\overline{N-n}|}")
                            st.markdown("Y si abonas un poco cada mes:")
                            st.latex(r"P = (R+E)\,a_{\overline{n}|}\quad\Rightarrow\quad E = \frac{P}{a_{\overline{n}|}} - R")
                            st.caption("La cifra exacta sale del mismo motor que replica al Excel: probamos montos (bisección) "
                                       "hasta que la tabla termina justo en el mes meta. La fórmula explica el concepto; el banco "
                                       "usa días reales / 360 y bajadas de tasa, por eso no coincide al centavo.")
                    else:
                        st.info("Con esa meta ya terminas a tiempo.")

            else:  # Pagar menos por un tiempo
                html(aviso("info", "Tu cotización no cambia · Ecuación de valor",
                           "Para bajar tu mensualidad una temporada, lo que dejas de pagar debe valer lo mismo, "
                           "en una fecha focal, que lo que pagas para compensarlo."))
                g1, g2, g3, g4 = st.columns(4)
                tope_D = max(100.0, float(mes1["Pago"]))
                D = g1.number_input("¿Cuánto menos al mes?", min_value=100.0, max_value=tope_D,
                                    value=min(3_000.0, tope_D), step=500.0, key=llave("pm_D_{:.0f}".format(tope_D)))
                tope_a = max(2, len(tabla) - 2)
                a = g2.number_input("¿Desde qué mes?", min_value=2, max_value=tope_a, value=min(13, tope_a), step=1,
                                    key=llave("pm_a_{}".format(len(tabla))))
                tope_n = max(1, len(tabla) - int(a))
                n = g3.number_input("¿Durante cuántos meses?", min_value=1, max_value=tope_n, value=min(24, tope_n), step=1,
                                    key=llave("pm_n_{}_{}".format(len(tabla), int(a))))
                m = g4.number_input("Si lo repartes, ¿en cuántos meses?", min_value=1, max_value=240, value=24, step=1, key=llave("pm_m"))
                pm = extras.pagar_menos(D, int(a), int(n), int(m), r["tasa"], con_penalizacion=cobra_pena)
                b = int(a + n - 1)
                if b + m > len(tabla):
                    html(aviso("warn", "El reparto termina después de tu último pago",
                               "Tu crédito acaba en el mes {}; elige menos meses.".format(len(tabla))))

                opcion = st.segmented_control("Opción", ["A", "B", "C"], default="A", key=llave("pm_opt"),
                                              format_func={"A": "A · Abonar antes", "B": "B · Pagar después",
                                                           "C": "C · Repartir después"}.get) or "A"
                tarjetas = [
                    ("A", "A · Abonar antes", "Un pago en el mes {} que paga por adelantado la diferencia.".format(pm["mes_A"]), money(pm["A"])),
                    ("B", "B · Pagar después en una sola vez", "Un pago en el mes {}, al terminar la temporada.".format(pm["mes_B"]), money(pm["B"])),
                    ("C", "C · Repartir después", "Pagas un poco más del mes {} al {}.".format(pm["desde_C"], pm["hasta_C"]), "+{} al mes".format(money(pm["C"]))),
                ]
                html('<div class="am-grid3">{}</div>'.format("".join(
                    '<div class="am-opt{}"><h5>{}</h5><p>{}</p><b>{}</b></div>'.format(" on" if k == opcion else "", t, p, v)
                    for k, t, p, v in tarjetas)))
                if cobra_pena:
                    html(aviso("warn", "Con «Pago anticipado» cada pago extra lleva 3 % + IVA",
                               "Las cifras ya están divididas entre 0.9652 para cubrirlo."))

                tramo = (int(a), b, "#FDEDEB", "pagas {} menos · meses {} a {}".format(money(D, 0), int(a), b))
                filas_svg = [
                    ("A · Abonar antes", [tramo], pm["mes_A"], "{} · mes {}".format(money(pm["A"], 0), pm["mes_A"])),
                    ("B · Pagar después", [tramo], pm["mes_B"], "{} · mes {}".format(money(pm["B"], 0), pm["mes_B"])),
                    ("C · Repartir", [tramo, (b + 1, b + int(m), "#F7C9C5", "+{} · meses {} a {}".format(money(pm["C"], 0), b + 1, b + int(m)))], None, ""),
                ]
                html('<div class="am-card" style="margin-top:14px"><h4>Línea del tiempo de tus pagos</h4>{}</div>'.format(
                    svg_planes(mes1["Pago"], filas_svg)))

                with st.expander("La ecuación de valor, paso a paso", expanded=True):
                    i = pm["i"]
                    st.latex(r"i = \frac{{{}}}{{12}} = {:.6f}".format(pct(r["tasa"]).replace(" %", r"\%"), i))
                    st.markdown("**A.** Fecha focal en el mes {}: lo que dejas de pagar, traído a ese mes.".format(pm["mes_A"]))
                    st.latex(r"X_A = D\,a_{{\overline{{n}}|i}} = {}\cdot\frac{{1-(1+{:.6f})^{{-{}}}}}{{{:.6f}}} = \${}".format(
                        tex_num(D), i, int(n), i, tex_num(pm["A"])))
                    st.markdown("**B.** Fecha focal en el mes {}: lo que dejas de pagar, llevado al final del tramo.".format(pm["mes_B"]))
                    st.latex(r"X_B = D\,s_{{\overline{{n}}|i}} = {}\cdot\frac{{(1+{:.6f})^{{{}}}-1}}{{{:.6f}}} = \${}".format(
                        tex_num(D), i, int(n), i, tex_num(pm["B"])))
                    st.markdown("**C.** Repartes X_B en {} pagos iguales:".format(int(m)))
                    st.latex(r"E\,a_{{\overline{{m}}|i}} = D\,s_{{\overline{{n}}|i}} \;\Rightarrow\; E = {}\cdot\frac{{{:.4f}}}{{{:.4f}}} = \${}".format(
                        tex_num(D), pm["s_n"], pm["a_m"], tex_num(pm["C"])))
                    st.markdown("**¿Por qué B cuesta más que A?** Porque pagas después y el dinero genera interés mientras tanto:")
                    st.latex(r"X_B = X_A\,(1+i)^{n}")
                    st.caption("Se usa la tasa nominal mensual del curso (tasa ÷ 12). El banco calcula con días reales / 360 y "
                               "bajadas de tasa, por eso en la tabla puede variar unos pesos. Pagar menos requiere un convenio "
                               "con el banco; aquí solo se mide su costo.")

        # --------------------------------------------------------- CARÁTULA
        with pes_caratula:
            clave, dato_reca = extras.reca(producto, esquema, con_conyuge and (ssv_c > 0 or cinf_c > 0))
            nombre_comercial = "{} {} {} años".format(extras.PROGRAMA[producto], esquema.upper(), plazo)
            pena_texto = "0%" if tipo_comision == "Contratación" else "3%"
            tardio = "$350.00" if clave == "7-2-2" else "$500.00"
            html(
                '<div class="am-card" style="max-width:860px;margin:auto">'
                '<h4 style="font-size:22px">Carátula de Crédito</h4><div class="sub">Réplica académica · cotización informativa</div>'
                + fila("Nombre comercial del producto", nombre_comercial)
                + fila("Tipo de crédito", "Crédito Hipotecario")
                + '<div class="am-grid4" style="margin:14px 0">'
                + '<div class="am-spec"><b>CAT</b><br><span style="font-size:22px;font-weight:800">{}</span><br>Sin IVA. Para fines informativos y de comparación</div>'.format(pct(ind["cat"], 1))
                + '<div class="am-spec"><b>Tasa de interés anual ordinaria</b><br><span style="font-size:22px;font-weight:800">{}</span><br>Tasa fija. No aplica tasa de interés moratoria</div>'.format(pct(r["tasa"]))
                + '<div class="am-spec"><b>Monto del crédito</b><br><span style="font-size:22px;font-weight:800">{}</span></div>'.format(money(r["credito_para_tabla"]))
                + '<div class="am-spec"><b>Monto total a pagar</b><br><span style="font-size:22px;font-weight:800">{}</span><br>Moneda: pesos mexicanos</div>'.format(money(ind["total_pagado"]))
                + "</div>"
                + fila("Plazo del crédito", "{} años".format(plazo))
                + fila("Fecha de corte y de pago", "Día {} de cada mes".format(tabla.iloc[0]["Fecha de pago"].day))
                + fila("Comisión por contratación", money(r["comision_contratacion"]))
                + fila("Penalización por pago anticipado de crédito", "{} más IVA sobre el monto anticipado".format(pena_texto))
                + fila("Penalización por pago tardío", "{} más IVA".format(tardio))
                + fila("Seguros obligatorios", "Vida y desempleo · Daños y responsabilidad civil")
                + fila("Registro de Contratos de Adhesión (RECA)", dato_reca[1] if dato_reca else "Revisa los datos ingresados")
                + '<p style="font-size:13px;color:#6A6A6A;margin-top:12px"><b>Advertencias:</b> incumplir tus obligaciones te puede generar '
                  'comisiones e intereses moratorios. Contratar créditos que excedan tu capacidad de pago afecta tu historial crediticio. '
                  'El avalista, obligado solidario o coacreditado responderá como obligado principal por el total del pago.</p>'
                + "</div>")
            st.caption("La penalización por pago anticipado se muestra como en el Excel: 0 % con comisión por contratación y 3 % con pago anticipado.")

    html('<div class="am-footer"><span><b>amortiza</b> · Réplica académica del simulador Scotiabank v2026.06 · '
         'Matemáticas Financieras I, FES Acatlán</span><span>Cotización informativa; no implica autorización del crédito.</span></div>')


# #############################################################################
# VISTA 2: MÁS INFORMACIÓN
# #############################################################################
with vista_info:
    st.markdown("## Más información")
    st.caption("Las palabras clave sin tecnicismos, la teoría del curso detrás de cada número y los datos del proyecto.")
    glosario, formulas, acerca = st.tabs(["01 · Glosario", "02 · Fórmulas del curso", "03 · Acerca del proyecto"])

    with glosario:
        terminos = [
            ("CAT", "Costo anual total del crédito, en porcentaje: intereses + comisión + avalúo + seguros, sin IVA. Sirve para comparar créditos."),
            ("Pago al millar", "Cuánto pagas al mes por cada $1,000 que te prestan. Mensualidad = crédito ÷ 1,000 × pago al millar."),
            ("Aforo (% de financiamiento)", "Qué parte del valor de la vivienda presta el banco. El resto es tu enganche."),
            ("PTI", "La parte máxima de tu ingreso que puede irse a la mensualidad (55 % o 65 % según el producto)."),
            ("Subcuenta de vivienda (SSV)", "Tu ahorro en el Infonavit. En Cofinavit y CI+CB se suma a tu crédito."),
            ("Pagos Oportunos", "Pago fijo. La tasa baja en los meses 13, 25, 37 y 49 si pagas a tiempo; por eso terminas antes del plazo."),
            ("Valora", "Tu mensualidad empieza más baja y sube un poco cada año; la tasa baja 0.25 % al año desde el mes 37 hasta un mínimo."),
            ("Prepago (abono a capital)", "Dinero extra que va directo a tu deuda, para terminar antes o pagar menos al mes."),
            ("Ecuación de valor", "Para comparar dinero de fechas distintas, todo se lleva a una misma fecha (fecha focal)."),
        ]
        html('<div class="am-grid3">' + "".join(
            '<div class="am-card"><h4>{}</h4><p style="font-size:13.5px;color:#6A6A6A;margin:6px 0 0">{}</p></div>'.format(t, d)
            for t, d in terminos) + "</div>")

    with formulas:
        st.markdown("**Factor de descuento** — trae a valor presente un peso pagadero en un periodo.")
        st.latex(r"v = (1+i)^{-1}")
        st.markdown("**Valor presente y monto de una anualidad vencida** de n pagos de 1.")
        st.latex(r"a_{\overline{n}|i} = \frac{1-v^{n}}{i} \qquad s_{\overline{n}|i} = \frac{(1+i)^{n}-1}{i}")
        st.markdown("**Saldo insoluto** — prospectivo (lo que falta) y retrospectivo (lo prestado menos lo pagado).")
        st.latex(r"S_k = R\,a_{\overline{n-k}|i} \qquad S_k = P(1+i)^{k} - R\,s_{\overline{k}|i}")
        st.markdown("**Interés y saldo del banco, mes a mes** (año de 360 días y días reales del mes).")
        st.latex(r"I_k = S_{k-1}\cdot\frac{r_k}{360}\cdot d_k \qquad A_k = R_k - I_k \qquad S_k = S_{k-1} - A_k - P_k + \text{pen}_k - \text{ap}_k")
        st.markdown("**CAT** — tasa interna de retorno de una ecuación de valor con fecha focal en 0.")
        st.latex(r"C - \text{comisión} - \text{avalúo} = \sum_{t=1}^{n}\frac{F_t}{(1+r)^{t}} \qquad \text{CAT} = (1+r)^{12} - 1")
        st.markdown("**Meta de liquidación** — el abono vale lo mismo que las mensualidades que te ahorras.")
        st.latex(r"P = R\,a_{\overline{n}|} + X\,v^{k} \;\Rightarrow\; X = R\,v^{\,n-k}\,a_{\overline{N-n}|}")
        st.info("Ejemplo del curso: i = 10 %, R = 100, N = 5, n = 3, k = 1  ⇒  X = 100·(1.10)⁻²·a₂ = 143.43")

    with acerca:
        html('<div class="am-card" style="max-width:720px">'
             + fila("Asignatura", "Matemáticas Financieras I")
             + fila("Institución", "Facultad de Estudios Superiores Acatlán (UNAM)")
             + fila("Profesor", "Arturo Vera Moreno")
             + fila("Semestre", "2027-I")
             + fila("Integrantes", "Cruz Salazar Alejandra<br>Escobar Enriquez Camila<br>Espinoza De la O Mounserrath<br>"
                    "Salazar Jimenez Astrid Celic<br>Velazquez Contreras Javier")
             + '<p style="margin-top:14px;font-weight:700">El núcleo replica el simulador al centavo; las funcionalidades adicionales '
               'son capas de análisis que usan el mismo motor, sin modificar los resultados validados.</p></div>')
