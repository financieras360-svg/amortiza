"""
parametros.py
=============
CATÁLOGO DE PARÁMETROS DE LA CALCULADORA HIPOTECARIA

¿Qué es este archivo?
    Es la "base de datos" de la calculadora. Aquí se guardan todos los números
    que el banco ya tiene definidos: tasas, pagos al millar, porcentajes de
    financiamiento, comisiones, costos de avalúo y seguros.

¿Por qué existe como un archivo aparte?
    1. Para no escribir números "sueltos" dentro de las fórmulas. Si el día de
       mañana cambia una tasa, solo se cambia aquí y todo el programa se
       actualiza solo.
    2. Para poder revisar fácilmente que los datos son iguales a los del Excel
       (cada bloque dice de qué celdas salió).
    3. Para separar los DATOS (este archivo) de los CÁLCULOS (los demás archivos).

¿Qué NO hace?
    No calcula nada. Solo guarda valores.

¿Cómo lo usan los demás archivos?
    Al inicio de credito.py o amortizacion.py se escribirá:
        import parametros
    y después se puede leer cualquier dato, por ejemplo:
        parametros.SEGURO_VIDA        ->  0.0006

Convención importante:
    Todos los porcentajes se guardan como decimales.
    13% se escribe 0.13 ;  1.25% se escribe 0.0125.
    Así se pueden usar directamente en las fórmulas (por ejemplo, credito * 0.0125).

Alcance:
    Los 7 productos del simulador, con los esquemas Pagos Oportunos y Valora:
      Sin Infonavit:  1 Adquisición Tradicional, 4 Pasivo Hipotecario, 6 Liquidez Vivienda
      Con Infonavit:  2 Adquisición con Apoyo Infonavit, 5 Pasivo con Apoyo Infonavit,
                      3 Cofinavit, 7 CI+CB Cofinavit
"""

# =============================================================================
# CÓMO FUNCIONA UN DICCIONARIO (lo usamos en casi todo el archivo)
# -----------------------------------------------------------------------------
# Un diccionario es una tabla de consulta: guarda parejas  llave -> valor.
# Se escribe entre llaves { } y cada pareja se separa con dos puntos:
#
#     edades = {"Ana": 20, "Luis": 22}
#     edades["Ana"]        ->  20
#
# Un diccionario puede tener otros diccionarios dentro (diccionarios anidados).
# Para llegar a un dato se van poniendo las llaves una tras otra:
#
#     PRODUCTOS[1]["pti"]  ->  primero busca el producto 1, luego su "pti"  ->  0.55
#
# Lo usamos porque nos permite buscar un dato por su NOMBRE (o número) en lugar
# de recordar en qué posición está, y porque se lee casi como una tabla.
# =============================================================================


# =============================================================================
# BLOQUE 1. PRODUCTOS
# Fuente: hoja Simulador, celdas S32:AC38 (un renglón por producto).
# -----------------------------------------------------------------------------
# ¿Por qué un diccionario de productos?
#   Los 7 productos se calculan con LAS MISMAS fórmulas; lo único que cambia
#   entre ellos son estos valores ("sus ajustes"). Guardándolos aquí, el resto
#   del programa solo tiene que preguntar:  "¿cuál es la comisión del producto
#   que eligió el usuario?"  ->  PRODUCTOS[producto]["comision_apertura"]
#   y no hay que escribir un programa distinto para cada producto.
#
# ¿Cómo está organizado?
#   - La LLAVE de afuera es el número del producto (1 a 7), el mismo número
#     que usa el Excel.
#   - El VALOR es otro diccionario con los ajustes de ese producto.
#
# ¿Qué significa cada ajuste y para qué se usa?
#   nombre                 Nombre que se mostrará en la página.
#
#   comision_apertura      % del crédito que el banco cobra al inicio por
#                          otorgarlo.  Se usa en F3 (gastos) y en F16 (CAT).
#                          Ej.: 0.0125 x 1,900,000 = 23,750 pesos.
#
#   penalizacion_prepago   % que se cobra sobre un pago anticipado si el cliente
#                          eligió el tipo de comisión "Pago anticipado".
#                          Se usa en F13.  Ej.: 0.03 x 50,000 = 1,500 (+ IVA).
#                          En los productos 2, 3 y 7 vale 0: no hay penalización.
#
#   pti                    "Pago sobre ingreso": % máximo del sueldo mensual que
#                          puede destinarse a la mensualidad.
#                          Se usa en F2 (crédito por ingreso) y F8 (ingreso mínimo).
#                          Ej.: con 0.55, quien gana 50,000 puede pagar hasta 27,500.
#
#   ingreso_minimo         Sueldo mensual mínimo que pide el producto, aunque la
#                          mensualidad sea pequeña.  Se usa en F8.
#
#   credito_minimo         Monto más pequeño que el banco presta en este producto.
#   credito_maximo         Monto más grande que el banco presta en este producto.
#                          Ambos se usan para mostrar avisos (validaciones).
#
#   valor_minimo_vivienda  Valor mínimo de la casa para este producto.
#                          Se usa para mostrar avisos.  En los productos de
#                          Pasivo (4 y 5) vale 0 porque el Excel no pide un mínimo.
#
#   cobra_avaluo           True (verdadero) si el producto cobra avalúo,
#                          False (falso) si no.  Los productos de Pasivo (4 y 5)
#                          no lo cobran (fórmula de la celda O30).  Se usa en F3.
#
#   requiere_enganche      True si el cliente está COMPRANDO una casa y debe poner
#                          enganche (valor de la casa - crédito).  Es False en:
#                            - Pasivo (4 y 5): la casa ya es del cliente; el crédito
#                              sirve para pagar una hipoteca que ya tiene.
#                            - Liquidez (6): la casa ya es del cliente; el crédito
#                              es dinero con la casa como garantía.
#                          (fórmulas L35 y L38).  Se usa en F4.
#
#   tipo_infonavit         Cómo participa el Infonavit en el producto:
#                            None               -> no participa (1, 4, 6)
#                            "apoyo"            -> el patrón aporta cada bimestre
#                                                  para pagar el crédito (2, 5).
#                                                  Ver BLOQUE 9.
#                            "cofinanciamiento" -> el banco y el Infonavit prestan
#                                                  juntos para la misma casa (3, 7).
#                                                  Ver BLOQUE 10.
#
#   financia_gastos        True si el cliente puede sumar los gastos notariales
#                          al crédito (solo Pasivo, 4 y 5).  Ver BLOQUE 11.
#
# Nota: 250_000 es lo mismo que 250000. Python permite el guion bajo solo para
# que los números grandes se lean más fácil.
# =============================================================================
PRODUCTOS = {
    1: {
        "nombre": "Adquisición Tradicional",
        "comision_apertura": 0.0125,
        "penalizacion_prepago": 0.03,
        "pti": 0.55,
        "ingreso_minimo": 10_000,
        "credito_minimo": 250_000,
        "credito_maximo": 15_000_000,
        "valor_minimo_vivienda": 400_000,
        "cobra_avaluo": True,
        "requiere_enganche": True,
        "tipo_infonavit": None,
        "financia_gastos": False,
    },
    2: {
        "nombre": "Adquisición con Apoyo Infonavit",
        "comision_apertura": 0.0075,
        "penalizacion_prepago": 0.0,
        "pti": 0.65,
        "ingreso_minimo": 7_500,
        "credito_minimo": 250_000,
        "credito_maximo": 15_000_000,
        "valor_minimo_vivienda": 400_000,
        "cobra_avaluo": True,
        "requiere_enganche": True,
        "tipo_infonavit": "apoyo",
        "financia_gastos": False,
    },
    3: {
        "nombre": "Cofinavit",
        "comision_apertura": 0.0075,
        "penalizacion_prepago": 0.0,
        "pti": 0.65,
        "ingreso_minimo": 7_500,
        "credito_minimo": 120_000,
        "credito_maximo": 15_000_000,
        "valor_minimo_vivienda": 300_000,
        "cobra_avaluo": True,
        "requiere_enganche": True,
        "tipo_infonavit": "cofinanciamiento",
        "financia_gastos": False,
    },
    4: {
        "nombre": "Pasivo Hipotecario",
        "comision_apertura": 0.0,
        "penalizacion_prepago": 0.03,
        "pti": 0.65,
        "ingreso_minimo": 10_000,
        "credito_minimo": 125_000,
        "credito_maximo": 15_000_000,
        "valor_minimo_vivienda": 0,
        "cobra_avaluo": False,
        "requiere_enganche": False,
        "tipo_infonavit": None,
        "financia_gastos": True,
    },
    5: {
        "nombre": "Pasivo Hipotecario con Apoyo Infonavit",
        "comision_apertura": 0.0,
        "penalizacion_prepago": 0.03,
        "pti": 0.65,
        "ingreso_minimo": 10_000,
        "credito_minimo": 125_000,
        "credito_maximo": 15_000_000,
        "valor_minimo_vivienda": 0,
        "cobra_avaluo": False,
        "requiere_enganche": False,
        "tipo_infonavit": "apoyo",
        "financia_gastos": True,
    },
    6: {
        "nombre": "Liquidez Vivienda",
        "comision_apertura": 0.0125,
        "penalizacion_prepago": 0.03,
        "pti": 0.55,
        "ingreso_minimo": 20_000,
        "credito_minimo": 150_000,
        "credito_maximo": 2_500_000,
        "valor_minimo_vivienda": 1_000_000,
        "cobra_avaluo": True,
        "requiere_enganche": False,
        "tipo_infonavit": None,
        "financia_gastos": False,
    },
    7: {
        "nombre": "CI+CB Cofinavit",
        "comision_apertura": 0.0125,
        "penalizacion_prepago": 0.0,
        "pti": 0.55,
        "ingreso_minimo": 10_000,
        "credito_minimo": 250_000,
        "credito_maximo": 15_000_000,
        "valor_minimo_vivienda": 300_000,
        "cobra_avaluo": True,
        "requiere_enganche": True,
        "tipo_infonavit": "cofinanciamiento",
        "financia_gastos": False,
    },
}


# =============================================================================
# BLOQUE 2. AFORO (porcentaje del valor de la casa que presta el banco)
# Fuente: hoja Simulador, celdas AH48:AJ52, X34 y X38.
# Se usa en: F1, y de ahí en F2 (crédito = valor de la casa x aforo).
# -----------------------------------------------------------------------------
# ¿Qué es?
#   El banco nunca presta el 100% de la casa. El aforo dice qué parte presta.
#   Con aforo 0.95 y una casa de 2,000,000, presta 1,900,000 y el cliente pone
#   100,000 de enganche.
#
# ¿De qué depende?
#   - De la actividad del cliente: "Empleado" (tiene patrón) o "Independiente"
#     (trabaja por su cuenta; se le presta un poco menos por tener ingresos
#     menos estables).
#   - Del nivel de tasa: la calificación de riesgo que el banco le da al cliente.
#     "1" = menor riesgo (mejores condiciones) ... "R" = mayor riesgo.
#
# ¿Cómo está organizado?
#   AFORO[actividad][nivel]  ->  aforo         (productos 1, 4 y 6)
#   Ej.: AFORO["Empleado"]["3"]  ->  0.90
#   Más abajo hay tres casos especiales: independientes de Pasivo, productos
#   con Apoyo Infonavit (2 y 5) y productos de cofinanciamiento (3 y 7).
#
# ¿Por qué estos valores y no los de la hoja Pol?
#   En el Excel, la fórmula de aforo tiene una segunda parte (por montos altos)
#   que apunta a una celda vacía y por eso nunca se ejecuta. Estos valores son
#   los que el Excel realmente usa (comprobado con casas de 2, 8, 12 y 18 millones).
#   La hoja Pol tiene otra tabla parecida, pero ninguna fórmula la usa.
# =============================================================================

# Lista con los niveles en orden; servirá para llenar la lista desplegable de
# la página. Se guardan como texto porque "R" no es un número.
NIVELES_DE_TASA = ["1", "2", "3", "4", "R"]

AFORO = {
    "Empleado":      {"1": 0.95, "2": 0.95, "3": 0.90, "4": 0.85, "R": 0.85},
    "Independiente": {"1": 0.90, "2": 0.90, "3": 0.85, "4": 0.75, "R": 0.75},
}

# Excepción del Excel: en el producto 4 (Pasivo Hipotecario) a los
# independientes se les presta más que en los productos 1 y 6.
# (En la celda AJ hay una condición especial "si es producto 4 o 5").
# Se guarda aparte para no repetir toda la tabla de AFORO solo por un caso.
AFORO_INDEPENDIENTE_PASIVO = {"1": 0.95, "2": 0.95, "3": 0.95, "4": 0.85, "R": 0.85}

# Productos con Apoyo Infonavit (2 y 5): usan su propia columna del Excel,
# "Emp. con apoyo" (celdas AH48:AH52). Solo aplican a empleados, porque el
# apoyo viene de las aportaciones que hace el patrón.
AFORO_EMPLEADO_CON_APOYO = {"1": 0.95, "2": 0.95, "3": 0.95, "4": 0.85, "R": 0.85}

# Productos de cofinanciamiento (3 y 7): el aforo es FIJO, no depende del nivel
# ni de la actividad (celdas X34 y X38).  Llave = número de producto.
AFORO_FIJO = {3: 0.95, 7: 0.90}

# En el Excel, los productos 2, 3 y 5 solo permiten elegir "Empleado" como
# actividad económica (celda W30). Se usará para limitar la lista de la página.
PRODUCTOS_SOLO_EMPLEADO = {2, 3, 5}


# =============================================================================
# BLOQUE 3. CAMPAÑAS (promociones del banco)
# Fuente: hoja Simulador, lista V63:V67 y fórmulas O30 (avalúo) y O31 (comisión).
# Se usa en: F3 (gastos) y para elegir la tasa en los bloques 4 y 5.
# -----------------------------------------------------------------------------
# ¿Qué son?
#   Promociones que cambian la tasa y, en algunos casos, eliminan cobros.
#   El Excel no explica a quién aplica cada una; para nosotros solo importan
#   sus efectos en los cálculos.
#
# ¿Por qué se usan conjuntos { } y no listas [ ]?
#   Un conjunto (set) sirve para preguntar rápidamente "¿esta campaña está
#   aquí?". Más adelante el programa hará algo como:
#       si la campaña está en CAMPANAS_SIN_AVALUO  ->  el avalúo vale 0
# =============================================================================

# Lista de campañas, en el mismo orden que el Excel. Servirá para la lista
# desplegable de la página.
CAMPANAS = ["Ninguna", "WM", "Premium", "Vinculación", "Hipoteca Green"]

# Campañas que NO cobran avalúo (fórmula O30).
CAMPANAS_SIN_AVALUO = {"WM", "Premium"}

# Campañas que NO cobran comisión por apertura (fórmula O31).
CAMPANAS_SIN_COMISION = {"WM", "Premium", "Vinculación", "Hipoteca Green"}

# Hipoteca Green solo aparece en el Excel cuando el producto es el 1.
CAMPANAS_SOLO_PRODUCTO_1 = {"Hipoteca Green"}


# =============================================================================
# BLOQUE 4. ESQUEMA PAGOS OPORTUNOS
# Fuente: hoja Simulador, celdas AR56:BQ59 (un bloque de celdas por campaña).
# Se usa en: F5 (tasa del mes) y F6 (pago mensual).
# -----------------------------------------------------------------------------
# ¿Qué es Pagos Oportunos?
#   Un esquema con mensualidad FIJA. Si el cliente paga puntual, la tasa baja
#   con el tiempo; como el pago no cambia, el crédito se termina antes.
#
# ¿Qué se guarda por cada campaña?
#   tasa             Tasa anual con la que empieza el crédito.
#                    En este esquema NO depende del nivel de tasa del cliente.
#   pago_al_millar   Pago mensual por cada 1,000 pesos de crédito, según el
#                    plazo en años. Es el factor de la anualidad que el banco
#                    ya tiene calculado (F6):   pago = crédito / 1000 x pago_al_millar
#                    Ej.: 1,900,000 / 1000 x 11.86 = 22,534 pesos al mes.
#
# ¿Cómo se lee?
#   PAGOS_OPORTUNOS[campaña]["tasa"]                  ->  tasa anual
#   PAGOS_OPORTUNOS[campaña]["pago_al_millar"][plazo] ->  pago al millar
#   Ej.: PAGOS_OPORTUNOS["Ninguna"]["pago_al_millar"][20]  ->  11.86
#
# ¿Por qué no se calcula el pago con la fórmula de anualidad del curso?
#   Porque el banco redondea su tabla y la fórmula no da exactamente el mismo
#   número. Para que nuestros resultados coincidan con el Excel usamos la tabla;
#   la fórmula del curso explica de dónde viene (ver documento de la Fase 2).
# =============================================================================

# Plazos permitidos en este esquema (en años).
PLAZOS_PAGOS_OPORTUNOS = [7, 10, 15, 20]

PAGOS_OPORTUNOS = {
    "Ninguna":        {"tasa": 0.13,   "pago_al_millar": {7: 18.38, 10: 15.05, 15: 12.79, 20: 11.86}},
    "WM":             {"tasa": 0.095,  "pago_al_millar": {7: 16.42, 10: 13.02, 15: 10.53, 20: 9.42}},
    "Premium":        {"tasa": 0.0975, "pago_al_millar": {7: 16.55, 10: 13.16, 15: 10.70, 20: 9.59}},
    "Vinculación":    {"tasa": 0.1075, "pago_al_millar": {7: 17.10, 10: 13.73, 15: 11.32, 20: 10.27}},
    "Hipoteca Green": {"tasa": 0.10,   "pago_al_millar": {7: 16.70, 10: 13.30, 15: 10.84, 20: 9.75}},
}

# Bajadas de tasa por pagar puntual (F5).
# Fuente: hoja Tabla de amortizacion, columnas AL:AO, filas 27, 39, 51 y 63.
# ¿Cómo se lee?  llave = mes en que baja la tasa ; valor = cuánto baja.
#   En el mes 13 baja 0.20%, en el 25 otro 0.20%, en el 37 otro 0.20% y en el
#   49 baja 0.15%. En total baja 0.75% (por ejemplo, de 13% a 12.25%).
# Se guarda así porque el programa recorrerá los meses y solo tendrá que
# preguntar: "¿este mes está en el diccionario?" -> si sí, resta ese valor.
BAJADAS_PAGOS_OPORTUNOS = {13: 0.0020, 25: 0.0020, 37: 0.0020, 49: 0.0015}


# =============================================================================
# BLOQUE 5. ESQUEMA VALORA
# Fuente: hoja Simulador, celdas AR48:BT54 (un bloque de celdas por campaña).
# Se usa en: F5 (tasa del mes), F6 (pago inicial) y F7 (pago que crece cada año).
# -----------------------------------------------------------------------------
# ¿Qué es Valora?
#   Un esquema en el que la mensualidad empieza más baja y SUBE un porcentaje
#   fijo cada 12 meses (anualidad geométrica). Pensado para quien espera que su
#   sueldo crezca.
#
# ¿Qué se guarda?
#   piso              Tasa mínima a la que puede bajar la tasa por pagar puntual.
#   niveles           Un diccionario con los datos para cada nivel de tasa:
#       tasa              Tasa anual inicial.
#       pago_al_millar    Pago por cada 1,000 pesos, según el plazo en años
#                         (Valora no tiene plazo de 7 años).
#       incremento_anual  g: cuánto sube el pago cada año (0.02 = 2%).
#
# ¿Por qué en "Ninguna" hay 5 niveles y en las demás campañas solo "todos"?
#   - Sin campaña, la tasa depende del riesgo del cliente: a mayor riesgo,
#     mayor tasa (nivel 1 = 12% ... nivel R = 14%).
#   - En las campañas el Excel trae un solo renglón, marcado como válido para
#     los niveles "1, 2, 3, 4, R". Por eso se guarda una sola vez con la llave
#     "todos" en lugar de repetirlo cinco veces (ver la nota al final).
#
# ¿Cómo se lee?
#   VALORA["Ninguna"]["niveles"]["3"]["tasa"]                 ->  0.13
#   VALORA["Ninguna"]["niveles"]["3"]["pago_al_millar"][20]   ->  11.28
#   VALORA["WM"]["niveles"]["todos"]["incremento_anual"]      ->  0.026
# =============================================================================

# Plazos permitidos en este esquema (en años).
PLAZOS_VALORA = [10, 15, 20]

VALORA = {
    "Ninguna": {
        "piso": 0.12,
        "niveles": {
            "1": {"tasa": 0.12,  "pago_al_millar": {10: 13.63, 15: 11.32, 20: 10.36}, "incremento_anual": 0.020},
            "2": {"tasa": 0.125, "pago_al_millar": {10: 14.06, 15: 11.78, 20: 10.84}, "incremento_anual": 0.017},
            "3": {"tasa": 0.13,  "pago_al_millar": {10: 14.50, 15: 12.22, 20: 11.28}, "incremento_anual": 0.014},
            "4": {"tasa": 0.135, "pago_al_millar": {10: 14.90, 15: 12.63, 20: 11.70}, "incremento_anual": 0.011},
            "R": {"tasa": 0.14,  "pago_al_millar": {10: 15.30, 15: 13.04, 20: 12.13}, "incremento_anual": 0.007},
        },
    },
    "WM": {
        "piso": 0.095,
        "niveles": {
            "todos": {"tasa": 0.095, "pago_al_millar": {10: 11.98, 15: 9.50, 20: 8.38}, "incremento_anual": 0.026},
        },
    },
    "Premium": {
        "piso": 0.0975,
        "niveles": {
            "todos": {"tasa": 0.0975, "pago_al_millar": {10: 12.12, 15: 9.64, 20: 8.54}, "incremento_anual": 0.026},
        },
    },
    "Vinculación": {
        "piso": 0.1075,
        "niveles": {
            "todos": {"tasa": 0.1075, "pago_al_millar": {10: 12.75, 15: 10.33, 20: 9.30}, "incremento_anual": 0.024},
        },
    },
    "Hipoteca Green": {
        "piso": 0.10,
        "niveles": {
            "todos": {"tasa": 0.10, "pago_al_millar": {10: 12.24, 15: 9.78, 20: 8.70}, "incremento_anual": 0.026},
        },
    },
}

# Bajada de tasa por pagar puntual en Valora (F5).
# Fuente: hoja Tabla de amortizacion, columna AJ (filas 51, 63, ..., 243).
# La tasa baja 0.25% cada 12 meses empezando en el mes 37 (meses 37, 49, 61...),
# pero nunca queda por debajo del "piso" de la campaña.
# Ej.: nivel 3 sin campaña empieza en 13%; en el mes 37 baja a 12.75%, en el 49
# a 12.50%, en el 61 a 12.25% y en el 73 llega al piso de 12% y ahí se queda.
BAJADA_ANUAL_VALORA = 0.0025
PRIMER_MES_BAJADA_VALORA = 37


# =============================================================================
# BLOQUE 6. GASTOS DE ORIGINACIÓN (lo que el cliente paga al inicio)
# Fuente: hoja Simulador, celdas AE65, AE66 y la tabla de avalúos S91:U112.
# Se usa en: F3 (gastos) y F4 (inversión total del cliente).
# -----------------------------------------------------------------------------
# IMPUESTOS_Y_DERECHOS    Impuesto por comprar la casa y su registro.
#                         3% del valor de la casa. Ej.: 0.03 x 2,000,000 = 60,000.
# HONORARIOS_NOTARIALES   Lo que cobra el notario.
#                         2% del valor de la casa. Ej.: 0.02 x 2,000,000 = 40,000.
# Son los porcentajes que el simulador usa si el usuario no captura otros.
# =============================================================================
IMPUESTOS_Y_DERECHOS = 0.03
HONORARIOS_NOTARIALES = 0.02

# Tabla de AVALÚOS (el perito que revisa cuánto vale la casa).
# ¿Por qué es una lista y no un diccionario?
#   Porque aquí no se busca un valor exacto, sino el RANGO en el que cae la casa.
#   Cada renglón es una pareja (valor "desde", costo del avalúo), en orden.
#   El programa recorrerá la lista y se quedará con el último renglón cuyo
#   "desde" sea menor o igual al valor de la casa. Así funciona también la
#   búsqueda del Excel (BUSCARV aproximado).
#   Ej.: una casa de 2,000,000 cae en el renglón (1_500_001, 5_750) -> 5,750.
# Las parejas ( , ) se llaman tuplas: son como listas pequeñas que no cambian.
AVALUOS = [
    (0.01,       1_900),
    (500_001,    3_300),
    (1_000_001,  4_500),
    (1_500_001,  5_750),
    (2_000_001,  7_000),
    (2_500_001,  8_200),
    (3_000_001, 11_000),
    (4_000_001, 13_900),
    (5_000_001, 15_500),
    (6_000_001, 19_000),
    (7_000_001, 20_000),
    (8_000_001, 20_100),
    (9_000_001, 22_500),
    (10_000_001, 27_000),
    (11_000_001, 28_000),
    (12_000_001, 30_000),
    (13_000_001, 32_000),
    (14_000_001, 34_000),
    (15_000_001, 37_000),
    (20_000_001, 39_000),
    (25_000_001, 40_000),
    (30_000_001, 40_000),
]

# Para casas de más de 30,000,000 el Excel muestra "Favor de cotizar"
# (el costo de 40,000 es solo aproximado). La página mostrará el mismo aviso.
AVALUO_LIMITE_COTIZAR = 30_000_000


# =============================================================================
# BLOQUE 7. SEGUROS, PENALIZACIÓN E IVA
# Fuente: hoja Simulador, celdas AA43 y AA44; hoja Tabla de amortizacion, col. L.
# Se usa en: F13 (penalización) y F15 (seguros y pago total del mes).
# -----------------------------------------------------------------------------
# SEGURO_VIDA    Seguro de vida y desempleo: si el cliente muere o pierde su
#                empleo, cubre la deuda. Se cobra cada mes sobre lo que se
#                debe al inicio del mes, por eso baja conforme baja la deuda.
#                Ej.: 0.0006 x 1,900,000 = 1,140 pesos en el primer mes.
#
# SEGURO_DANOS   Seguro de daños: protege la casa (incendio, sismo). Se cobra
#                cada mes sobre el valor de la casa, así que es fijo.
#                Ej.: 0.00031 x 2,000,000 = 620 pesos al mes.
#
# IVA            Impuesto de 16%. Se suma a la penalización por prepago:
#                penalización = prepago x 0.03 x (1 + 0.16).
# =============================================================================
SEGURO_VIDA = 0.0006
SEGURO_DANOS = 0.00031
IVA = 0.16


# =============================================================================
# BLOQUE 8. CONVENCIONES DE CÁLCULO DEL BANCO
# Fuente: hoja Tabla de amortizacion, columnas F, BD, BE y BJ.
# Se usa en: F5 y F9 (interés del mes) y para generar las fechas de pago.
# -----------------------------------------------------------------------------
# DIAS_ANIO_BANCARIO  El banco calcula el interés de cada mes así:
#                     saldo x tasa anual / 360 x días reales del mes.
#                     Por eso un mes de 31 días cobra más interés que uno de 28.
#
# MESES_POR_ANIO      Se usa para saber cuándo empieza un año nuevo del crédito
#                     (meses 13, 25, 37...), que es cuando sube el pago en Valora
#                     y cuando baja la tasa.
#
# Fecha de corte: el día del mes en que vence cada pago.
#   Si el crédito se calcula entre el día 3 y el 16 del mes -> vence el día 3.
#   En cualquier otro día                                    -> vence el día 17.
#   Importa porque decide cuántos días tiene el primer periodo de interés.
# =============================================================================
DIAS_ANIO_BANCARIO = 360
MESES_POR_ANIO = 12
DIA_CORTE_TEMPRANO = 3
DIA_CORTE_TARDIO = 17


# =============================================================================
# BLOQUE 9. INFONAVIT: APORTACIÓN PATRONAL (productos 2 y 5, "con Apoyo")
# Fuente: hoja Simulador, celdas AF70:AF77; hoja Tabla de amortizacion, col. N.
# Se usa en: la tabla de amortización (una columna extra que baja el saldo).
# -----------------------------------------------------------------------------
# ¿Qué es?
#   Todo patrón en México aporta al Infonavit el 5% del sueldo de su trabajador.
#   En los productos "con Apoyo Infonavit", esas aportaciones se mandan al banco
#   cada BIMESTRE (cada 2 meses) y se usan como abono a la deuda. Así el crédito
#   se paga más rápido sin que el cliente ponga más dinero.
#
# ¿Cómo se calcula en el simulador?  (lo programaremos en amortizacion.py)
#   aportación bimestral = sueldo mensual x 5% x 2 meses
#   pero con un TOPE (no puede pasar de):
#   tope = UMA diaria x 25 veces x 5% x 60.8 días = 117.31 x 25 x 0.05 x 60.8 = 8,915.56
#   - Se abona en los meses 2, 4, 6, 8... (cada 2 meses).
#   - Sube 3% cada año (meses 14, 26, 38...), como los aumentos de sueldo.
#
# Ej.: sueldo de 30,000 -> 30,000 x 0.05 x 2 = 3,000 pesos cada bimestre,
#      que se restan del saldo además de la amortización normal.
#
# ¿Qué es la UMA?  Unidad de Medida y Actualización: un valor oficial en pesos
#   que se usa para calcular límites de leyes y prestaciones. El simulador usa
#   la UMA diaria de enero de 2026: 117.31 pesos.
# =============================================================================
UMA_DIARIA = 117.31                    # celda AF70
TOPE_VECES_UMA = 25                    # celda AF71: tope de 25 veces la UMA
APORTACION_PATRONAL_VIVIENDA = 0.05    # celda AE74: 5% del sueldo
DIAS_BIMESTRE = 60.8                   # celda AF75: días promedio de un bimestre
INCREMENTO_ANUAL_APORTACION = 0.03     # celda AF77: sube 3% cada año

MES_PRIMERA_APORTACION = 2             # la primera llega en el mes 2 (fila 16)
MESES_ENTRE_APORTACIONES = 2           # después, cada 2 meses
MES_PRIMER_INCREMENTO_APORTACION = 14  # primer aumento en el mes 14 (fila 28), luego cada 12


# =============================================================================
# BLOQUE 10. INFONAVIT: COFINANCIAMIENTO (productos 3 y 7)
# Fuente: hoja Simulador, celdas AE24:AE29, AE63, AE64 y O26.
# Se usa en: F2 (crédito del banco) y F3 (gastos).
# -----------------------------------------------------------------------------
# ¿Qué es?
#   El banco y el Infonavit prestan JUNTOS para comprar la misma casa. Además, el
#   cliente puede usar el ahorro de su Subcuenta de Vivienda (SSV): el dinero que
#   su patrón ha depositado en el Infonavit a lo largo de los años.
#   La casa se paga con:  crédito del banco + crédito Infonavit + SSV + enganche.
#
# Diferencia entre los dos productos:
#   3 Cofinavit:  el usuario CAPTURA cuánto le presta el Infonavit (lo sabe por
#                 su precalificación). El banco presta hasta el 95% de la casa,
#                 y si banco + Infonavit + SSV pasan del 100%, el banco presta
#                 menos para no rebasar el valor de la casa.
#                 Ej.: casa de 2,000,000, Infonavit 300,000 (15%) y SSV 100,000 (5%):
#                      95% + 15% + 5% = 115% -> se pasa 15% -> banco presta 80%
#                      = 1,600,000.  Total: 1,600,000 + 300,000 + 100,000 = 2,000,000.
#   7 CI+CB:      el simulador CALCULA el crédito Infonavit: es el 30% del
#                 crédito total, con un máximo de 885,445.60. El crédito total
#                 es el 90% de la casa (o la casa menos la SSV, si es menor).
#                 El banco presta el resto (70%).
#                 Ej.: casa de 2,000,000 -> total 1,800,000 -> Infonavit 540,000
#                      y banco 1,260,000.
#
# Gasto extra: el Infonavit cobra 3% de SU crédito como gastos de originación.
#   Ej.: crédito Infonavit de 540,000 -> 540,000 x 0.03 = 16,200.
# =============================================================================
COMISION_INFONAVIT = 0.03                  # celda AE64 (se aplica en O26)
MONTO_MAXIMO_CREDITO_INFONAVIT = 885_445.60  # celdas AE29 y AE63
PORCENTAJE_INFONAVIT_CI_CB = 0.30          # celda AE27: Infonavit presta el 30% (producto 7)


# =============================================================================
# BLOQUE 11. PASIVO: GASTOS NOTARIALES FINANCIADOS (productos 4 y 5)
# Fuente: hoja Simulador, celdas I27, I28 y O25.
# Se usa en: F2 (crédito total) y F3 (gastos).
# -----------------------------------------------------------------------------
# ¿Qué es?
#   En los productos de Pasivo el cliente puede pedir que los gastos de la
#   escrituración se sumen al crédito, en lugar de pagarlos de su bolsillo.
#   El banco agrega hasta un 6% del crédito solicitado, sin pasar del crédito
#   máximo. Si el cliente ya pide el crédito máximo, no queda espacio y se
#   financia 0.
#   Ej.: crédito máximo 1,900,000 y crédito solicitado 1,500,000:
#        6% de 1,500,000 = 90,000  ->  crédito total 1,590,000.
# =============================================================================
PORCENTAJE_GASTOS_FINANCIADOS = 0.06       # celda I28


# =============================================================================
# NOTA SOBRE VALORA CON CAMPAÑA
# En el Excel, si se elige Valora con una campaña distinta de "Ninguna" y un
# nivel de tasa diferente de 1, la tasa queda vacía y el simulador no da
# resultado. Como el propio Excel marca ese renglón como válido para los niveles
# "1, 2, 3, 4, R", aquí se aplica a todos los niveles (llave "todos").
# Para el nivel 1 el resultado es idéntico al del Excel.
# =============================================================================


# -----------------------------------------------------------------------------
# REVISIÓN RÁPIDA DEL CATÁLOGO
# ¿Qué hace este bloque?
#   Imprime un resumen para comprobar que los datos se cargaron bien.
# ¿Cuándo se ejecuta?
#   Solo cuando se corre este archivo directamente con:  python parametros.py
#   Python le da a la variable especial __name__ el valor "__main__" únicamente
#   en ese caso. Cuando otro archivo hace "import parametros", esta condición
#   es falsa y el bloque se ignora (así no ensucia la página con estos textos).
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    print("Productos cargados:")
    # .items() recorre el diccionario dando cada pareja (llave, valor):
    # aquí "numero" es 1, 4 o 6 y "datos" es el diccionario de ese producto.
    for numero, datos in PRODUCTOS.items():
        # Las f-strings (f"...") insertan variables dentro del texto.
        # ":.2%" muestra un decimal como porcentaje con 2 decimales (0.0125 -> 1.25%).
        print(f"  {numero}. {datos['nombre']}  (comisión {datos['comision_apertura']:.2%}, PTI {datos['pti']:.0%})")

    print("\nPagos Oportunos, sin campaña:")
    po = PAGOS_OPORTUNOS["Ninguna"]
    for plazo in PLAZOS_PAGOS_OPORTUNOS:
        print(f"  {plazo} años: tasa {po['tasa']:.2%}, pago al millar {po['pago_al_millar'][plazo]}")

    print("\nValora, sin campaña, 20 años:")
    for nivel, datos in VALORA["Ninguna"]["niveles"].items():
        print(f"  Nivel {nivel}: tasa {datos['tasa']:.2%}, pago al millar {datos['pago_al_millar'][20]}, "
              f"incremento anual {datos['incremento_anual']:.1%}")

    print("\nInfonavit:")
    tope = UMA_DIARIA * TOPE_VECES_UMA * APORTACION_PATRONAL_VIVIENDA * DIAS_BIMESTRE
    print(f"  Tope de aportación bimestral: {tope:,.2f}")
    print(f"  Monto máximo de crédito Infonavit: {MONTO_MAXIMO_CREDITO_INFONAVIT:,.2f}")

    # len() cuenta cuántos elementos tiene una lista.
    print(f"\nRangos de avalúo cargados: {len(AVALUOS)}")
    print("Revisión terminada: el catálogo se cargó sin errores.")
