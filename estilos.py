"""
estilos.py
==========
ASPECTO VISUAL DE LA CALCULADORA (no calcula nada)

¿Qué hace este archivo?
    Guarda el diseño "amortiza" que aprobó el equipo:
        - CSS: colores (paleta Scotiabank suave), tarjetas, avisos, pestañas
        - funciones para dar formato a números ($1,900,000.00 · 13.00 %)
        - funciones que arman pedazos de HTML (tarjetas, filas, avisos)
        - gráficas en SVG (dibujos hechos con texto, sin librerías externas)

    app.py solo llama estas funciones. Así la lógica (cálculos) y la
    apariencia (este archivo) quedan separadas.
"""

from html import escape          # convierte < > & en texto seguro para HTML


# =============================================================================
# 1. FORMATO DE NÚMEROS
# =============================================================================

def money(x, decimales=2):
    """1900000 -> '$1,900,000.00'.  Los negativos llevan signo menos."""
    signo = "−" if x < -0.004 else ""
    return "{}${:,.{}f}".format(signo, abs(x), decimales)


def num(x, decimales=2):
    return "{:,.{}f}".format(x, decimales)


def pct(x, decimales=2):
    """0.13 -> '13.00 %'"""
    return "{:.{}f} %".format(x * 100, decimales)


def tex_num(x, decimales=2):
    """Número para LaTeX: la coma de miles va entre llaves para que no deje espacio."""
    return "{:,.{}f}".format(x, decimales).replace(",", "{,}")


def anios_meses(meses):
    a, m = divmod(int(meses), 12)
    partes = []
    if a:
        partes.append("{} {}".format(a, "año" if a == 1 else "años"))
    if m or not a:
        partes.append("{} {}".format(m, "mes" if m == 1 else "meses"))
    return " ".join(partes)


MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]


def fecha_corta(f):
    """date(2026, 12, 17) -> '17 dic 2026'"""
    return "{} {} {}".format(f.day, MESES[f.month - 1], f.year)


def fecha_dmy(f):
    return f.strftime("%d/%m/%Y")


# =============================================================================
# 2. CSS (paleta Scotiabank en versión suave, tomada de la página del equipo)
# =============================================================================
CSS = """
<style>
:root{
  --red-600:#EC111A; --red-700:#C70E17; --plum-700:#7A2E39;
  --rose-50:#FFF8F7; --rose-100:#FDEDEB; --rose-200:#F3D2CD;
  --charcoal:#333333; --ink:#1D1D1D; --ink-2:#6A6A6A; --ink-900:#1E1117;
  --line:#D6D6D6; --bg-2:#F6F6F6; --cream:#FAF6F5; --link:#007EAB;
  --err:#A50010; --err-bg:#FDECEE; --warn:#8A5A00; --warn-bg:#FFF4D6; --warn-bar:#F2A900;
  --info:#007EAB; --info-bg:#E6F4F9; --ok:#00664F; --ok-bg:#E3F0EB;
}
/* Barra roja arriba y fondo */
[data-testid="stAppViewContainer"]{border-top:4px solid var(--red-600);}
[data-testid="stHeader"]{background:transparent;}
.block-container{padding-top:1.2rem; max-width:1240px;}

/* Encabezado */
.am-header{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;
  padding:4px 0 14px;border-bottom:1px solid var(--line);margin-bottom:4px}
.am-logo{display:flex;align-items:center;gap:10px}
.am-logo b{font-size:26px;font-weight:800;color:var(--red-600);letter-spacing:-.02em}
.am-logo span{font-size:13px;color:var(--ink-2)}

/* Hero */
.am-hero{background:linear-gradient(180deg,#FFEFED 0%,#FFF8F7 60%,#FFFFFF 100%);
  border-radius:14px;padding:34px 24px 26px;text-align:center;margin:8px 0 18px}
.am-hero h1{margin:0;font-size:38px;font-weight:800;color:var(--ink-900);letter-spacing:-.02em}
.am-hero .pill{display:inline-block;margin:12px 0 8px;padding:5px 14px;border-radius:999px;
  background:#fff;border:1px solid var(--rose-200);color:var(--red-700);font-weight:700;font-size:14px}
.am-hero p{margin:0;color:var(--ink-2);font-size:15px}

/* Títulos de paso */
.am-step{display:flex;align-items:center;gap:10px;font-size:21px;font-weight:800;color:var(--ink-900);margin:18px 0 6px}
.am-step .n{width:28px;height:28px;border-radius:50%;background:var(--red-600);color:#fff;
  display:inline-flex;align-items:center;justify-content:center;font-size:14px;font-weight:800}

/* Insignia del producto */
.am-badge{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;margin:8px 0 4px}
.am-badge .b{padding:6px 14px;border-radius:999px;background:#fff;border:1px solid var(--line);font-size:14px}
.am-badge .b b{color:var(--red-700)}
.am-badge small{color:var(--ink-2);font-size:12.5px}

/* Tarjeta de resultado (panel derecho) */
.am-result{background:var(--rose-50);border:1px solid var(--rose-200);border-top:4px solid var(--red-600);
  border-radius:12px;padding:22px 22px 18px}
.am-result .eyebrow{font-size:13px;font-weight:700;color:var(--plum-700)}
.am-result .lbl{font-size:13px;color:var(--ink-2);margin-top:12px}
.am-result .big{font-size:46px;font-weight:800;color:var(--red-700);line-height:1.05;font-variant-numeric:tabular-nums}
.am-result .sentence{font-size:14.5px;margin:12px 0 0;color:var(--ink)}
.am-kpis{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;margin-top:16px}
.am-kpis div{background:#fff;border:1px solid var(--rose-200);border-radius:10px;padding:10px 12px}
.am-kpis b{display:block;font-size:19px;color:var(--ink-900);font-variant-numeric:tabular-nums}
.am-kpis span{font-size:12px;color:var(--ink-2)}
.am-idle{background:var(--rose-50);border:1px dashed var(--rose-200);border-radius:12px;padding:40px 24px;text-align:center;color:var(--ink-900)}

/* Avisos */
.am-alert{display:flex;gap:10px;align-items:flex-start;border-left:4px solid;border-radius:6px;
  padding:10px 14px;margin:10px 0;font-size:14px}
.am-alert .ic{font-weight:800;width:20px;height:20px;border-radius:50%;display:inline-flex;
  align-items:center;justify-content:center;font-size:12px;color:#fff;flex:none;margin-top:1px}
.am-alert.err{background:var(--err-bg);border-color:var(--err);color:var(--err)} .am-alert.err .ic{background:var(--err)}
.am-alert.warn{background:var(--warn-bg);border-color:var(--warn-bar);color:var(--warn)} .am-alert.warn .ic{background:var(--warn-bar)}
.am-alert.info{background:var(--info-bg);border-color:var(--info);color:#00516F} .am-alert.info .ic{background:var(--info)}
.am-alert.ok{background:var(--ok-bg);border-color:var(--ok);color:var(--ok)} .am-alert.ok .ic{background:var(--ok)}

/* Banda de resultados */
.am-banner{background:var(--rose-50);border-radius:12px;padding:18px;text-align:center;
  font-size:24px;font-weight:800;color:var(--ink-900);margin:26px 0 8px}

/* Tarjetas del resumen */
.am-grid3{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:16px}
.am-grid2{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px}
.am-grid4{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}
@media (max-width:900px){.am-grid3,.am-grid2,.am-grid4,.am-kpis{grid-template-columns:1fr}}
.am-card{background:#fff;border:1px solid var(--line);border-radius:10px;padding:16px 18px}
.am-card h4{margin:0;font-size:17px;color:var(--ink-900)}
.am-card .sub{font-size:12.5px;color:var(--ink-2);margin:2px 0 8px}
.am-row{display:flex;justify-content:space-between;gap:10px;padding:7px 0;border-bottom:1px solid #EEE;font-size:14px}
.am-row span:last-child{font-variant-numeric:tabular-nums;font-weight:600;text-align:right}
.am-row.total{border-top:2px solid var(--charcoal);border-bottom:none;font-weight:800}
.am-row.total span{font-weight:800}
.am-spec{background:var(--bg-2);border-radius:8px;padding:12px;font-size:13px}

/* Indicadores grandes de la tabla */
.am-stat{border-radius:10px;padding:14px 16px;border:1px solid var(--line);background:#fff}
.am-stat .t{font-size:12.5px;color:var(--ink-2);font-weight:700}
.am-stat .v{font-size:28px;font-weight:800;color:var(--ink-900)}
.am-stat.ok{background:var(--ok-bg);border-color:#B8DDD2} .am-stat.ok .t,.am-stat.ok .v{color:var(--ok)}

/* Opciones A, B, C */
.am-opt{border:1px solid var(--line);border-radius:10px;padding:14px;background:#fff}
.am-opt.on{border:2px solid var(--red-600);background:var(--rose-100)}
.am-opt h5{margin:0 0 4px;font-size:15px;color:var(--ink-900)}
.am-opt p{margin:0 0 6px;font-size:13px;color:var(--ink-2)}
.am-opt b{font-size:21px;color:var(--red-700)}

/* Línea del tiempo */
.am-tl{background:#fff;border:1px solid var(--line);border-radius:12px;padding:6px 22px}
.am-tl .ev{display:grid;grid-template-columns:120px 22px minmax(0,1fr) auto;gap:14px;align-items:center;
  padding:14px 0;border-bottom:1px solid #F0F0F0}
.am-tl .ev:last-child{border-bottom:none}
.am-tl .f{font-size:13.5px;font-weight:700;color:var(--charcoal)}
.am-tl .dot{width:14px;height:14px;border-radius:50%;justify-self:center}
.am-tl .tt{font-weight:700;font-size:15.5px;color:var(--ink-900)}
.am-tl .dd{font-size:13.5px;color:var(--ink-2)}
.am-tl .m{font-size:12.5px;color:var(--ink-2)}

/* "¿Por qué?" */
.am-why{background:#fff;border:1px solid var(--rose-200);border-left:4px solid var(--red-600);border-radius:8px;padding:12px 16px;margin:8px 0}
.am-why b{color:var(--plum-700)}
.am-why code{background:var(--rose-50);color:var(--ink);font-size:13.5px}

/* Pie */
.am-footer{margin-top:36px;padding:16px 0;border-top:1px solid var(--line);font-size:12.5px;color:var(--ink-2);
  display:flex;justify-content:space-between;flex-wrap:wrap;gap:8px}

/* Widgets de Streamlit: opción elegida en rosa */
[data-testid="stBaseButton-segmented_controlActive"], [data-testid="stBaseButton-pillsActive"]{
  background:var(--rose-100)!important;color:var(--plum-700)!important;border-color:var(--red-600)!important;font-weight:700}
.stTabs [data-baseweb="tab"] p{font-weight:700;font-size:15px}
</style>
"""


# =============================================================================
# 3. PEDAZOS DE HTML
# =============================================================================

def fila(etiqueta, valor, total=False):
    clase = "am-row total" if total else "am-row"
    return '<div class="{}"><span>{}</span><span>{}</span></div>'.format(clase, escape(etiqueta), valor)


def tarjeta(titulo, subtitulo, filas_html):
    return ('<div class="am-card"><h4>{}</h4><div class="sub">{}</div>{}</div>'
            .format(escape(titulo), escape(subtitulo), "".join(filas_html)))


ICONO_AVISO = {"err": "!", "warn": "!", "info": "i", "ok": "✓"}


def aviso(nivel, titulo, texto=""):
    """nivel = 'err', 'warn', 'info' u 'ok'."""
    cuerpo = "<b>{}</b>".format(escape(titulo)) + ("<br>{}".format(escape(texto)) if texto else "")
    return '<div class="am-alert {}" role="alert"><span class="ic">{}</span><div>{}</div></div>'.format(
        nivel, ICONO_AVISO[nivel], cuerpo)


def paso(numero, texto):
    return '<div class="am-step"><span class="n">{}</span>{}</div>'.format(numero, escape(texto))


# =============================================================================
# 4. GRÁFICAS EN SVG (dibujadas aquí mismo, sin librerías)
# =============================================================================

def svg_saldos(series, ancho=760, alto=280):
    """
    Gráfica de líneas del saldo al final de cada año.
    series = lista de (nombre, color, [saldos por año])
    """
    izq, der, arr, aba = 70, 20, 40, 34
    n = max(len(s[2]) for s in series)
    maximo = max(max(s[2]) for s in series) or 1
    w, h = ancho - izq - der, alto - arr - aba

    def x(i):
        return izq + (w * i / max(n - 1, 1))

    def y(v):
        return arr + h - h * v / maximo

    partes = ['<svg viewBox="0 0 {} {}" width="100%" role="img" aria-label="Saldo por año" '
              'style="font-family:inherit">'.format(ancho, alto)]
    for k in range(5):                                       # líneas guía
        v = maximo * k / 4
        partes.append('<line x1="{}" x2="{}" y1="{:.1f}" y2="{:.1f}" stroke="#E6E6E6"/>'.format(izq, ancho - der, y(v), y(v)))
        partes.append('<text x="{}" y="{:.1f}" font-size="11" fill="#6A6A6A" text-anchor="end">${:,.1f} M</text>'.format(
            izq - 8, y(v) + 4, v / 1e6))
    for i in range(0, n, max(1, n // 6)):
        partes.append('<text x="{:.1f}" y="{}" font-size="11" fill="#6A6A6A" text-anchor="middle">Año {}</text>'.format(
            x(i), alto - 10, i + 1))
    for nombre, color, datos in series:
        pts = " ".join("{:.1f},{:.1f}".format(x(i), y(v)) for i, v in enumerate(datos))
        partes.append('<polyline points="{}" fill="none" stroke="{}" stroke-width="3" stroke-linejoin="round"/>'.format(pts, color))
    leyenda_x = izq + 10
    for nombre, color, _ in series:
        partes.append('<rect x="{}" y="6" width="12" height="12" rx="3" fill="{}"/>'.format(leyenda_x, color))
        partes.append('<text x="{}" y="17" font-size="12" fill="#333">{}</text>'.format(leyenda_x + 18, escape(nombre)))
        leyenda_x += 40 + 7 * len(nombre)
    partes.append("</svg>")
    return "".join(partes)


def svg_barra_tiempo(meses_total, plazo_meses, marcas, ancho=760, alto=74):
    """
    Barra horizontal del plazo con marcas.
    marcas = lista de (mes, color, etiqueta)
    """
    izq, der = 10, 10
    w = ancho - izq - der

    def x(m):
        return izq + w * m / plazo_meses

    p = ['<svg viewBox="0 0 {} {}" width="100%" role="img" aria-label="Línea del tiempo">'.format(ancho, alto)]
    p.append('<rect x="{}" y="22" width="{}" height="14" rx="7" fill="#F3D2CD"/>'.format(izq, w))
    p.append('<rect x="{}" y="22" width="{:.1f}" height="14" rx="7" fill="#EC111A"/>'.format(izq, x(meses_total) - izq))
    for mes, color, etiqueta in marcas:
        p.append('<rect x="{:.1f}" y="16" width="4" height="26" rx="2" fill="{}"/>'.format(x(mes) - 2, color))
    p.append('<circle cx="{:.1f}" cy="29" r="11" fill="#00664F"/>'.format(x(meses_total)))
    p.append('<text x="{}" y="62" font-size="12" fill="#6A6A6A">Hoy</text>'.format(izq))
    p.append('<text x="{}" y="62" font-size="12" fill="#6A6A6A" text-anchor="end">Plazo contratado: {} años</text>'.format(
        ancho - der, plazo_meses // 12))
    p.append("</svg>")
    return "".join(p)


def svg_planes(pago_normal, filas, ancho=760):
    """
    Línea del tiempo de "pagar menos por un tiempo": una fila por opción.
    filas = lista de (nombre, lista de (mes_inicio, mes_fin, color, etiqueta), marca_mes o None, texto_marca)
    """
    alto_fila, izq = 58, 150
    meses = max(fin for _, tramos, _, _ in filas for _, fin, _, _ in tramos) + 6
    w = ancho - izq - 16
    alto = alto_fila * len(filas) + 26

    def x(m):
        return izq + w * m / meses

    p = ['<svg viewBox="0 0 {} {}" width="100%" role="img" aria-label="Opciones para compensar">'.format(ancho, alto)]
    for k, (nombre, tramos, marca, texto) in enumerate(filas):
        y0 = 10 + k * alto_fila
        p.append('<text x="0" y="{}" font-size="13" font-weight="700" fill="#1E1117">{}</text>'.format(y0 + 24, escape(nombre)))
        p.append('<rect x="{}" y="{}" width="{}" height="20" rx="4" fill="#E9E4E3"/>'.format(izq, y0 + 10, w))
        for ini, fin, color, etiqueta in tramos:
            p.append('<rect x="{:.1f}" y="{}" width="{:.1f}" height="20" rx="4" fill="{}"/>'.format(
                x(ini - 1), y0 + 10, max(2, x(fin) - x(ini - 1)), color))
            p.append('<text x="{:.1f}" y="{}" font-size="11" fill="#7A2E39">{}</text>'.format(
                x(ini - 1) + 4, y0 + 44, escape(etiqueta)))
        if marca is not None:
            p.append('<line x1="{0:.1f}" x2="{0:.1f}" y1="{1}" y2="{2}" stroke="#EC111A" stroke-width="3"/>'.format(
                x(marca), y0 + 2, y0 + 34))
            p.append('<text x="{:.1f}" y="{}" font-size="11" font-weight="700" fill="#C70E17">{}</text>'.format(
                x(marca) + 5, y0 + 6, escape(texto)))
    for a in range(0, meses + 1, 12):
        p.append('<text x="{:.1f}" y="{}" font-size="10" fill="#6A6A6A" text-anchor="middle">{}</text>'.format(
            x(a), alto - 4, "mes {}".format(a)))
    p.append("</svg>")
    return "".join(p)
