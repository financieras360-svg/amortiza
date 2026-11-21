# amortiza · Calculadora hipotecaria

Réplica académica del simulador «Financiamiento Residencial» v2026.06 (Scotiabank).
Matemáticas Financieras I · FES Acatlán (UNAM) · Prof. Arturo Vera Moreno · semestre 2027-I.

El núcleo replica el simulador al centavo; las funcionalidades adicionales son capas de
análisis que usan el mismo motor, sin modificar los resultados validados.

## Archivos

| Archivo | Qué hace |
|---|---|
| `parametros.py` | Datos del banco (productos, tasas, pago al millar, aforos, avalúos, seguros, Infonavit) |
| `credito.py` | Crédito máximo, gastos, enganche, ingreso mínimo y avisos (hoja «Simulador») |
| `amortizacion.py` | Tabla de amortización mes a mes (hoja «Tabla de amortizacion») |
| `indicadores.py` | CAT (TIR por bisección), plazo real y totales |
| `extras.py` | Mensualidad deseada, meta de liquidación, pagar menos, línea del tiempo, «¿por qué?» |
| `estilos.py` | Diseño de la página (colores, tarjetas, gráficas SVG) |
| `app.py` | La página en Streamlit (solo muestra; no calcula) |
| `validacion.py` | Comparación contra el Excel: 12 escenarios de tabla y 18 de crédito |

## Cómo correrla

```
pip install -r requirements.txt
python validacion.py
python -m streamlit run app.py
```
