# Financiera Yamaha — Consolidación de cartera (clientes antiguos)

Consolidación de la información de créditos del archivo `FINANCIERA_2025.xlsx`
(control manual de la financiera en Pitalito/Garzón) para migrar los clientes
antiguos al sistema del PROYECTO-YAMAHA. Los clientes antiguos no tienen
documentos digitalizados (están en físico), por eso las columnas `cedula` y
`telefono` quedan vacías para completarse después.

## Archivos

| Archivo | Contenido |
|---|---|
| `origen/FINANCIERA_2025.xlsx` | Archivo original, sin modificar (fuente) |
| `consolidar.py` | Script que genera la consolidación (reproducible y auditable) |
| `verificar.py` | Verificación independiente: relee los archivos generados y valida 27 invariantes |
| `FINANCIERA_2025_CONSOLIDADO.xlsx` | Excel de revisión: RESUMEN, CLIENTES, CREDITOS, CUOTAS, MORA, HALLAZGOS |
| `csv/clientes.csv` | 132 clientes (UTF-8 BOM, separador `;`) |
| `csv/creditos.csv` | 133 créditos consolidados |
| `csv/cuotas.csv` | 1,601 cuotas (plan de pagos completo, pagadas y pendientes) |

## Cómo regenerar

```bash
pip install openpyxl
python3 consolidar.py                 # genera Excel + CSVs (corte 2026-07-15)
python3 verificar.py                  # valida todo; termina con error si algo no cuadra
python3 consolidar.py --corte 2026-08-01   # otro corte para estados de mora
```

## Fuente de los datos y confiabilidad

- **Fuente autoritativa**: hoja `LETRAS` del archivo original (una fila por cuota:
  vencimiento, nº crédito, ciudad, cliente, valor letra, valor pagado, fecha de pago, nº cuota).
- **Desembolsos** (capital prestado, fecha y monto por crédito): extraídos de la sección
  "PAGOS EFECTUADOS" de las 31 hojas de caja diaria.
- **Prueba de cuadre**: la suma de todos los pagos de letras registrados en las 31 hojas
  de caja diaria es **exactamente igual** al total pagado según `LETRAS`:
  **$662,178,582 — diferencia $0**. `verificar.py` recalcula este cuadre desde cero.
- Secuencias de cuotas completas (1..plazo) en los 133 créditos, sin duplicados,
  valor de letra constante por crédito, números de crédito sin reuso entre clientes.

## Cifras (corte 2026-07-15)

- 132 clientes, 133 créditos, 1,601 cuotas (889 pagadas, 4 con abono parcial,
  660 pendientes, 48 vencidas sin pagar)
- Total desembolsado: **$883,710,100**
- Total recaudado: **$662,178,582**
- Saldo por cobrar: **$433,414,843** (créditos en mora: $140,852,243; al día: $292,562,600).
  Incluye el faltante de las 4 cuotas con abono parcial ($2,224,696).
- Estados: 47 CANCELADO (20 con cierre anticipado) · 61 AL DIA · 25 EN MORA

## Reglas de interpretación importantes

- **Abonos parciales**: en un crédito activo, una cuota con pago menor a la letra
  es un abono parcial (la caja los rotula "ABONO LETRA"): el faltante sigue siendo
  deuda y cuenta para la mora. Hay 4 casos vivos (créditos #78, #2640, #2731, #2948),
  detallados en HALLAZGOS.
- **Cancelación anticipada**: en un crédito ya liquidado, el tramo final pagado en un
  solo día con descuento o antes del vencimiento es el pago del saldo
  (CIERRE=ANTICIPADO); no deja deuda.

## Correcciones aplicadas (todas documentadas en la hoja HALLAZGOS)

Nada se modificó en silencio. Únicas 4 correcciones, todas de digitación evidente:

1. Fecha de pago `2005-03-28` → `2025-03-28` (LETRAS fila 13, crédito #2234).
2. Fecha de vencimiento `2006-06-30` → `2026-06-30` (LETRAS fila 856, crédito #2674).
3. Nombre `NOHORA ELCY MADRIGAL VALENCUA` → `...VALENCIA` (unificación entre hojas).
4. Nombre `ANYI PAOLA PEFARAN HOYOS` → `...PERAFAN HOYOS` (unificación entre hojas).

Ninguna corrección altera montos. La hoja `HALLAZGOS` también documenta anomalías
que NO se corrigieron (cruces de nombres en caja — incluido un pago de $5,472,500
del crédito #2205 anotado bajo otro cliente —, homónimos que no deben fusionarse,
los abonos parciales vivos y los casos fuera de alcance NORMA/APORTES).

## Nota sobre los IDs de cliente

Los IDs `C001..C132` se derivan de (fecha de primera cuota, nombre) de ESTE corte:
son estables si se reordena la hoja LETRAS, pero pueden correrse si entran clientes
nuevos con fechas anteriores. La clave definitiva de cada cliente debe ser la
**cédula** cuando se digite desde los documentos físicos; hasta entonces, usar
`nombre` o `credito` (únicos en este corte) como referencia cruzada.

## Pendiente (fase B)

Adaptar `csv/*.csv` al esquema real del módulo de financiera de
`mapeba/PROYECTO-YAMAHA` (no accesible desde esta sesión por permisos).
Cuando el repo esté disponible: mapear columnas a sus tablas/campos y generar
el formato de importación exacto (CSV con sus encabezados, seed SQL o API).
