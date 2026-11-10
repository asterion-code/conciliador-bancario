# Conciliador Bancario — Diagnóstico y Plan v2.4

**Objetivo:** extender el conciliador (hoy solo cubre `Dep. bancario`) para incorporar `T. Credito` y `Nota de credito`, manteniendo la pestaña de carga del Registro de Cobranzas intacta en el Excel de salida y sumando las hojas ya construidas en v2.3.

---

## 0. Estado real vs. plan original (actualizado)

Este documento se quedó desactualizado a medida que la implementación avanzó — varias fases se completaron con más profundidad de la que el plan original anticipaba, a partir de lo que fueron revelando los datos reales. Este bloque reemplaza el resumen de progreso; §1-§3 más abajo quedan como registro histórico del diagnóstico original.

| Fase | Alcance original del plan | Estado real |
|---|---|---|
| 0 | Base v2.4 | ✅ Hecho |
| 1 | Hoja "Registro Original" intacta en el export | ✅ Hecho |
| 2 | Notas de crédito | ⏸️ **Sigue en pausa** — sin empezar, a pedido explícito del usuario |
| 3B | Depósitos de tarjeta ↔ Banco, multi-operador | ✅ Hecho — Niubiz + Amex + Diners, con acumulación multi-día |
| 3A | Cobranzas T.Credito ↔ Ventas, "match aproximado con niveles de confianza" | ✅ Hecho, pero mucho más profundo de lo previsto (ver abajo) |

**Lo que el plan no anticipaba y se construyó a partir de los datos reales:**

- **`Ventas` trae las 4 marcas de tarjeta juntas** (columna `Marca tarjeta`: VI/MC → Niubiz, AM → Amex, DN → Diners) — el plan asumía que había que combinar `Ventas` + `Ventas amex` + archivos Diners por separado. Se simplificó a una sola fuente de Ventas.
- **`ID operación` (Ventas) = `ID Operación` (Depositos)** — es la misma clave de transacción real en ambas hojas (confirmado con datos reales). Permite un join exacto para trazar el depósito bancario real de una venta puntual, en vez de aproximar por importe+fecha dentro de un grupo agregado.
- **Matching de "split" en ambas direcciones**: Cobranza partida en varias filas para 1 venta (mismo NRO/DOC real, o mismo día+marca+importe), y Venta partida en varias filas para 1 Cobranza. Confirmado con casos reales (REALIZACIONES INMOBILIARIAS, ROSARIO DONAYRE, Fila 1913+1914).
- **Observación de posible error de tipeo en DOCUMENTO**: cuando una Cobranza sin venta candidata en su propia marca calza exacto (importe+fecha) con una venta de OTRA marca, se señala como sospecha de dato mal tipeado en vez de dejarla como "sin explicación". Confirmado en ~10 filas reales (debían decir Amex o Diners, no Visa).
- **Ventana de match ampliada a 10 días** (Cobranzas puede registrar una venta hasta 8 días después de la fecha real) — pero los matches de "split" se mantienen en 5 días, porque combinar 2-3 importes para que sumen exacto es más propenso a coincidencia; una ventana de 10 ahí produjo un falso positivo real (ver commits).
- **Reporte consolidado**: en vez de 3 pestañas separadas (Ventas, Depósitos, OBS_TC) que había que cruzar a mano, una sola hoja "TCredito Consolidado" con las 305 filas de Cobranzas T.Credito + su relación con Ventas + su relación con Bancos, formato definido por el usuario.
- **Bug de punto flotante corregido**: `Math.abs(243.08-243.07)` da `0.626018159083016613` en JS, no `0.01` — un `> 0.01` plano rechazaba matches reales de 1 centavo.

**Archivo de datos cambió a mitad de camino**: el usuario consolidó `Cruce tarjetas.xlsx` (Ventas + Depositos + Ventas amex) + los 4 archivos Diners sueltos en un solo archivo nuevo, **`Ventas y depositos.xlsx`**, con 4 hojas: `Ventas`, `Depositos` (ahora solo Niubiz — Amex salió a su propia hoja), `Depositos amex`, `Depositos Diners`. Estructura de columnas distinta a la anterior en varios puntos clave — ver §4.

---

## 4. Migración a `Ventas y depositos.xlsx` (en curso)

Reemplaza como única fuente de datos de operadores: `Cruce tarjetas.xlsx` + los 4 archivos Diners sueltos quedan obsoletos.

| Hoja | Filas | Cambios clave vs. la fuente anterior |
|---|---|---|
| `Ventas` | 2,168 | Encabezado ahora en la fila 1 (antes fila 3). Sin `FECHA CORTA` (queda `Fecha y hora de Transacción`). Mismas columnas relevantes (`Moneda`, `Monto`, `Marca tarjeta`, `ID operación`) — la detección por palabra clave ya usada sigue funcionando sin cambios de código. |
| `Depositos` | 2,296 | Encabezado en fila 1. **Ya NO trae Amex** (antes tenía un bloque final sin `Cód. Comercio` que se detectaba como Amex — ahora ese bloque se fue a su propia hoja). **Ya NO trae `Cód. Comercio`** — el canal POS vs Pago Link ahora se lee directo de la columna `Producto`: `"Niubiz POS Basico"` vs `"Niubiz Pago Link"` (confirmado: 1,678 POS / 618 Pago Link) — más simple y explícito que el mapeo de códigos de comercio anterior. |
| `Depositos amex` | 71 | Estructura totalmente distinta: `CODIGO DEL ESTABLECIMIENTO, FECHA DEL DEPOSITO, MONTO TOTAL, ..., MONTO DEPOSITADO, BANCO, ESTADO`. Sin `ID Operación` — es un resumen de liquidación, no tiene detalle a nivel de transacción individual (a diferencia de Niubiz). El join exacto por ID operación no aplica aquí; toca seguir aproximando por importe+fecha para Amex. |
| `Depositos Diners` | 16 | Combina venta + pago en una sola fila (`Importe del consumo` bruto + `Importe neto de pago` neto + `Fecha de pago`) — más completo que las 4 hojas Diners sueltas de antes (que había que cruzar por `Orden de pago`). |

**Pendiente de implementar** (siguiente paso inmediato):
1. Adaptar `buildNiubizAmexDepositos()` para leer `Depositos` de este archivo (sin bloque Amex mezclado, canal por `Producto` en vez de `Cód. Comercio`).
2. Nueva función para `Depositos amex` (estructura propia, sin `ID Operación`).
3. Nueva función para `Depositos Diners` (reemplaza la lógica de `pagos` sueltos — ahora ya viene con bruto+neto+fecha de pago en una fila).
4. Actualizar la UI de carga: una sola dropzone para `Ventas y depositos.xlsx` en vez de "Cruce Tarjetas" + "Diners".
5. Verificar con datos reales que el join por `ID operación` (Niubiz) sigue funcionando igual o mejor.

---

## 1. Diagnóstico (registro histórico — visto contra los archivos originales, antes de la consolidación)

### 1.1 Registro de Cobranzas — las 3 formas

| Forma | Filas (jul-2026) | Identificador único | Observación |
|---|---|---|---|
| `Dep. bancario` | 987 | Nro. Operación (NRO/DOC) | Resuelto en v2.3 |
| `T. Credito` | 305 | **Ninguno** — NRO/DOC = `-9999` en 298 de 305 filas | Requiere matching aproximado |
| `Nota de credito` | 236 | SERIE + NRO/DOC (se repite: 776×10, 389×7...) | No hay banco asociado (DOCUMENTO = NULL en el 100%) |

**T. Credito** — DOCUMENTO trae la marca de tarjeta, no un banco: Visa (280), Mastercard (11), American Express (9), Diners (4). Moneda: US$ 229 / S/ 76.

**Nota de credito** — anula o reduce una factura; no ingresa dinero al banco. Su "conciliación" es una validación estructural interna (que Importe Parcial sume el Monto Total, reutilizando `checkFacturasCompletas` que ya soporta Tipo de Cambio), no un cruce bancario.

### 1.2 Archivo `Cruce tarjetas.xlsx` (obsoleto — reemplazado por `Ventas y depositos.xlsx`, ver §4)

Tres operadores de tarjeta, cada uno con su propio patrón de liquidación. Los tres depositan en las **mismas 4 cuentas bancarias** (BBVA PEN, BBVA USD, BCP PEN, BCP USD).

#### Niubiz (Visa/Mastercard) — hojas `Ventas` / `Depositos`

- `Ventas`: 2,173 filas, encabezado en fila 4. Columnas clave: `FECHA CORTA`, `Moneda`, `Monto` (bruto), `Estado`, `N° lote`, `Cod.Comercio`.
- `Depositos`: 2,372 filas, encabezado en fila 3. Trae la cadena completa: `Importe de Operación` (bruto) → `Comisión Total` + `IGV` → `Suma Depositada` (neto) → `Banco Pagador` + `Fecha de depósito`.
- **Fecha de depósito viene como texto `DD-MM-YYYY`**, no como fecha nativa — hay que parsearla.
- 8 códigos de comercio, todos del mismo RUC (EMPRESA DEMO). Mapeo confirmado:

  | Cód. Comercio | Banco | Moneda | Canal | Glosa en extracto |
  |---|---|---|---|---|
  | 209281590 | BBVA | PEN | POS-POS | `ABONO VISANET 5656321223` |
  | 952538885 | BBVA | USD | POS-POS | `ABONO VISANET 4079244026` |
  | 315472803 | BCP | PEN | POS-POS | `DE COMPANIA PERUANA DE` |
  | 966661760 | BCP | USD | POS-POS | `DE COMPANIA PERUANA DE` |
  | 952808414 | BBVA | PEN | **Pago Link** | `R##########COMPAÑIA PERUANA DE MEDIOS D` |
  | 811777741 | BBVA | USD | **Pago Link** | `R##########COMPAÑIA PERUANA DE MEDIOS D` |

  **✅ Resuelto — no faltaba ningún extracto.** BBVA usa una glosa distinta según el canal de venta: `ABONO VISANET` para POS presencial, y `R20341198217COMPAÑIA PERUANA DE MEDIOS D` para ventas por **Pago Link** (comercio electrónico). Mismo operador (Niubiz), misma cuenta bancaria, glosa diferente. Verificado: BBVA USD → 18 movimientos de julio con esta glosa, suma US$ 64,481.73; BBVA PEN → 7 movimientos, S/ 31,742.70 (ambos en el rango esperado del código de comercio correspondiente).

- **Algoritmo de cruce validado empíricamente** (probé 6 estrategias contra el banco real):
  - Agrupar por **fecha de depósito** (no por lote — 38 de 71 grupos-día combinan más de un lote).
  - Usar **`Suma Depositada` (neto)** — el bruto da 0% de aciertos.
  - Ventana de **±2 días**, permitiendo acumulación entre días (el banco suele abonar 1 día antes, y a veces junta 2 días de Niubiz en un solo abono).
  - Resultado con glosa `ABONO VISANET` (canal POS-POS): ~93-100% de aciertos.
  - Resultado con glosa `COMPAÑIA PERUANA DE MEDIOS` (canal Pago Link, BBVA USD 811777741): 15/21 exacto directo + 6/21 explicados por acumulación de 2 días — sin residuo sin explicar.
  - ⚠️ La brecha inicial que parecía "desfase de corte" (Depositos 2.4x el extracto) en realidad era estar sumando códigos de comercio de cuentas/canales distintos sin filtrar, y buscar solo una de las dos glosas posibles por cuenta.
  - **Conclusión de diseño:** cada cuenta bancaria puede tener **más de un patrón de glosa** para el mismo operador (según canal de venta o sistema de origen) — la tabla de configuración debe soportar una lista de patrones por código de comercio, no uno solo. Mismo hallazgo aplica a BCP (`DE COMPANIA PERUANA DE` y `##########  COMPAN` conviven para Amex).

#### American Express — hojas `Ventas amex` / `Dep amex` (+ cola de `Depositos`)

- Mismo patrón de agregación por fecha que Niubiz, pero:
  - Fecha en formato `DD/MM/YYYY` (con barras, no guiones).
  - Sin `Cód. Comercio` ni `Marca de Tarjeta` — bloque separado dentro de la hoja `Depositos` (filas ~2340-2372) y en `Dep amex`.
  - `Banco Pagador` = `BBVA` / `BCP` (texto plano, no `BANCO CONTINENTAL` / `BANCO DE CREDITO` como Niubiz).
  - Glosa en BCP: dos formatos — `DE COMPANIA PERUANA DE` **y** `NNNNNNNNNNNNNNN COMPAN` (el número codifica fecha: `DDMM` + secuencia).
  - Comisión ~4.9% (vs ~2.9% Niubiz), consistente con tarifa Amex.

#### Diners — 4 archivos nuevos (`Diners ventas_mar_jul_*.xlsx`, `diners pagos_mar_jul_*.xlsx`)

- **El más simple de los tres.** `Orden de pago` es clave exacta que conecta `ventas` ↔ `pagos` 1:1.
- `pagos` = **una fila por pago individual**, sin agregación: `Banco`, `N° Cuenta bancaria`, `Fecha de pago efectivo`, `Importe total de abono` (neto).
- Las cuentas bancarias (`2860913909960308`, `3462819482199351`, `1752917034236`, `8114710497465`) **coinciden exactamente** con las 4 cuentas ya mapeadas para Niubiz.
- Glosa en banco: `DINERS CLUB PERU...` — inconfundible.
- **Verificado:** 3 de 4 pagos de muestra calzan exacto (importe + fecha + glosa) directo en el extracto, sin necesidad de agrupar.

### 1.3 Resumen comparativo — Fase 3B (Depósitos ↔ Banco)

| Operador | Agregación | Ventana | Confianza esperada |
|---|---|---|---|
| Niubiz | Suma por fecha de depósito | ±2 días | Alta (93-100% validado) |
| Amex | Suma por fecha de depósito | ±2 días | Por confirmar (mismo patrón) |
| Diners | Sin agregar — 1:1 | 0-1 día | Alta (75% en muestra, sin ambigüedad) |

### 1.4 Fase 3A (Cobranzas T.Credito ↔ Ventas) — la más incierta

- Sin identificador común explotable (NRO/DOC = `-9999`). **Actualización:** sí existe — `ID operación`, ver §0.
- Conteo de filas julio: Ventas 307 vs Cobranzas T.Credito 305 — muy cercano, buena señal.
- Montos: Ventas (bruto, jul) PEN 91,491.54 / USD 174,137.03 vs Cobranzas PEN 81,731.46 / USD 184,289.72 — diferencia de ~10K en direcciones opuestas por moneda. Hipótesis más probable: mismo problema de Tipo de Cambio/moneda mal clasificada que ya resolvimos para depósitos. **Actualización: investigado a fondo — no era eso.** La diferencia real se debía a que `Ventas` traía las 4 marcas mezcladas y se estaban perdiendo Amex/Diners al asumir que toda la hoja era Niubiz (ver §0). El patrón de Tipo de Cambio > 3 solo explica casos aislados, no la anomalía agregada.
- Match debe ser aproximado (importe + fecha + moneda, con ventana de tolerancia) y clasificado por confianza, no binario. **Hecho**, con 3 niveles (Alta/Media/Baja) — ver §0.

### 1.5 Deuda técnica

`executeCoreLogic()` asume un único flujo (filtra con `isDepositoBancario()` y descarta el resto). Hay que generalizar el filtro de forma de cobranza antes de sumar los nuevos flujos, o el código se vuelve inmanejable. **Sigue pendiente** — Fase 3A/3B se implementaron como flujos independientes (no integrados a `executeCoreLogic`) precisamente para evitar este riesgo, a costa de que el % de conciliados general necesite un paso de fusión aparte (`incorporateTarjetasIntoMetrics`, ya implementado).

---

## 2. Plan de trabajo (histórico)

| Fase | Alcance | Depende de | Riesgo | Estado |
|---|---|---|---|---|
| **0** | Base v2.4: copiar desde v2.3, generalizar filtro de forma de cobranza | — | Bajo | ✅ Hecho |
| **1** | Hoja "Registro Original" intacta en el Excel exportado | — | Bajo | ✅ Hecho |
| **2** | Notas de crédito: pestaña + validación interna (Importe Parcial vs Monto Total, agrupado por SERIE+NRO/DOC) | — | Medio | ⏸️ En pausa (a pedido del usuario) |
| **3B** | Depósitos de tarjeta ↔ Banco, multi-operador (Niubiz + Amex + Diners), con soporte de múltiples glosas por cuenta | Confirmar estructura Amex | Medio | ✅ Hecho |
| **3A** | Cobranzas T.Credito ↔ Ventas (bruto), matching aproximado con niveles de confianza | — | Alto | ✅ Hecho (más profundo de lo previsto) |
| **4** | Migración a `Ventas y depositos.xlsx` como fuente única | 3A/3B ya construidos | Medio | 🔧 **En curso** — ver §4 |

---

## 3. Abierto / pendiente de tu confirmación

1. ~~Extracto bancario faltante para los códigos 952808414/523~~ — **Resuelto.**
2. ~~Alcance de operadores en 3B~~ — **Resuelto: los tres (Niubiz + Amex + Diners) en esta versión.**
3. ~~Anomalía de ~10K PEN/USD entre Ventas y Cobranzas T.Credito~~ — **Resuelto, no era Tipo de Cambio.** Ver §1.4.
4. *(En pausa junto con la Fase 2)* Validaciones deseadas para Notas de Crédito: ¿solo la suma Importe Parcial vs Monto Total, o también detectar NC duplicadas y verificar que referencien una factura existente en el mismo archivo?
5. **Nuevo:** confirmar el plan de migración de §4 antes de tocar código — en particular, cómo tratar `Depositos amex` (sin `ID Operación`, sin detalle a nivel de transacción individual) para el cruce de Fase 3B.
