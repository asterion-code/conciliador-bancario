# `portable_2.5.html` — Conciliador General (Cobranzas + Tarjetas)

Herramienta **portable de un solo archivo HTML** para conciliar el
**Registro de Cobranzas** de la empresa contra los extractos bancarios (BCP,
Scotiabank, BBVA/Continental, Interbank, Yape), incluidos los **pagos con
tarjeta** hasta su abono en el banco. Sin instalación, sin servidor y sin
enviar datos a ningún tercero: todo ocurre en el navegador.

Es la **versión activa** del conciliador general. `portable_2.4.html`,
`portable_2.3.html` y `portable_2.2.html` se conservan solo como referencia.

**Contenido**

1. [Objetivo](#objetivo)
2. [Cómo usarlo](#cómo-usarlo)
3. [Archivos de entrada](#archivos-de-entrada)
4. [Qué hace con los depósitos bancarios](#qué-hace-con-los-depósitos-bancarios)
5. [Qué hace con los pagos con tarjeta](#qué-hace-con-los-pagos-con-tarjeta)
6. [Cruce contra asientos ya contabilizados](#cruce-contra-asientos-ya-contabilizados-revisión-quincenal)
7. [El Excel que exporta](#el-excel-que-exporta)
8. [Resultados de referencia](#resultados-de-referencia)
9. [Referencia técnica](#referencia-técnica)
10. [Historial de cambios](#historial-de-cambios)

---

## Objetivo

Cada mes el área de Cobranzas registra en un Excel los cobros recibidos
(depósitos bancarios, pagos con tarjeta, notas de crédito…) y por separado se
descargan los extractos de cada banco. Verificar a mano que cada depósito
aparezca en el banco con su importe, fecha y N° de operación, y que cada pago
con tarjeta haya llegado realmente al banco, es lento y propenso a error con
cientos de filas al mes, dos monedas y cinco bancos.

Desde que la revisión bancaria pasó a ser **quincenal**, la herramienta
también marca qué movimientos ya están contabilizados, para no duplicar
asientos.

## Cómo usarlo

1. Abre `portable_2.5.html` con un navegador moderno.
2. Carga (o arrastra) los archivos:
   - **Registro de Cobranzas** (obligatorio).
   - **Extractos bancarios**: archivos sueltos, un Excel consolidado con una
     hoja por cuenta (p. ej. `Bancos_Q1_20102026.xlsx`) o libros de trabajo
     `B_Bancos_XX.xlsx`.
   - **Ventas y depósitos de tarjetas** (opcional; p. ej.
     `Tarjeta_de_credito.xlsx`), para el cruce de tarjetas.
   - **Asientos contables ya subidos** (opcional), para el estado contable.
3. Revisa el mapeo de columnas que se detecta solo. En Cobranzas: Fecha,
   Cuenta o Forma de Cobranza, Documento (Banco), Nro/Doc, Monto Total
   Cobrado, Importe Parcial, Moneda, Tipo de Cambio. En cada extracto: Fecha,
   N° de operación e Importe.
4. Define las tolerancias (fecha ±2 días, importe ±0.01) y los bancos a
   incluir.
5. Pulsa **Ejecutar Conciliación** (depósitos bancarios).
6. Si cargaste tarjetas, pulsa **Conciliar Depósitos de Tarjeta**. Esto cruza
   los pagos con tarjeta, saca de "Pendiente en Bancos" los abonos de tarjeta
   que el cruce explica y actualiza los porcentajes en pantalla.
7. Revisa las pestañas y **exporta el Excel**. Para tarjetas, empieza por la
   hoja **Resumen Tarjetas → Pendientes a revisar**.

## Archivos de entrada

| Archivo | Contenido |
|---|---|
| `Registro de Cobranzas *.xlsx` | Cobranzas del periodo: depósitos bancarios directos (`Dep. bancario`), pagos con tarjeta (`T. Credito`), etc. |
| Extractos bancarios | Archivos `.xls`/`.xlsx` por banco/moneda (incluye el `.xls` de BBVA que es HTML), un Excel consolidado con una hoja por cuenta, o libros de trabajo `B_Bancos_XX.xlsx`. |
| Ventas y depósitos de tarjetas (opcional) | Hojas `Ventas`, `Ventas amex`, `Depositos`, `Depositos amex`, `Depositos Diners` (Niubiz, Amex, Diners). |
| Asientos contables (opcional) | Export del sistema contable: `Numero` (N° de operación), `Fecha`, `Debe`/`Haber`. |

Ejemplo completo: carpeta `Conciliacion octubre q1/`.

## Qué hace con los depósitos bancarios

1. **Filtra** las filas `Dep. bancario` de Cobranzas: tarjetas, notas de
   crédito o efectivo nunca aparecerán como depósito en el extracto.
2. **Concilia** cada cobranza con su movimiento usando el **N° de operación**
   como ancla; banco, fecha, importe y moneda sirven para calificar el cruce,
   no para descartarlo.
3. **Corrige** monedas mal etiquetadas cuando el Tipo de Cambio delata la
   moneda contraria.
4. **Alerta, no oculta:** si algo casi coincide (fecha con desfase, banco
   distinto, moneda que no cuadra), queda conciliado **con una alerta** y el
   detalle de la diferencia.
5. **Ningún registro desaparece en silencio:** las filas sin un dato
   indispensable (banco vacío o no reconocido, fecha inválida) van a **Sin
   Clasificar** con el motivo exacto.
6. **Errores de digitación:** monto inconsistente entre filas de un mismo
   Nro/Doc, registros duplicados (mismo N° de operación + Monto Total + Importe
   Parcial repetido) y filas sin banco.
7. **Bancos reconocidos:** BCP, Scotiabank (también por las siglas `SCT`/`SKT`),
   BBVA/Continental, Interbank y **Yape** como banco propio.

### Lectura de extractos

- **Excel consolidado:** cada hoja con tabla de movimientos es una cuenta; las
  de resumen o metadata se ignoran.
- **Libros de trabajo** `B_Bancos_XX.xlsx` (`extractosDesdeLibro`): el extracto
  viene pegado sin espacios en los encabezados y repetido junto a columnas
  propias (`Glosa`, `IMPORTE`, `OBS`). Se usa solo el primer bloque, gana la
  columna del banco (la más a la izquierda) y se omiten las hojas que no
  identifican un banco.
- **Columnas** (`autoMapExtractoColumns`): fecha e importe por la columna más a
  la izquierda que calce; N° de operación por encabezado exacto
  (`columnaNroOperacion`: `N° OPER.`, `Operación - Número`, `Nro. de
  operación`, `Num.Mvto`…).
- **Duplicados:** si varios libros repiten hojas de meses anteriores, un mismo
  movimiento (banco, moneda, fecha, importe, N° de operación) cuenta una vez.

## Qué hace con los pagos con tarjeta

El cruce tiene dos tramos y una regla de fondo.

**Regla de fondo:** el banco no abona venta por venta. Abona **la suma de todo
lo que el operador liquida ese día para ese código de comercio**: ventas de
cualquier mes menos costos, traslados de deuda, regularizaciones y
contracargos. Por eso un abono casi nunca tiene una cobranza de su mismo
importe, y hay que llegar a él a través del depósito del operador.

### Tramo 1 — Cobranza T.Crédito → Venta del operador (`matchTCreditoVentas`)

- Por marca (Visa/Mastercard → Niubiz, Amex, Diners), moneda e importe exacto,
  con una ventana de hasta 10 días. **Se asignan primero los pares más cercanos
  en fecha**, no en el orden de las filas (en octubre Q1, la fila 1100 del 05/10
  se quedaba con la venta del 13/10 de la fila 1172).
- Confianza: Alta (mismo día), Media (≤ 3 días), Baja (4–10 días). Un match
  "Baja" se descarta si hay una venta de **otra marca** más cercana: apunta a un
  error de tipeo en la marca.
- Una cobranza **partida en 2–3 filas** para una venta (mismo cliente, misma
  fecha, suma exacta), o una venta partida en 2–3 filas para una cobranza,
  también se reconoce.
- Una cobranza sin venta trae una pista: posible error de marca, posible error
  de moneda/tipo de cambio, posible desfase de fecha, o depósito fuera del
  periodo.

### Tramo 2 — Venta → Depósito → Abono en el banco (módulo `TJ`)

Es el mismo motor de `portable_tarjetas.html` (ver
`README_portable_tarjetas.md`, "Reglas del cruce"):

- La venta llega a su **fila exacta de depósito** por `ID operación` (Niubiz);
  Amex y Diners, por importe bruto, moneda y fecha.
- El día de ese depósito (banco + moneda + canal + fecha, con **todas** las
  filas del operador) se cruza con el abono en cuatro pasadas: directo (Diners
  con o sin cargo fijo de S/ 5.90 / US$ 2.07), acumulado de 2–3 días, **día
  pagado en 2 abonos**, y casos especiales (Amex abonado hasta 5 días después;
  Diners abonado en el otro banco). Cada caso especial lleva su comentario.
- Una quinta pasada marca los abonos que Niubiz pagó y **revirtió con una
  regularización** (los pagos duplicados por el bruto del 27–28/08/2026): en la
  hoja del banco salen **"REVERTIDO (TARJETAS)"** con la explicación, y en
  "Días vs Banco" como "Abono revertido por el operador".

### Dónde se ve el resultado

- **Tarjetas - Trazabilidad:** una fila por cada pago con tarjeta de Cobranzas
  **y por cada venta del operador que no está en Cobranzas** (solo las de las
  fechas de las cobranzas cargadas).
- **Estado de Cobranzas:** las filas T.Crédito entran con su estado: Cadena
  completa → *Cruce Exacto*; Contracargo, Depósito sin abono, Venta sin depósito
  → *Requiere Revisión*; Cobranza sin venta → *No Conciliado*. Columna
  "» N° op. abono (tarjeta)".
- **Hojas de cada banco:** los abonos de tarjeta explicados salen
  **"CONCILIADO (TARJETAS)"**, con los depósitos que los forman; los que ningún
  depósito explica, "SIN CONCILIAR" con alerta "Abono de tarjeta". Esto sale en
  el Excel aunque no se pulse "Conciliar Depósitos de Tarjeta".

## Cruce contra asientos ya contabilizados (revisión quincenal)

El asiento contable guarda el N° de operación bancario en la columna `Numero`.
Con el export de asientos cargado:

- Cada fila de **Conciliados** se busca exigiendo **Numero + Importe + Fecha a
  la vez**, para evitar falsos "Ya Contabilizado" si dos asientos comparten el
  número por un error de tipeo. El importe es el que no es cero entre
  `Debe`/`Haber`.
- **Abonos de tarjeta:** se buscan con el N° de operación del banco, el
  importe del abono (la suma del día) y su fecha. Un día pagado en 2 abonos
  queda "Ya Contabilizado" solo si lo están los dos; si no, "Parcial".
- La columna **Estado Contable** (Ya Contabilizado / Nuevo para Contabilizar)
  sale en Conciliados, en Tarjetas - Trazabilidad, en Días vs Banco y en la
  alerta de las filas "CONCILIADO (TARJETAS)". Sin el archivo, dice "N/A".

## El Excel que exporta

| Hoja | Contenido |
|---|---|
| Registro Original | El Registro de Cobranzas tal como se cargó. |
| Resumen | Métricas: % de cobranzas conciliadas (incluye T.Crédito confirmados en banco), % del banco conciliado (incluye abonos de tarjeta). |
| Conciliados | Pares cobranza ↔ banco, con alerta, facturas completas y estado contable. |
| Pendientes Cobranzas | Cobranzas sin movimiento en el banco. |
| Sin Clasificar | Filas sin un dato indispensable, con el motivo. |
| Una hoja por extracto (`BCP_PEN`, `BBVA_USD`, `SKT_PEN`, `IBK_USD`, `YAPE_PEN`…) | Todas las filas del extracto con columnas "»": estado (CONCILIADO / CONCILIADO (TARJETAS) / SIN CONCILIAR), cobranza asignada, diferencia, alerta y detalle. |
| Estado de Cobranzas | Vista solo de Cobranzas, deduplicada, con estado y motivo; incluye las filas T.Crédito. |
| Errores de Digitacion | Montos inconsistentes, duplicados, filas sin banco. |
| **Resumen Tarjetas** | **Pendientes a revisar** (qué, casos, importe, dónde verlo, qué hacer); ventas por mes, operador y moneda; pagos T.Crédito por mes; abonos de tarjeta por mes y banco con su estado contable. |
| **Tarjetas - Totales por Mes** | Por mes y moneda: ventas, depósitos netos del operador y abonos de tarjeta (BBVA + BCP), cada uno conciliado / no conciliado; la diferencia depósitos conciliados − abonos conciliados se explica por desfase entre meses y cargo fijo de Diners, y la "Diferencia sin explicar" debe ser 0. Octubre Q1: 0 en soles, 4.14 USD (cargo Diners). |
| **Tarjetas - Conciliados** | También como pestaña en pantalla. Solo las filas T. Credito con la cadena completa: primero las columnas originales de Cobranzas (Cob_), luego venta, depósito, comisión, abono y regla del cruce. **Validación**: "Válido" (regla directa, todo exacto) o "Válido con nota" (otra regla, cobranza o venta partida, o algo que corregir). **Qué corregir**: fecha de Cobranzas distinta a la de la venta, redondeo, o marca que no corresponde al operador. Octubre Q1: 168 filas, 131 válidas y 37 con nota (3 de fecha, 2 de redondeo). |
| **Tarjetas - Trazabilidad** | Cobranza → Venta → Depósito → Abono, una fila por pago o venta, con estado, comentario y estado contable. |
| **Tarjetas - Días vs Banco** | Un día de depósito por código de comercio frente a su abono, con composición, método, estado y explicación; al final, los abonos que ningún día explica. |
| **Tarjetas - Ajustes** | Contracargos, traslados de deuda, cargos, regularizaciones y saldo deudor; al final, regularizaciones vs abonos revertidos. |
| *Depositos + Banco* (×3) | **Ocultas** (clic derecho en una pestaña → Mostrar): las hojas del operador con todas sus columnas originales y su abono. |

Las hojas de tarjetas solo salen si se cargaron ventas y depósitos de tarjetas
**y** extractos.

**Periodo de las hojas de tarjetas:** "Días vs Banco", "Ajustes" (también sus
resúmenes), las regularizaciones, las ventas retenidas y las hojas ocultas
muestran solo los **meses calendario de los abonos de tarjeta de los extractos
cargados** (con extractos de octubre, solo octubre). La Trazabilidad cubre las
fechas de las cobranzas cargadas.

Estados de la Trazabilidad: Cadena completa · Contracargo · Depósito sin abono ·
Venta sin depósito (avisa si el archivo de depósitos aún no llega a esa fecha) ·
Cobranza sin venta. El significado de cada estado y de cada método de cruce
está en `README_portable_tarjetas.md`, "El Excel que exporta".

## Resultados de referencia

**Octubre Q1 2026** (`Conciliacion octubre q1/`: Cobranzas al 20/10,
`Bancos_Q1_20102026.xlsx` con 9 cuentas, `Tarjeta_de_credito.xlsx`):

| Métrica | Resultado |
|---|---|
| Cobranzas conciliadas | 509 de 612 (incluye 168 de 209 T.Crédito confirmados en banco) |
| Movimientos del banco conciliados | 408 de 497 (incluye 67 abonos de tarjeta) |
| Dep. bancario en Estado de Cobranzas | 292 Cruce Exacto · 49 Requiere Revisión · 62 No Conciliado |
| Abonos de tarjeta | Todos explicados (BBVA y BCP, soles y dólares) |
| T.Crédito pendientes | 7 sin venta (3 con posible error de marca), 21 ventas de los últimos días sin depósito aún, 14 depositadas sin abono (extracto BBVA hasta el 16/10) |
| Ventas del operador sin registro en Cobranzas | 4 |

**Julio 2026** (Cobranzas de julio, libros `B_Bancos`): 306 T.Crédito → 286
cadena completa, 6 depósito sin abono, 14 sin venta; 14 ventas del operador sin
registro en Cobranzas.

Los resultados de abril a octubre del lado del operador están en
`README_portable_tarjetas.md`, "Resultados de referencia".

## Referencia técnica

### Módulo `TJ` (código compartido con `portable_tarjetas.html`)

`const TJ = (function () { … })()` contiene una copia de los lectores y del
cruce de `portable_tarjetas.html` (con `operadorFile` → `cruceTarjetasFile`):
`leerVentas`, `leerDepositosNiubiz`, `leerDepositosAmex`, `leerDepositosDiners`,
`leerAbonosBancarios`, `construirGruposBanco`, `cruzarDepositosBanco`,
`cargosFijosDiners`, `estadosDeGrupos`, `categoriaDeposito`,
`construirDepositosConBanco`, `construirDiasSinAbono`,
`construirAjustesOperador(filtro)`. **Si cambias esa lógica en una herramienta,
haz el mismo cambio en la otra.**

### Funciones propias del 2.5

- **Extractos:** `extractosDesdeLibro`, `recortarPrimerBloque`,
  `autoMapExtractoColumns`, `primeraColumna`, `columnaNroOperacion`,
  `findGlosaColumnIndex`.
- **Tarjetas, tramo 1:** `buildCobranzasTCredito`,
  `buildCobranzasTCreditoUnits`, `buildVentasPool`, `matchTCreditoVentas`,
  `findSplitCobranzaMatches`, `findVentaSplitMatches`, `posibleErrorMarcaObs`,
  `pistaMonedaOFecha`, `ventaFueraDePeriodoObs`.
- **Tarjetas, tramo 2:** `matchTarjetasDepositos` (traduce el resultado de `TJ`
  al formato de la pestaña, del enlace con T.Crédito y de las métricas),
  `linkTCreditoToBankDeposits`, `incorporateTarjetasIntoMetrics`.
- **Hojas de tarjetas:** `construirTrazabilidadTarjetas`,
  `construirDiasVsBanco`, `construirResumenTarjetas`, `construirTotalesPorMes`, `infoAbonosTarjeta`,
  `estadoContableAbono`, `calcularRangoExtractosTarjeta`, `enPeriodoExtractos`.
- **Asientos:** `handleAsientosFileSelect`, `buscarAsientoContable`.

### Cómo verificar sin navegador

Igual que en `README_portable_tarjetas.md`, "Cómo verificar cambios sin
navegador": se ejecuta el último bloque `<script>` con `vm.runInContext` y un
`document` simulado, se cargan `cobranzasFile`, `extractosFiles` y
`cruceTarjetasFile`, se fijan los selects de mapeo (`cobColFecha`…,
`filterBCP`…), y se llama a `executeCoreLogic()`,
`processTarjetasReconciliation()` y `exportResultsToExcel()` (las dos últimas
con `setTimeout`) interceptando `XLSX.writeFile`.

### Limitaciones

- **No se probó en un navegador**: todo se verificó con Node sobre archivos
  reales.
- La pestaña "Pendiente en Bancos" y los porcentajes en pantalla solo
  descuentan los abonos de tarjeta después de pulsar "Conciliar Depósitos de
  Tarjeta"; el Excel los marca siempre.
- Si se cargan varios libros que repiten meses, las copias repetidas de un
  extracto salen "SIN CONCILIAR" en su propia hoja (el movimiento se concilia
  en la primera copia).

## Historial de cambios

Detalle completo en `git log` (rama `feature/v24-tarjetas-notas-credito`).

**Novedades de 2.5 frente a 2.4:** Sin Clasificar; Excel consolidado multi-hoja;
Yape como banco propio; Scotiabank por `SCT`/`SKT`; errores de digitación
ampliados (duplicados y sin banco); cruce contra asientos ya contabilizados.

**Cambios posteriores (octubre 2026):**

| Cambio | Detalle |
|---|---|
| Lectura de libros de trabajo | Primer bloque, columna más a la izquierda, N° de operación por encabezado, duplicados entre libros (sin ceros a la izquierda). |
| Motor de tarjetas unificado | La Fase 3B usa el módulo `TJ`; se quitó el cruce duplicado anterior. |
| Tarjetas - Trazabilidad | Cobranza → Venta → Depósito → Abono, más ventas del operador sin registro en Cobranzas. |
| Estado de Cobranzas y hojas de banco | T.Crédito con su estado; abonos "CONCILIADO (TARJETAS)". |
| Estado contable de abonos y Resumen Tarjetas | Cruce con asientos; pendientes a revisar. |
| Hojas de tarjetas de 11 a 4 (+3 ocultas) | Se eliminaron OBS_TC, Tarjetas Depositos, TCredito Consolidado, Días sin Abono, Regularizaciones vs Abonos; solo el periodo de los extractos. |
| T.Crédito ↔ Ventas | Pares más cercanos en fecha primero. |
| Motor del abono | Día pagado en 2 abonos; cargo fijo de Diners; Amex hasta 5 días; Diners en el otro banco; abonos revertidos por regularización. |
| Totales por mes | Hoja "Tarjetas - Totales por Mes"; `Depositos amex` ignora la tabla pegada debajo con otro formato. |
| Amex por suma | En la trazabilidad, una venta Amex se enlaza a la fila diaria de `Depositos amex` por la suma de las ventas del día, o de las retenidas pagadas juntas (mismo código, hasta 60 días). |
| Estados de cuenta desde PDF | Hojas `EECC_BBVA_2026-01_MN` / `EECC_BCP_…`: `OPER.` y `FECHA PROC.` como fecha, `NUM OP`, `CARGO / ABONO`, `MN`/`ME`; grupos del día de tarjeta separados por código de comercio. |
