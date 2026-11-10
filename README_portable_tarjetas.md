# `portable_tarjetas.html` — Conciliador Independiente de Tarjetas

Herramienta **portable de un solo archivo HTML**, separada de
`portable_2.5.html`, para validar **exclusivamente** la cadena de tarjetas de
crédito/débito: **Ventas → Depósitos del operador → Abono en banco**. No lee
ni depende del Registro de Cobranzas. Corre 100% en el navegador: no instala
nada ni envía datos a ningún servidor.

**Contenido**

1. [Por qué existe](#por-qué-existe-aparte-de-portable_25html)
2. [Qué hace, en una página](#qué-hace-en-una-página)
3. [Cómo usarlo](#cómo-usarlo)
4. [Archivos de entrada](#archivos-de-entrada)
5. [El Excel que exporta](#el-excel-que-exporta)
6. [Reglas del cruce](#reglas-del-cruce)
7. [Cómo leer una trazabilidad (ejemplos reales)](#cómo-leer-una-trazabilidad-ejemplos-reales)
8. [Contracargos, traslados de deuda y regularizaciones](#contracargos-traslados-de-deuda-y-regularizaciones)
9. [Asientos contables por periodo](#asientos-contables-por-periodo)
10. [Resultados de referencia (enero–octubre 2026)](#resultados-de-referencia-eneroctubre-2026)
11. [Referencia técnica y conocimiento acumulado](#referencia-técnica-y-conocimiento-acumulado)
12. [Historial de cambios](#historial-de-cambios)

---

## Por qué existe (aparte de `portable_2.5.html`)

En `portable_2.5.html`, el cruce de tarjetas pasa por el Registro de
Cobranzas, donde puede haber errores de tipeo en la marca de tarjeta o en el
número de documento. Un descuadre ahí puede ser un error interno de registro,
no un problema real de depósito o banco.

`Ventas`, `Depositos` y el extracto bancario, en cambio, vienen de fuentes
oficiales: los propios operadores (Niubiz, Amex, Diners) y el banco. Ese cruce
debería cuadrar prácticamente al 100%, así que sirve como verificación limpia:
si algo no cuadra aquí, es un problema real (depósito que no llegó, cargo del
operador, ajuste), no un error de tipeo.

Desde octubre 2026 las dos herramientas usan **el mismo motor de depósitos ↔
banco y la misma estructura de hojas de tarjetas** en el Excel (ver
[Historial de cambios](#historial-de-cambios)).

## Qué hace, en una página

1. **Lee el archivo del operador** (`Ventas`, `Ventas amex`, `Depositos`,
   `Depositos amex`, `Depositos Diners`).
2. **Lee los extractos bancarios** en cualquiera de sus formatos: archivos
   sueltos por banco/moneda, un Excel consolidado con una hoja por cuenta, o un
   libro de trabajo `B_Bancos_XX.xlsx`. Reconoce por la glosa los abonos de
   tarjeta (`ABONO VISANET`, `COMPAÑIA PERUANA DE MEDIOS`, `DINERS CLUB`,
   `COMPAÑIA DE SERVICIOS CONEXOS`…).
3. **Venta → Depósito:** une cada venta con su fila de depósito (Niubiz por
   `ID operación`; Amex por `Código de autorización` vía `Ventas amex`; Diners
   por importe y fecha).
4. **Depósito → Banco:** suma todo lo que el operador liquida en el día por
   código de comercio (ventas de cualquier mes, cargos, traslados, ajustes) y
   lo cruza con el abono del extracto, en cuatro pasadas (ver
   [Reglas del cruce](#reglas-del-cruce)).
5. **Explica lo que no cuadra:** días con neto 0 o negativo, regularizaciones,
   contracargos, saldo deudor, cargo fijo de Diners, abonos tardíos o
   partidos.
6. **Exporta un Excel** con un resumen de pendientes, la trazabilidad venta por
   venta, el cruce día por día, los ajustes del operador, las hojas por banco,
   los asientos contables y el cierre de la cuenta por cobrar al operador.

## Cómo usarlo

1. Abre `portable_tarjetas.html` con un navegador moderno.
2. Carga el archivo del operador en "Ventas y Depósitos (Operadores)".
3. Carga los extractos bancarios (uno o varios archivos, o un Excel
   consolidado). Incluye el extracto del **mes siguiente** si quieres
   confirmar los depósitos de fin de mes, que se abonan después.
4. Elige el periodo: un mes, o **"Todos los periodos"** para analizar varios
   meses de una vez. Ajusta las tolerancias si hace falta (por defecto ±2 días
   y ±0.01).
5. Ejecuta la conciliación y revisa las pestañas en pantalla.
6. Exporta el Excel. Empieza por la hoja **Resumen Tarjetas → Pendientes a
   revisar**: cada fila dice cuántos casos hay, cuánto suman, en qué hoja
   verlos y qué hacer.

## Archivos de entrada

| Archivo | Contenido |
|---|---|
| Archivo del operador: `Conciliacion octubre q1/Tarjeta_de_credito.xlsx` (vigente, octubre) o `bancos/Ventas y depositos_tarjetas.xlsx` (12/10, hasta septiembre) | Hojas `Ventas`, `Ventas amex`, `Depositos`, `Depositos amex`, `Depositos Diners`. Las demás hojas (`TD`, `SALDOS`, `COB`, `4_Bancos`, `Contabilidad`…) se ignoran. |
| Extractos bancarios | Abonos de tarjeta identificables por glosa. Formatos: archivos sueltos BBVA/BCP/Scotiabank/Interbank, Excel consolidado (`Bancos_Q1_20102026.xlsx`, una hoja por cuenta) o libros de trabajo `bancos/B_Bancos_04…09.xlsx`. |

Los `*.xlsx` están en `.gitignore`: viven solo en la carpeta local.

## El Excel que exporta

9 hojas visibles y 3 ocultas, en este orden:

| # | Hoja | Qué contiene |
|---|---|---|
| 1 | **Resumen Tarjetas** | Periodo y corte del archivo de depósitos. **Pendientes a revisar** (qué, casos, importe por moneda, dónde verlo, qué hacer). Ventas por mes, operador y moneda con su estado y % de cadena completa. Abonos de tarjeta por mes, banco y moneda: conciliados y sin depósito, con importes. |
| 1b | **Tarjetas - Totales por Mes** | Por mes y moneda: **ventas** (fecha de venta), **depósitos netos del operador** (fecha de depósito) y **abonos de tarjeta en BBVA + BCP** (fecha del abono), cada uno en total, **conciliado y no conciliado**. Abonos no conciliados separados en revertidos por el operador y sin depósito. La diferencia *depósitos conciliados − abonos conciliados* se descompone en desfase entre meses y cargo fijo de Diners / redondeo; la última columna, **Diferencia sin explicar**, debe ser 0. Al final, total del periodo por moneda. |
| 2 | **Tarjetas - Trazabilidad** | Una fila por venta del periodo: operador, marca, producto, ID de operación, código de autorización, fila de `Ventas`, fecha, moneda, importe → fecha y neto del depósito, código de comercio, canal, banco, neto del día → fecha, N° de operación e importe del abono → **estado de la cadena** y comentario → **comisión, regla del cruce, validación y qué corregir** (ver [Validación de la cadena](#validación-de-la-cadena)). Con filtro automático. |
| 3 | **Tarjetas - Días vs Banco** | Una fila por día de depósito y código de comercio (la unidad con la que paga el banco), de los tres operadores: composición del día (ventas, costos del operador, traslado de deuda, regularización, contracargo, saldo deudor, otros), neto, abono (fecha, N°, importe, diferencia), método de cruce, estado, explicación y detalle. Al final, los **abonos de tarjeta que ningún día explica**. |
| 4 | **Tarjetas - Ajustes** | Contracargos, traslados de deuda (con su contrapartida), costos de disputa, regularizaciones y saldo deudor; resumen por categoría y efecto por banco; al final, **regularizaciones vs abonos revertidos**. |
| 5–6 | **BBVA - Tarjetas**, **BCP - Tarjetas** | Solo los movimientos de tarjeta de cada banco, con su estado (Conciliado / No conciliado). |
| 7 | **Asientos por Periodo** | Asientos de venta y de depósito por mes, operador y moneda (ver [Asientos](#asientos-contables-por-periodo)). |
| 8 | **Cierre CxC Operador** | Por mes de venta: cuánto se depositó dentro del mes, en el mes siguiente, cuánto está en desfase y cuánto sin explicación. |
| — | *Depositos + Banco*, *Depositos amex + Banco*, *Depositos Diners + Banco* | **Ocultas.** Las hojas del operador con todas sus columnas originales (tarjeta, autorización, lote, comisión, IGV…) y al final la trazabilidad hasta el banco. Para verlas: clic derecho en cualquier pestaña → **Mostrar**. |

### Estados de la Trazabilidad

| Estado | Significa | Qué hacer |
|---|---|---|
| Cadena completa | La venta tiene depósito y el día de ese depósito tiene abono. | Nada. Si el cruce usó una regla especial, el comentario lo dice (Amex tardío, Diners en otro banco, cargo fijo). |
| Contracargo | Llegó al banco y después el operador la descontó. | Ver el comentario (fecha, importe y abono del que se descontó) y "Tarjetas - Ajustes". |
| Depósito sin abono | El operador la depositó, pero su día no tiene abono. | Ver el motivo del día en "Días vs Banco" (neto 0, regularización, fuera de extractos…). |
| Falta en Depositos amex | Amex: `Ventas amex` confirma la fecha de depósito, pero la fila no está en `Depositos amex`. | Pedir a Amex el detalle de depósitos faltante. |
| Venta sin depósito (corte) | La venta es de los últimos días antes del corte del archivo de depósitos. | Normal: se deposita en el siguiente corte. |
| Venta sin depósito (revisar) | La venta es anterior al corte y no tiene depósito. | Reclamar al operador. |

### Validación de la cadena

La pestaña **Cadena Completa** y las cuatro últimas columnas de "Tarjetas - Trazabilidad" dicen si cada venta con cadena completa se da por válida:

| Columna | Contenido |
|---|---|
| Comisión (bruto − neto) | Lo que el operador descontó de esa venta. |
| Regla del cruce | Pasada que concilió el día con el banco: `Directo`, `Acumulado (2 días)`, `Dividido (2 abonos)`, Amex/Diners con desfase o en otro banco, o `Sin depósito del operador (abono directo)`. |
| Validación | **Válido**: venta, depósito y abono cuadran exactos por la regla directa. **Válido con nota**: cuadra por otra regla o hay algo que corregir. Vacía si la cadena no está completa. |
| Qué corregir | **Redondeo** entre la venta y el bruto depositado, u **Operador** si la marca no es del archivo en que está la venta (p. ej. Amex en Ventas de Niubiz). |

Sin Registro de Cobranzas esta herramienta no puede revisar fecha ni redondeo contra la cobranza: esas correcciones salen en la hoja "Tarjetas - Conciliados" de `portable_2.5.html`.

Octubre Q1: 174 ventas en Cadena Completa; 140 válidas, 21 válidas con nota (19 por acumulado de 2 días y 2 por día pagado en 2 abonos), 13 con depósito aún sin abono; nada que corregir.

### Estados de Días vs Banco

| Estado | Significa |
|---|---|
| Abonado | El día tiene su abono. La columna **Método** dice cómo: `Directo`, `Acumulado (N días)`, `Dividido (2 abonos)`, `Directo (cargo fijo Diners ×k)`, `Directo (Amex: abono N días después)`, `Directo (Diners en BBVA/BCP…)`. |
| Sin abono: explicado | Compensación interna del operador: neto 0, neto negativo (cargo), regularización, contracargo cubierto, crédito aplicado a un saldo deudor, ventas retenidas, neto menor que el cargo fijo de Diners. |
| Sin abono: revisar | Sin explicación automática. La columna Detalle propone el abono más cercano y la diferencia. |
| Abono sin depósito | (Al final de la hoja) un abono con glosa de tarjeta que ningún día explica. |
| Abono revertido por el operador | (Al final de la hoja) un abono que Niubiz pagó y luego retiró con una regularización del mismo importe (pasada 5). |

### Periodo del reporte

- Con **un mes** elegido, las hojas 1 a 4 y las ocultas muestran ese mes
  (la Trazabilidad, por fecha de venta; Días vs Banco y Ajustes, por fecha de
  depósito).
- Con **"Todos los periodos"**, muestran los **meses calendario completos que
  cubren los abonos de tarjeta de los extractos cargados**. Con extractos de
  abril a octubre, de abril a octubre: las ventas de enero a marzo no aparecen,
  porque no hay banco contra qué cruzarlas.
- `Asientos por Periodo` y `Cierre CxC Operador` cubren siempre todo el archivo
  del operador.

## Reglas del cruce

### Venta → Depósito (`cruzarVentasDepositos`)

- **Niubiz:** `ID operación` exacto. Es la misma clave en `Ventas` y en
  `Depositos`; no hay ambigüedad.
- **Amex:** `Código de autorización` contra `Ventas amex`, que da la fecha de
  depósito real por transacción (`buildFechaDepositoAmexPorAuth`). Respaldo:
  operador + moneda + importe bruto + fecha (el depósito nunca antes de la
  venta, hasta 15 días).
  - `Depositos amex` trae **una fila por establecimiento y día con la suma de las
    ventas** (MONTO TOTAL). Si una venta no calza sola, se busca el grupo de ventas
    Amex de la misma moneda, hasta 7 días antes y con la fecha de `Ventas amex`, cuya
    suma sea ese bruto (24/09/2026: 886.37 = 688.84 + 197.53).
  - **Ventas retenidas y pagadas juntas:** si tampoco hay grupo del día, se busca por
    **código de establecimiento, hasta 60 días antes** (19/06/2026: Amex pagó por código
    las "Retención abonada" con ventas del 23/05 al 16/06; 05/08: las dos "Venta
    retenida" del 12–13/06 por 7,295.18). Método "Suma de N ventas Amex…".
- **Diners:** operador + moneda + importe bruto + fecha (1 fila de
  `Depositos Diners` = 1 transacción).
- **Ventas sin depósito** (`clasificarVentaSinDeposito`): "corte" si son
  posteriores o cercanas (≤ 3 días) al último depósito del archivo;
  "Depositado (falta en Depositos amex)" si `Ventas amex` confirma la fecha;
  "Sin explicación" en otro caso.

### Depósito → Banco (`construirGruposBanco` + `cruzarDepositosBanco`)

**Regla de fondo:** el banco no abona venta por venta. Abona **la suma de todo
lo que el operador liquida ese día para ese código de comercio** (cada código por
separado: en marzo hubo dos códigos POS en la misma cuenta BBVA y se abonaron por separado): ventas (de
cualquier mes) menos costos del operador, traslados de deuda, regularizaciones
y contracargos. Por eso los grupos se arman con **todas** las filas de
`Depositos` de ese banco, moneda, canal y fecha de depósito. Las filas que no
son venta no traen producto: su canal (POS / Pago Link) sale del código de
comercio.

Un grupo calza con un abono si coinciden banco, moneda y familia de glosa
(Niubiz / Amex / Diners). En BBVA POS la glosa además debe contener el código
de comercio (`ABONO VISANET 5656321223`); Pago Link y BCP no lo traen. El
cruce corre en cuatro pasadas, cada una solo sobre lo que la anterior dejó sin
abono:

| Pasada | Regla | Método en el Excel |
|---|---|---|
| 1 | Un grupo = un abono, fecha ±2 días, importe ±0.01. **Diners:** también si el abono es el neto menos 1 a 3 cargos fijos (S/ 5.90 / US$ 2.07); gana siempre el exacto. | `Directo`, `Directo (cargo fijo Diners ×k: importe)` |
| 2 | Un abono que junta 2 o 3 días de depósito. | `Acumulado (N días)` |
| 3 | Un día pagado en **2 abonos** del mismo banco, moneda y operador, dentro de ±2 días, cada uno menor que el neto, que juntos lo cuadran. | `Dividido (2 abonos)` |
| 4 | **Amex:** abono del mismo banco e importe hasta **5 días después** del depósito (feriados, fines de semana). **Diners:** abono en el **otro banco** que cuadra exacto o con el cargo fijo. Cada caso lleva su comentario. | `Directo (Amex: abono N días después)`, `Directo (Diners en BBVA…)` |
| 5 | **Abono revertido por el operador:** un abono de Niubiz que sigue sin depósito y tiene, en los 7 días siguientes, una "Regularización gestión operativa" del mismo banco, moneda e importe. Si además es el **bruto** de las ventas de un día cuyo neto también se abonó, el comentario lo identifica como **pago duplicado por el bruto**. No cuenta como conciliado: queda al final de "Días vs Banco" con estado "Abono revertido por el operador" y su explicación. | (sin método: estado propio) |

Otras reglas:

- **Fechas:** el banco abona entre 1 día antes y el siguiente día hábil de la
  "fecha de depósito" de Niubiz (domingo → lunes; sábado → viernes anterior).
  Por eso la tolerancia es ±2 días.
- **Cobertura de extractos** (`hayExtractoCerca`): para decir que un abono no
  llegó se exigen movimientos del banco a ambos lados de la fecha. Un mes sin
  extracto, o el borde de un extracto, sale como "Fuera de los extractos
  cargados".
- **Lectura de columnas:** gana la columna más a la izquierda que calce
  (fecha, importe, glosa); el N° de operación se busca por encabezado exacto
  (`N° OPER.`, `Operación - Número`, `Nro. de operación`, `Num.Mvto`).
- **Duplicados:** un movimiento (banco, moneda, fecha, importe, N° de
  operación sin ceros a la izquierda) cuenta una vez aunque esté en varios
  libros.

## Cómo leer una trazabilidad (ejemplos reales)

Para demostrar que una venta llegó al banco no basta con su fila de
`Depositos`: hay que mostrar el **total del día de ese código** y el abono que
lo iguala. Eso es lo que hace cada fila de "Tarjetas - Días vs Banco", y lo
que se ve en las pestañas `Día dd-mm (código)` de los archivos de trazabilidad
armados aparte.

### Ejemplo 1 — venta normal con abono: ID `423194875749118` (parte 1)

| Eslabón | Fuente | Dato |
|---|---|---|
| Venta | `Ventas` fila 2337 | 07/08/2026, USD 2,100.00, Mastercard, Pago Link, código 811777741, DCC = Sí (Monto DCC 11,576.04) |
| Depósito del operador | `Depositos` fila 2512 | fecha de depósito 09/08, bruto 2,100.00 − comisión 86.73 − IGV 8.13 = neto **2,005.14** |
| Total del día | `Depositos`, código 811777741, 09/08 | 6 filas (3 ventas + 2 "Uso del Sist CE otras marcas ME" + 1 "Costo por trx liquidada") = **4,582.79** |
| Abono en el banco | BBVA USD, N° operación **59582** | 10/08/2026, USD 4,582.79, glosa `R20341198217COMPA#IA PERUANA DE MEDIOS D`; diferencia 0.00 |

El 09/08/2026 fue domingo: Niubiz fecha el depósito el domingo y el banco
abona el lunes 10/08.

### Ejemplo 2 — el contracargo del mismo ID (25/09)

Un **contracargo** es una venta que el titular de la tarjeta no reconoce y por
la que pide el reembolso; Niubiz lo descuenta de lo que deposita. En este ID,
`Depositos` fila 2882 ("Cargos por Contracargos"): operación del 23/09, fecha
de depósito **25/09/2026, −USD 2,236.57**, 136.57 más que la venta.

Día completo del código 811777741 (BBVA / USD / Pago Link) el 25/09:

| Fila | Concepto | Neto |
|---|---|---|
| 2882 | Cargos por Contracargos (ID 423194875749118) | −2,236.57 |
| 2883 | Costo por trx liquidada Com Electronico | −0.18 |
| 2885 | Traslado de Deuda | +0.18 |
| 2886 | Traslado de Deuda | +2,122.59 |
| 3002 | Venta del día (ID 725276018955597) | +113.98 |
| | **Neto del día** | **0.00** |

Niubiz cubrió el contracargo con la venta del propio día (113.98) y con un
**traslado de deuda** desde otro código (2,236.57 − 113.98 = **2,122.59**).
Como el día suma cero, **no hay abono en BBVA**.

El otro lado es el código 966661760 (POS BCP USD): entrega los traslados
(−2,122.59 y −0.18), sus 3 ventas netas suman 5,310.12 y el neto del día es
**3,187.35**, exactamente el abono de BCP USD del 24/09 (**N° 8353799**, glosa
`DE COMPANIA PERUANA DE`). El débito del contracargo lo absorbió la cuenta de
BCP, aunque la venta era de BBVA.

**Hipótesis pendiente (los 136.57):** la venta tenía DCC (el titular pagó en
su propia moneda, Monto DCC 11,576.04). Es probable que el contracargo se haya
calculado reconvirtiendo ese monto. No está confirmado; es la pregunta para
Niubiz.

### Ejemplo 3 — contracargo absorbido por las ventas del día: ID `919093786579754`

| Eslabón | Dato |
|---|---|
| Venta | `Ventas` fila 330, 17/06/2026, USD 2,218.70, Visa internacional, Pago Link, código 811777741 |
| Depósito | `Depositos` fila 1922, depósito 19/06, bruto 2,218.70 − comisión 88.53 − IGV 8.07 = neto 2,122.10 |
| Total del día 19/06 | 3 filas (2 ventas + 1 cargo) = 2,183.92 |
| Abono | BBVA USD 18/06/2026, **N° 59274**, USD 2,183.92 |
| Contracargo | `Depositos` fila 2014: operación 30/07, depósito 01/08, **−2,218.70** (43 días después de la venta; la comisión y el IGV del depósito original, 96.60, no se devuelven) |
| Total del día 01/08 | 6 ventas del 30/07 (2,552.78) + 2 cargos (−1.24) + contracargo (−2,218.70) = **332.84** |
| Abono | BBVA USD 31/07/2026 (01/08 fue sábado), **N° 59532**, USD 332.84 |

Aquí no hubo traslado: el débito se descontó de las ventas del mismo día.
Pérdida neta: USD 96.60 (cobró 2,122.10 y devolvió 2,218.70).

### Ejemplo 4 — día pagado en dos abonos: Amex BCP USD, 14/10/2026

Neto del día 782.02; el banco pagó **527.18 (N° 85004743) + 254.84
(N° 49336338)** el mismo día. Método `Dividido (2 abonos)`. En el histórico
hay otros 5 días Amex iguales (por ejemplo S/ 406.98 del 19/06 = 259.95 +
147.03), siempre con N° de operación casi consecutivos.

### Ejemplo 5 — cargo fijo de Diners

- 01/10/2026, BCP USD: neto 180.98, abono **176.84** = 180.98 − 2 × 2.07.
- 05/10/2026, BBVA PEN: neto 5.70, menor que el cargo de 5.90 → no hay abono
  ("Neto menor que el cargo fijo de Diners").
- 20/04/2026: el operador indica BCP PEN (2,415.71) pero el abono llegó a
  **BBVA PEN** como 2,409.81 (−5.90). Comentario: confirmar con Diners la cuenta
  de pago de ese código.

## Contracargos, traslados de deuda y regularizaciones

**Contracargo:** venta que el titular no reconoce; el operador la descuenta
(importe negativo en `Depositos`, tipo "Cargos por Contracargos"). El débito no
siempre cae en la cuenta de la venta.

**Traslado de deuda** (inferido de los datos; confirmar con el operador): pares
del mismo día, de igual monto y signo opuesto, entre códigos de comercio de la
misma empresa. El código con importe negativo es la cuenta que absorbe el
débito: el operador usa lo que le debe pagar a un código para cubrir el saldo
deudor de otro. Esos montos no se abonan al banco mientras haya deuda: no es un
depósito perdido, es una compensación interna.

**Regularización del 29/08/2026:** el 27–28/08 el banco recibió abonos en los
seis códigos; el 29/08 el operador los revirtió con una "Regularización gestión
operativa" por el mismo monto en cada código (Pago Link USD: **USD 20,992.89**,
abono N° 59690). Pago Link quedó con saldo deudor, repagado con ventas propias
y traslados desde los códigos POS: −16,707.33 → −16,038.05 → −11,252.83 → 0.00
el 02/09. Las ventas POS retenidas esos días (S/ 7,640.76 y USD 8,549.25) se
cobraron en el sistema pero no llegaron al banco. Por eso agosto tiene ~40
ventas "Depósito sin abono".

**Qué revirtió la regularización (resuelto, octubre 2026):** los abonos del
27–28/08 eran **pagos duplicados**. El 27/08 BBVA abonó US$ 20,992.89 (N° 59690),
que es el **bruto** de las 4 ventas Pago Link USD del 26/08 (48.85 + 4,290.56 +
16,143.72 + 509.76); el 28/08 llegó además su **neto**, US$ 20,310.81 (N° 59694).
La diferencia de 682.08 son las comisiones e IGV de esas ventas (510.46) más
"Cuotas sin intereses Abonos" (171.62). La regularización retiró el pago de más
y el saldo deudor posterior lo recuperó. Lo mismo en los otros cinco códigos
(898.71, 779.01, 515.04, 507.60, 146.16). Detalle paso a paso:
`Trazabilidad_Abono_59690.xlsx`. Queda por confirmar con Niubiz el motivo del
pago adelantado ("Tipo de Abono" = "Adelanto de Abono").

"Tarjetas - Ajustes" lista todo esto; "Días vs Banco" explica cada día afectado.

## Asientos contables por periodo

La hoja `Asientos por Periodo` arma los asientos con cuentas PCGE
referenciales (confirmar subcuentas con el contador):

- **Moneda original.** Los archivos del operador no traen tipo de cambio: el
  asiento va en soles y dólares por separado, con una columna `TC` para
  convertir a PEN (fórmulas ya puestas). El asiento mensual es una
  aproximación; para precisión, asentar por día.
- **Ventas** (por mes de la venta): CxC al operador contra Ventas e IGV
  (importe / 1.18). Excluye anuladas.
- **Depósitos** (por mes del depósito, incluye ventas de meses anteriores):
  Bancos, comisión e IGV de la comisión contra la CxC.
- **Otros conceptos del operador** (mantenimiento, costo por transacción,
  traslado de deuda…) van en asientos aparte porque también cambian el abono.
- **Líneas "revisar":** lo que no se clasifica solo (por ejemplo la "Deuda
  anterior" de Amex, US$ 2.01 / S/ 5.90).
- Al final, comprobación Debe = Haber por moneda.

`Cierre CxC Operador` da, por mes de venta, el saldo de la CxC al operador al
cierre: debe cuadrar con el mayor contable.

**Por qué importa la fecha de venta vs. la de depósito:** la venta se reconoce
el día en que ocurre; el dinero llega 1–3 días después y a veces cambia de mes.
Entre ambas fechas el saldo vive en la CxC al operador.

```
Venta (hoja Ventas)
  Debe   12 CxC Operador                      total cobrado
  Haber  70 Ventas                            base imponible
  Haber  40 IGV                               IGV

Depósito (hoja Depositos: Suma Depositada, Comisión, IGV de la misma fila)
  Debe   10 Bancos                            neto depositado
  Debe   63 Comisión                          comisión sin IGV
  Debe   40 IGV crédito fiscal                IGV de la comisión
  Haber  12 CxC Operador                      total cobrado
```

El abono del extracto no genera un asiento nuevo: confirma el del depósito. Si
el banco abona distinto, la diferencia está en "Días vs Banco" (columna
`Diferencia (abono − neto)` y la explicación del método).

**Ventas en dólares:** la CxC en USD es partida monetaria; la diferencia entre
el tipo de cambio de la venta y el del depósito va a 776 / 676 (diferencia de
cambio). Ejemplo, venta US$ 100, comisión 3.5% + IGV, TC 3.700 → 3.750:

```
Venta:     Debe 12 CxC 370.00 | Haber 70 Ventas 313.56 | Haber 40 IGV 56.44
Depósito:  Debe 10 Bancos 359.51 | Debe 63 Comisión 13.13 | Debe 40 IGV 2.36   (= 375.00)
           Haber 12 CxC 370.00 | Haber 776 Diferencia de cambio 5.00
```

**Contracargos (propuesta, confirmar con el contador):** en la fecha del
contracargo, Debe 1219 Contracargos en disputa / Haber 1212 CxC Operador; si se
pierde la disputa, Debe 6591 / Haber 1219. Una diferencia entre el contracargo
y la venta (los 136.57 del ID 423194875749118) va a gasto provisional hasta que
el operador la explique. El traslado de deuda no lleva asiento propio.

## Resultados de referencia (enero–octubre 2026)

Corrida "Todos los periodos" con `Tarjeta_de_credito.xlsx` (octubre; depósitos
hasta el 20/10) y todos los extractos del año: `bancos/Bancos enero_marzo_2026.xlsx`
(estados de cuenta BBVA y BCP de enero a marzo), `B_Bancos_04…09` y
`Bancos_Q1_20102026`. Excel: `Conciliacion_Tarjetas_Anual_Ene-Oct2026.xlsx`.

**Ventas con cadena completa** (incluye las de contracargo, que sí llegaron al banco):

| Mes | Niubiz | Amex | Diners |
|---|---|---|---|
| Enero | 305 / 305 | 15 / 15 | — |
| Febrero | 263 / 265 | 10 / 10 | 0 / 5 (*) |
| Marzo | 284 / 284 | 14 / 14 | 5 / 5 |
| Abril | 253 / 253 | 12 / 12 | 3 / 3 |
| Mayo | 302 / 302 | 11 / 11 | — |
| Junio | 285 / 285 | 19 / 19 | 2 / 2 |
| Julio | 280 / 280 | 16 / 16 | 5 / 5 |
| Agosto | 272 / 312 | 11 / 11 | — |
| Septiembre | 288 / 289 | 14 / 14 | 1 / 2 |
| Octubre (al 20/10) | 155 / 186 | 6 / 6 | 0 / 1 |

(*) `Depositos Diners` empieza el 05/03/2026: las 5 ventas Diners de febrero no tienen
depósito en el archivo, y sus abonos (BCP USD 188.01 y 195.30 del 16/02, BCP PEN 320.23
del 26/02, BBVA PEN 1,481.07 del 02/03) quedan "sin depósito". Es falta de datos del
operador, no un faltante.

**Abonos de tarjeta del banco conciliados:** enero 83/83, febrero 101/104, marzo 123/124,
abril 100/100, mayo 101/101, junio 109/109, julio 105/105, agosto 106/112 (los 6 restantes
son los revertidos por la regularización), septiembre 112/113, octubre 66/67.

**Totales por mes** (hoja "Tarjetas - Totales por Mes"): la diferencia sin explicar entre
depósitos conciliados y abonos conciliados es **0 en todos los meses, de enero a octubre**.
En el total del año la diferencia es **S/ 29.50 y US$ 14.49**, solo cargos fijos de Diners
(el desfase entre meses se compensa). En agosto y septiembre aparecen −/+16,038.05 USD
como depósitos no conciliados: es la fila informativa "Saldo deudor al 30/08" y su
contrapartida.

**Pendientes a revisar del año** (2,912 ventas):

| Qué | Casos | Explicación |
|---|---|---|
| Venta sin depósito y sin explicación | 5 | Las ventas Diners de febrero (archivo de depósitos Diners desde marzo). |
| Venta sin depósito por corte | 19 | Ventas de los últimos días: el depósito aún no llega. |
| Falta en Depositos amex | 0 | Antes 39: no faltaban filas; `Depositos amex` suma las ventas del día, y en junio Amex pagó juntas ventas retenidas (ver "Reglas del cruce"). |
| Venta depositada sin abono | 57 | Agosto ~40 (regularización del 29/08); octubre: el extracto BBVA llega hasta el 16/10. |
| Día sin abono y sin explicación | 0 | — |
| Abono revertido por el operador | 6 (S/ 2,185.32 · US$ 21,654.09) | Los pagos duplicados del 27–28/08 que revirtió la regularización del 29/08. |
| Abono sin depósito | 6 (S/ 2,077.21 · US$ 488.61) | 4 Diners de febrero–marzo (ver *), Diners BCP del 17/09 y Amex BBVA del 01/10. |
| Contracargos | 2 | Los de junio y agosto (ejemplos 2 y 3). |

**Cifras anteriores (archivo `Ventas y depositos_tarjetas.xlsx`, por mes):**
junio 288 / 306 (94.1%), julio 283 / 301 (94.0%), agosto 274 / 323 (84.8%),
septiembre 257 / 305 (84.3%). Antes de corregir el motor (2026), julio daba
158 / 301 (52.5%) y Pago Link 0 / 25 grupos.

## Referencia técnica y conocimiento acumulado

Lo aprendido analizando datos reales de EMPRESA DEMO SCRL (RUC 77127684268), enero–
octubre 2026. Los números pueden cambiar con nuevos exports.

### Fuentes de datos y sus formatos

| Fuente | Qué es | Particularidades |
|---|---|---|
| `Ventas y depositos.xlsx` (17/08) | Versión antigua. | **Export incompleto**: le faltaban los depósitos del 01/08. |
| `bancos/Ventas y depositos_tarjetas.xlsx` (12/10) | `Depositos` hasta 01/10. | `Fecha de depósito` como serial de Excel (46192 = 19/06/2026). |
| `Conciliacion octubre q1/Tarjeta_de_credito.xlsx` (20/10) | Vigente: ventas y depósitos hasta el 20/10, más hojas contables propias. | Algunas hojas tienen espacios al final del nombre (`Ventas amex `): se comparan sin espacios. |
| Hoja `Ventas` | Una fila por transacción Niubiz/Amex/Diners. | `MARCA TARJETA`: VI/MC = Niubiz, AM = Amex, DN = Diners. Anulaciones (monto negativo o "Anulada") se descartan. |
| Hoja `Depositos` (Niubiz) | Una fila por transacción **y por cada cargo o ajuste**. | `Tipo de Operación`: `Ventas`, `Costo por trx liquidada`, `Uso del Sist CE otras marcas`, `Mantenimiento Pago Link`, `Cobro Mensualidad POS`, `Traslado de Deuda`, `Regularización gestión operativa`, `Cargos por Contracargos`, `Procesamiento de Disputas`, `Saldo deudor al dd/mm`. En ventas: neto = bruto + comisión + IGV (comisión e IGV negativos). `Es DCC` / `Monto DCC`: conversión de moneda al pagar. |
| Hoja `Depositos amex` | **Resumen diario por establecimiento**, sin ID de operación. | `DEUDA ANTERIOR`: descuento fijo US$ 2.01 / S/ 5.90. En el archivo de octubre trae **pegada debajo (desde la fila 86) otra tabla** con otro orden de columnas (detalle por venta): esas filas se descartan porque su "moneda" no es PEN/USD; antes se leían como depósitos duplicados. |
| Hoja `Ventas amex` | Detalle Amex con `Código de Autorización` y `Fecha de Depósito`. | Da la fecha real aunque falte la fila en `Depositos amex`. |
| Hoja `Depositos Diners` | 1 fila = 1 transacción. | Sin banco: se infiere del código (`8000*` = BCP, `1751163*` = BBVA). Comisión porcentual e IGV ya descontados; **el cargo fijo por pago (S/ 5.90 / US$ 2.07) no viene en el archivo.** |
| Extractos BBVA | `.xls` que es HTML paginado. | `mergeHtmlPaginatedSheets` junta las páginas. Formato nuevo: `F.Operac. / F.Valor / Referencia / Importe / ITF / Num.Mvto`. |
| Extractos BCP | `Fecha, Descripción operación, Monto, Operación - Número`. | Algunos libros guardan el N° con ceros a la izquierda (`67551313`) y otros sin ellos. |
| Libros `bancos/B_Bancos_04…09.xlsx` | Extracto pegado sin espacios en los encabezados, repetido en bloques junto a columnas contables (`Glosa`, `IMPORTE`, `OBS`), más hojas sueltas (`EFECTIVO`, `OBS`, `Anticipo - No id`, `Hoja1`, `TD`). | `B_Bancos_05/06/07` repiten hojas de abril. Ver `extractosDesdeLibro`. |
| `bancos/Bancos enero_marzo_2026.xlsx` | Estados de cuenta BBVA y BCP de enero a marzo **convertidos desde PDF**: una hoja por banco, mes y moneda (`EECC_BBVA_2026-01_MN`, `EECC_BCP_2026-03_USD`; `MN` = soles) y una hoja `Consolidado` que repite todo (se ignora porque su nombre no identifica un banco). | Fechas **sin año** (`08-01` = 08/01; el año sale del nombre del archivo). BBVA: `OPER.` (fecha de operación), `VALOR` (fecha valor, **no** importe), `DESCRIPCIÓN SALDO ANTERIOR` (glosa), `N° OPER.`, `CARGO/ABONO`. BCP: `FECHA PROC.`, `DESCRIPCION`, `NUM OP`, `CARGO / ABONO`. |
| `Bancos_Q1_20102026.xlsx` | Consolidado de octubre: una hoja por cuenta (`BCP SOLES`, `SCT USD`, `IBK SOLES`, `BBVA USD`, `YAPE`…). | BBVA hasta el 16/10; BCP hasta el 18/10. |

### Códigos de comercio y glosas

| Código | Banco | Moneda | Canal | Glosa del abono |
|---|---|---|---|---|
| 209281590 | BBVA | PEN | POS | `ABONO VISANET 5656321223` |
| 952538885 | BBVA | USD | POS | `ABONO VISANET 4079244026` |
| 315472803 | BCP | PEN | POS | `DE COMPANIA PERUANA DE` (sin código) |
| 966661760 | BCP | USD | POS | `DE COMPANIA PERUANA DE` (sin código) |
| 952808414 | BBVA | PEN | Pago Link | `R20341198217COMPA#IA PERUANA DE MEDIOS D` (sin código) |
| 811777741 | BBVA | USD | Pago Link | `R20341198217COMPA#IA PERUANA DE MEDIOS D` (sin código) |

`239624595` (BBVA PEN) y `413721590` (BBVA USD) son **códigos POS que se usaron en
marzo de 2026** junto con 209281590 / 952538885, con su propia glosa (`ABONO VISANET
9599528907`); el resto del año aparecen solo como contrapartida de traslados de
deuda. Amex en BBVA: `R20467737300COMPA#IA DE SERVICIOS CONEXO`; Amex en BCP:
`NNNNNNNNNNNNNNN COMPAN` o `DE COMPANIA PERUANA DE`. Diners:
`R20100118760DINERS CLUB PERU SA` (BBVA) o `DE DINERS CLUB PERU S.` (BCP).
**BBVA escribe `COMPA#IA` con `#`**: los patrones usan `COMPA[ÑN#]IA`.

### Hallazgos y decisiones

- **El 52.5% inicial no era un problema de fechas.** Ampliar la ventana de 2 a
  20 días no cambiaba nada. Causas: el abono es la suma de todo lo del día por
  código (no solo las ventas del periodo), y BBVA escribe `COMPA#IA`.
- **Las "24 ventas Niubiz sin depósito" de julio eran un export incompleto** de
  `Depositos`. No se reclama a Niubiz; conviene volver a descargar cuando
  aparezcan "sin explicación" de fin de mes.
- **Cargo fijo de Diners** (abril–octubre): 7 pagos con 1 cargo y 1 con 2; el
  archivo no lo trae.
- **Días pagados en 2 abonos** (Amex): 6 casos, mismo día, N° casi
  consecutivos.
- **Amex tardío:** 311.34 (01→04/05) y 258.60 (25→30/07).
- **Diners en otro banco:** 20/04, BCP según el operador, abonado en BBVA.
- **Duplicados con ceros a la izquierda** inflaban 26 abonos de BCP USD de
  abril como "sin depósito" (corregido).

### Limitaciones y pendientes

- **No se probó en un navegador:** todo se verificó con Node sobre los archivos
  reales.
- Un día pagado en **3 o más** abonos no se reconoce (solo 2).
- `Asientos por Periodo` va en moneda original, sin tipo de cambio.
- Pendiente con el operador: los 136.57 del contracargo 423194875749118 (DCC),
  el motivo del pago adelantado por el bruto del 27/08 (y de su regularización)
  y la cuenta de pago de Diners.

### Cómo verificar cambios sin navegador

El entorno usado no puede abrir `file://`; se prueba con Node:

1. `npm install xlsx jsdom@22 --no-save` en la carpeta del proyecto (**las dos
   juntas**: instalar una por separado borra la otra; `jsdom@25+` falla en Node
   18).
2. Un script lee el último bloque `<script>` del HTML y lo ejecuta con
   `vm.runInContext` en un contexto con `document` simulado (`getElementById`
   devuelve objetos con `value`/`style`/`classList`), `XLSX` y `DOMParser` de
   jsdom. Al final del código se agrega
   `globalThis.__x = { extractosDesdeLibro, ejecutarConciliacion, exportar, … }`
   para llamar a las funciones.
3. Se cargan `operadorFile` y `extractosFiles` como la página, se fija
   `periodoMes` (`ALL` o `AAAA-MM`), se llama a `ejecutarConciliacion()` y a
   `exportar()` (ambas usan `setTimeout`: esperar) interceptando
   `XLSX.writeFile` para leer el libro.
4. Antes de hacer commit: `node --check` de cada bloque `<script>`, comparar con
   las cifras de referencia y borrar los scripts temporales.

Ojo con el shell: `\n` y `\d` dentro de heredocs o `node -e` se pierden. Los
scripts de parcheo se escriben con un editor de archivos, y los archivos HTML
usan CRLF (normalizar a LF al editar y restaurar al guardar).

### Mapa de funciones

- **Lectura:** `leerVentas`, `leerDepositosNiubiz(todos)`, `leerDepositosAmex`,
  `leerDepositosDiners`, `leerAbonosBancarios`, `extractosDesdeLibro`,
  `recortarPrimerBloque`, `primeraColumna`, `columnaNroOperacion`,
  `findGlosaColumnIndex`, `parseHTMLStatement`, `mergeHtmlPaginatedSheets`.
- **Cruce:** `cruzarVentasDepositos`, `buildFechaDepositoAmexPorAuth`,
  `clasificarVentaSinDeposito`, `construirGruposBanco`, `agruparDepositos`,
  `cruzarDepositosBanco` (5 pasadas), `cargosFijosDiners`, `familiaDeGlosa`,
  `glosaContieneCodigo`, `estadosDeGrupos`, `ejecutarConciliacion`.
- **Reportes:** `construirResumenTarjetas`, `construirTotalesPorMes`, `construirTrazabilidad`,
  `construirDiasVsBanco`, `construirAjustesOperador(filtro)`,
  `construirDiasSinAbono`, `construirDepositosConBanco`, `construirAsientos`,
  `leerLiquidaciones`, `construirCierre`, `calcularRangoReporte`,
  `enPeriodoReporte`, `coberturaExtractos`, `hayExtractoCerca`,
  `categoriaDeposito`, `exportar`.
- **Constantes:** `CARD_GLOSA_PATTERNS`, `CARGO_FIJO_DINERS`,
  `CATEGORIAS_AJUSTE`, `COLUMNAS_COMPOSICION`, `COLUMNAS_TRAZA_BANCO`,
  `COLUMNAS_TRAZABILIDAD`, `IGV_TASA`.

**Código compartido con `portable_2.5.html`:** el 2.5 tiene una copia de los
lectores y del cruce en su módulo `TJ`. Si cambias `cruzarDepositosBanco`,
`estadosDeGrupos`, `cargosFijosDiners`, `construirAjustesOperador` o los
lectores aquí, haz el mismo cambio en el `TJ` del 2.5.

**El Excel exportado lo usa la opción 14 del repositorio contable**
(`4_Repositorio_FI/src/tarjetas_mayor.py`), que cruza la conciliación
contra el mayor de las cuentas 1031/1032. Lee las hojas `Tarjetas -
Trazabilidad`, `BBVA - Tarjetas`, `BCP - Tarjetas`, `Tarjetas - Totales por
Mes` y `Tarjetas - Días vs Banco`, por nombre de hoja y de columna (`Fecha
venta`, `Importe venta (bruto)`, `Fecha abono`, `N° operación abono`,
`Estado de la cadena`, `N° Operación`, `(Comisiones e IGV…`, `Estado`…).
Si cambias esos nombres, actualiza también ese script.

### Entregables generados (no versionados)

`Conciliacion_Tarjetas_Historico_Abr-Oct2026.xlsx`, `Trazabilidad_Abono_59690.xlsx`,
`Trazabilidad_423194875749118.xlsx`, `Trazabilidad_919093786579754.xlsx`,
`Informe_Trazabilidad_811777741_Jun2026.xlsx`,
`Contracargos_y_Ajustes_Niubiz.xlsx`, `Reporte_Depositos_con_Banco.xlsx`,
`Pendientes_Septiembre_Niubiz.xlsx`.

## Historial de cambios

Detalle completo en `git log` (rama `feature/v24-tarjetas-notas-credito`).

| Cuándo | Cambio |
|---|---|
| Ago–sep 2026 | Grupos del día con **todas** las filas de `Depositos`; patrón `COMPA[ÑN#]IA`; lectura de libros de trabajo; N° de operación por encabezado; julio de 52.5% a 94.0%. |
| Sep–oct 2026 | Hojas de contracargos y ajustes, días sin abono, regularizaciones vs abonos, depósitos + banco; "Todos los periodos"; asientos y cierre. |
| Oct 2026 | Excel con la estructura del 2.5: Resumen Tarjetas, Trazabilidad, Días vs Banco, Ajustes; Depositos + Banco ocultas; periodo por meses de los extractos. |
| Oct 2026 | Motor: día pagado en 2 abonos; cargo fijo de Diners; pasada 4 (Amex hasta 5 días, Diners en el otro banco) con comentario; duplicados sin ceros a la izquierda. |
| Oct 2026 | Pasada 5: abonos revertidos por una regularización (pagos duplicados por el bruto del 27–28/08); columna `Tipo de Abono` leída. |
| Oct 2026 | Hoja "Tarjetas - Totales por Mes" (conciliado / no conciliado); `Depositos amex` ignora la tabla pegada debajo con otro formato. |
| Oct 2026 | Amex: ventas enlazadas a la fila diaria de `Depositos amex` por suma, y ventas retenidas pagadas juntas (por establecimiento, hasta 60 días); "Falta en Depositos amex" pasa de 39 a 0. |
| Oct 2026 | Enero–marzo: estados de cuenta convertidos desde PDF (`OPER.`, `NUM OP`, `CARGO / ABONO`, `MN`/`ME`, fechas sin año); grupos del día separados por **código de comercio** (dos códigos POS en la misma cuenta en marzo). |
| Nov 2026 | Cadena Completa y Trazabilidad con comisión, regla del cruce, validación ("Válido" / "Válido con nota") y qué corregir (redondeo, operador). |
