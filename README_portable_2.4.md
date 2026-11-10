# `portable_2.4.html` — Conciliador General (Cobranzas + Tarjetas)

> **Superseded por [`portable_2.5.html`](README_portable_2.5.md)** — se
> mantiene solo como referencia histórica. Usa la v2.5 para trabajo activo.

Herramienta **portable de un solo archivo HTML** para conciliar el
**Registro de Cobranzas** de la empresa contra los extractos bancarios (BCP,
Scotiabank, Continental/BBVA, Interbank), sin instalación, sin servidor y sin
enviar datos a ningún tercero: todo el procesamiento ocurre en el navegador
del usuario.

## Objetivo

Cada mes, el área de Cobranzas registra los depósitos bancarios recibidos de
clientes en un Excel ("Registro de Cobranzas"), y por separado se descargan
los extractos de movimientos de cada banco. Verificar a mano que cada
depósito registrado realmente aparezca en el banco —con el mismo importe,
fecha y número de operación— es lento y propenso a error humano cuando se
manejan cientos o miles de filas por mes, en dos monedas (Soles y Dólares) y
hasta cinco bancos distintos.

Además, este conciliador incorpora el cruce de **pagos con tarjeta de
crédito** (T.Crédito) registrados en Cobranzas contra las ventas de los
operadores (Niubiz, Amex, Diners) — ver `DIAGNOSTICO_Y_PLAN_v2.4.md` para el
detalle de esa fase.

## Qué hace

1. **Carga** el Registro de Cobranzas (`.xlsx`) y uno o varios extractos
   bancarios:
   - Archivos individuales por banco/moneda (`.xls`/`.xlsx`, incluyendo los
     que en realidad son HTML disfrazado de `.xls`, como exporta BBVA).
   - **Un único Excel consolidado con varias hojas** (una por cuenta/banco,
     p.ej. `BANCOS.xlsx` con hojas `SCOTIABANK SOLES`, `BBVA DOLARES`, `IBK
     SOLES`, `BCP DOLARES`, etc.) — el motor detecta automáticamente qué
     hojas traen tabla de movimientos y cuáles son solo metadata/resumen
     (saldo, tipo de cambio), y trata cada hoja válida como una cuenta
     independiente.
2. **Filtra** solo las filas que corresponden a depósitos bancarios reales
   (`Dep. bancario` en "Cuenta o Forma de Cobranza") — así no compite contra
   pagos con tarjeta, notas de crédito, efectivo, etc., que nunca aparecerán
   en un extracto.
3. **Concilia** cada cobranza con su movimiento bancario, usando el **Número
   de Operación** como identificador único (ancla), y usando banco, fecha,
   importe y moneda como criterios de calidad/desempate — no como filtros que
   descarten un match legítimo.
4. **Corrige automáticamente** errores conocidos de origen: monedas mal
   etiquetadas cuando el Tipo de Cambio delata que en realidad es la moneda
   contraria.
5. **Cruza tarjetas de crédito**: enlaza los pagos con tarjeta de Cobranzas
   contra las ventas de los operadores y, de ahí, contra su depósito y abono
   bancario real (reconociendo la glosa del extracto por patrones conocidos:
   `ABONO VISANET`, `COMPAÑIA PERUANA DE MEDIOS...`, `DINERS CLUB`, etc.).
6. **Alerta, no oculta**: cuando algo casi coincide pero no exactamente (fecha
   con unos días de desfase, banco distinto al esperado, moneda que no
   cuadra), lo muestra como un registro conciliado con una alerta visible y
   el detalle de la diferencia — para que el usuario decida si es un error de
   tipeo corregible o una discrepancia real.
7. **Reporta** el resultado en pantalla (dashboard + tablas navegables) y en
   un Excel descargable, listo para archivar o enviar.

## Qué NO hace (por diseño)

- No modifica ni sube tus archivos a ningún servidor — es una página HTML
  estática que corre 100% en tu navegador.
- No reemplaza el criterio contable: las alertas están para que una persona
  las revise, no para conciliar "a la fuerza" cualquier coincidencia.

## Cómo usarlo

1. Abre `portable_2.4.html` con cualquier navegador moderno (no requiere
   internet, salvo la primera carga de fuentes si están disponibles).
2. Arrastra o selecciona el archivo de Cobranzas y los extractos bancarios
   (sueltos o el Excel consolidado multi-hoja).
3. Revisa el mapeo de columnas auto-detectado (ajústalo si hace falta) en el
   panel de cada extracto.
4. Define tolerancias de fecha/redondeo y qué bancos incluir.
5. Ejecuta la conciliación y revisa las pestañas de resultados:
   - **Conciliados**: pares cruzados, con su alerta si aplica.
   - **Pendiente en Cobranzas** / **Pendiente en Bancos**: lo que no encontró
     contraparte.
   - **Estado de Cobranzas**: vista resumida solo del lado de Cobranzas
     (Forma de Cobranza, Banco, Nro/Doc, Monto), deduplicada, con estado
     Conciliado / No Conciliado / Parcial.
6. Exporta el reporte a Excel para archivo o envío.

## Archivos de entrada

| Archivo | Contenido |
|---|---|
| `Registro de Cobranzas *.xlsx` | Cobranzas del periodo — depósitos bancarios directos y pagos con tarjeta. |
| `bancos/*.xls`/`.xlsx` o un Excel consolidado (p.ej. `BANCOS.xlsx`) | Extractos bancarios del periodo, una o varias cuentas. |
| `Ventas y depositos.xlsx` | Ventas y depósitos de los operadores de tarjeta (Niubiz/Amex/Diners), para el cruce de T.Crédito. |

## Diferencia con `portable_tarjetas.html`

Este conciliador (`portable_2.4.html`) es el más completo: cubre Cobranzas
(depósitos directos + tarjetas) y produce el reporte consolidado final. El
cruce de tarjetas que hace aquí depende del Registro de Cobranzas, donde
pueden existir errores de tipeo en marca/número de documento.

`portable_tarjetas.html` es una herramienta **aparte**, más chica y enfocada
solo en validar Ventas → Depósitos → Banco para tarjetas, sin pasar por
Cobranzas — ver `README_portable_tarjetas.md`.

## Estado del proyecto

Activo. Ver `DIAGNOSTICO_Y_PLAN_v2.4.md` para el estado real de avance por
fase, y `git log` para el historial detallado de cambios.
