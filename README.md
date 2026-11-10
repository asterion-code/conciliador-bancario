# Conciliación Bancaria — Empresa Demo

Dos herramientas **portables de un solo archivo HTML** para conciliar los
registros de la empresa contra los extractos bancarios (BCP, Scotiabank,
BBVA/Continental, Interbank, Yape). Sin instalación, sin servidor, sin enviar
datos a ningún tercero: todo el procesamiento ocurre en el navegador.

| Herramienta | Para qué | README |
|---|---|---|
| **`portable_2.5.html`** | Conciliador general: Registro de Cobranzas (depósitos directos y pagos con tarjeta) vs. extractos bancarios, con estado contable para la revisión quincenal. Versión activa. | [README_portable_2.5.md](README_portable_2.5.md) |
| **`portable_tarjetas.html`** | Conciliador independiente, solo tarjetas: Ventas → Depósitos del operador → Banco, sin pasar por Cobranzas. Incluye asientos contables y cierre de la cuenta por cobrar al operador. | [README_portable_tarjetas.md](README_portable_tarjetas.md) |

Las dos usan **el mismo motor de depósitos de tarjeta ↔ banco** (en el 2.5 vive
copiado en el módulo `TJ`) y exportan **las mismas hojas de tarjetas**: Resumen
Tarjetas, Tarjetas - Trazabilidad, Tarjetas - Días vs Banco y Tarjetas - Ajustes,
más las hojas "Depositos + Banco" ocultas.

## ¿Cuál uso?

- Para conciliar el mes (o la quincena) y saber qué falta registrar o
  contabilizar → **`portable_2.5.html`**.
- Para saber si un descuadre de tarjetas es real (depósito que no llegó,
  ajuste del operador) o solo un error de tipeo en Cobranzas, para analizar
  varios meses de tarjetas a la vez, o para los asientos de tarjetas →
  **`portable_tarjetas.html`**.

## Estructura del repositorio

| Archivo/carpeta | Contenido |
|---|---|
| `portable_2.5.html` | Conciliador general. Versión activa. |
| `portable_tarjetas.html` | Conciliador independiente de tarjetas. |
| `portable_2.4.html`, `portable_2.3.html`, `portable_2.2.html` | Versiones anteriores del general, solo como referencia histórica. |
| `README_portable_2.5.md`, `README_portable_tarjetas.md` | Documentación detallada de cada herramienta. |
| `README_portable_2.4.md`, `DIAGNOSTICO_Y_PLAN_v2.4.md` | Documentación histórica de la versión 2.4. |
| `bancos/` | Estados de cuenta de enero a marzo (`Bancos enero_marzo_2026.xlsx`), libros de trabajo `B_Bancos_04…09.xlsx`, `Ventas y depositos_tarjetas.xlsx` (no versionados). |
| `Conciliacion octubre q1/` | Juego completo de octubre (primera quincena): Cobranzas, `Bancos_Q1_20102026.xlsx`, `Tarjeta_de_credito.xlsx` y los reportes generados. |
| `conciliacion parcial agosto/` | Datos parciales de agosto para el conciliador general. |

Los `*.xlsx` están en `.gitignore`: viven solo en la carpeta local.

## Por dónde seguir (para no perder el hilo)

- **Archivo de operadores vigente:** `Conciliacion octubre q1/Tarjeta_de_credito.xlsx`
  (depósitos hasta el 20/10/2026).
- **Último análisis de tarjetas (año completo):** `Conciliacion_Tarjetas_Anual_Ene-Oct2026.xlsx`
  y la sección "Resultados de referencia" de `README_portable_tarjetas.md`.
- **Cómo leer una trazabilidad paso a paso:** `README_portable_tarjetas.md`,
  "Cómo leer una trazabilidad" (IDs `423194875749118` y `919093786579754`).
- **Pendiente principal:** probar las dos herramientas en un navegador con los
  archivos de `Conciliacion octubre q1/` (todo se verificó con Node).
- **Preguntas abiertas para los operadores:** los 136.57 del contracargo
  423194875749118 (probable efecto del DCC), el motivo del pago duplicado por el
  bruto del 27/08/2026 que revirtió la regularización del 29/08 (ver
  `Trazabilidad_Abono_59690.xlsx`) y la cuenta de pago de Diners del 20/04.
- **Al cambiar el motor de tarjetas:** hacer el mismo cambio en
  `portable_tarjetas.html` y en el módulo `TJ` de `portable_2.5.html`.
- **Rama de trabajo:** `feature/v24-tarjetas-notas-credito` (sin remoto
  configurado).
