# Costos Tabmex — Nómina por proyecto

Control de **gastos de nómina por proyecto** (sueldo + bonos + otras prestaciones)
en Excel, con una macro que carga cada registro a la hoja del proyecto
correspondiente y consolida los totales automáticamente.

## Archivos

| Archivo | Qué es |
|---|---|
| `Nomina_por_proyecto.xlsm` | El libro de Excel (macro-habilitado), con un botón **CARGAR GASTO** en la hoja Captura y la hoja **Pagos fijos**. |
| `CargarGasto.bas` | El módulo de la macro `CargarGasto` (ya incluida en el libro). |
| `patch_pagos_y_boton.py` | Script que aplicó la última corrección: arregló el botón y agregó la hoja **Pagos fijos**. No necesitas tocarlo. |
| `build_nomina.py` | Script original que generó la primera versión del `.xlsm`. Quedó desfasado respecto al libro actual (no incluye Quincena/Puesto ni la hoja Pagos fijos). |
| `finalize.py` | Ajuste original del `.xlsm` (content-type macro y botón). No necesitas tocarlo. |

## Novedades (jun-2026)

- **Botón CARGAR GASTO arreglado.** En la hoja Captura se habían acumulado
  4 botones encimados y el que quedaba arriba no tenía macro asignada, por eso
  el clic no hacía nada. Ahora hay **un solo botón**, con la macro `CargarGasto`.
- **Nueva hoja `Pagos fijos`** (cronograma de pagos fijos): renta, servicios,
  impuestos y demás, con **fecha de pago** y asignación a un **proyecto**.

## Cómo descargarlo

1. Arriba, en la lista de archivos de GitHub, haz clic en `Nomina_por_proyecto.xlsm`.
2. Pulsa el botón **Download raw file** (icono de descarga ⤓, arriba a la derecha).
3. Repite con `CargarGasto.bas`.

Ambos caen en tu carpeta **Descargas / Downloads**.

## Puesta en marcha (solo la primera vez)

1. Abre `Nomina_por_proyecto.xlsm`. Si sale la barra amarilla de seguridad,
   pulsa **Habilitar contenido**.
2. Importa la macro: `Alt+F11` → **Archivo ▸ Importar archivo…** → elige
   `CargarGasto.bas` → cierra el editor.
3. Guarda el libro **conservando el formato `.xlsm`** (si Excel sugiere `.xlsx`,
   dile que no: ese formato borra las macros).

## Uso diario

1. Ve a la hoja **Captura** y llena: Fecha, Proyecto, Empleado, Concepto,
   Sueldo, Bonos y Otras prestaciones. La celda `C16` muestra el total como
   vista previa (no es un botón).
2. Pulsa el botón azul **CARGAR GASTO** de la hoja Captura (o, si prefieres,
   `Alt+F8` → **CargarGasto** → **Ejecutar**). El botón ya trae la macro
   asignada; solo funciona después de haber importado `CargarGasto.bas` la
   primera vez.
3. El registro se manda a la hoja del proyecto (la crea si no existía) y al
   libro maestro `_Movimientos`. El **Resumen** y los totales por proyecto se
   actualizan solos.

## Hojas del libro

- **Inicio** — guía rápida.
- **Captura** — formulario de captura (con botón **CARGAR GASTO**).
- **Resumen** — consolidado por proyecto (fórmulas `SUMIFS`). Columnas:
  Total percibido · Carga social · **Pagos fijos (pagados)** · Costo total.
- **Pagos fijos** — cronograma de pagos fijos (ver abajo).
- **Catálogos** — Proyectos, Empleados, Conceptos, Puestos y el **Factor de
  carga social patronal (%)**.
- **Tabmex / Contrato voceo / Producción** — una hoja por proyecto.
- **_Movimientos** (oculta) — libro maestro / auditoría.
- **_Plantilla** (oculta) — base para crear hojas de proyecto nuevas.

## Hoja "Pagos fijos" (cronograma)

Registra aquí los gastos fijos recurrentes (renta, luz, agua, internet,
impuestos, mantenimiento…). Una fila por pago, con estas columnas:

| Columna | Para qué |
|---|---|
| **Fecha de pago** | Cuándo se paga / venció. |
| **Concepto** | Texto libre (p. ej. "Renta oficina", "Luz CFE"). |
| **Categoría** | Lista: Renta, Servicio, Impuesto, Predial, Mantenimiento, Otro. |
| **Proyecto** | Lista de proyectos. **A qué proyecto se carga el costo.** |
| **Periodicidad** | Único, Mensual, Bimestral, Trimestral, Semestral, Anual. |
| **Monto** | Importe del pago. |
| **Estatus** | **Pendiente / Pagado.** |
| **Notas** | Observaciones. |

Cómo se cargan a los proyectos: en cuanto marcas un pago como **`Pagado`**, su
monto se suma al proyecto elegido en la columna **Pagos fijos (pagados)** del
**Resumen**, y entra en el **Costo total** de ese proyecto. Los pagos en
**`Pendiente`** no afectan el costo todavía (solo aparecen en el cronograma y
en el "Total programado" de la propia hoja). Todo es por **fórmulas**: no hay
que pulsar ningún botón ni importar macros nuevas.

## Factor de carga social patronal

En **Catálogos** hay un parámetro `Factor de carga social patronal (%)`,
en **0 %** por defecto. Si lo subes a tu estimado de IMSS, Infonavit y
provisiones (aguinaldo, prima vacacional), la columna *Costo c/carga social* de
cada proyecto y el *Costo total* del Resumen reflejan el **costo patronal real**,
no solo lo pagado en mano. Para costeo de proyecto eso suele ser lo que importa.
