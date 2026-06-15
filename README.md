# Costos Tabmex — Nómina por proyecto

Control de **gastos de nómina por proyecto** (sueldo + bonos + otras prestaciones)
en Excel, con una macro que carga cada registro a la hoja del proyecto
correspondiente y consolida los totales automáticamente.

## Archivos

| Archivo | Qué es |
|---|---|
| `Nomina_por_proyecto.xlsm` | El libro de Excel (macro-habilitado), con un botón **CARGAR GASTO** en la hoja Captura. |
| `CargarGasto.bas` | El módulo de la macro `CargarGasto`. Se importa una sola vez. |
| `build_nomina.py` | Script que genera el `.xlsm` (no necesitas tocarlo). |
| `finalize.py` | Ajusta el `.xlsm` (content-type macro y botón). No necesitas tocarlo. |

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
- **Captura** — formulario de captura.
- **Resumen** — consolidado por proyecto (fórmulas `SUMIFS`).
- **Catálogos** — Proyectos, Empleados, Conceptos y el **Factor de carga social
  patronal (%)**.
- **Tabmex / Contrato voceo / Producción** — una hoja por proyecto.
- **_Movimientos** (oculta) — libro maestro / auditoría.
- **_Plantilla** (oculta) — base para crear hojas de proyecto nuevas.

## Factor de carga social patronal

En **Catálogos** hay un parámetro `Factor de carga social patronal (%)`,
en **0 %** por defecto. Si lo subes a tu estimado de IMSS, Infonavit y
provisiones (aguinaldo, prima vacacional), la columna *Costo c/carga social* de
cada proyecto y el *Costo total* del Resumen reflejan el **costo patronal real**,
no solo lo pagado en mano. Para costeo de proyecto eso suele ser lo que importa.
