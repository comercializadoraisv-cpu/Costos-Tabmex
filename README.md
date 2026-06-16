# Costos Tabmex — Nómina por proyecto

Control de **gastos de nómina por proyecto** (sueldo + bonos + otras prestaciones)
en Excel, con una macro que carga cada registro a la hoja del proyecto
correspondiente y consolida los totales automáticamente.

## Archivos

| Archivo | Qué es |
|---|---|
| `Nomina_por_proyecto.xlsm` | El libro de Excel macro-habilitado, **con la macro `CargarGasto` ya incrustada** y un botón **CARGAR GASTO** en la hoja Captura. Ya no hay que importar nada. |
| `CargarGasto.bas` | Copia legible del código de la macro (referencia / respaldo). El libro ya lo trae adentro. |
| `build_nomina.py` | Genera el `.xlsm` con openpyxl (no necesitas tocarlo). |
| `build_vba.py` | Construye el proyecto VBA (`vbaProject.bin`) desde cero según MS-OVBA (no necesitas tocarlo). |
| `finalize.py` | Incrusta el `vbaProject.bin`, fija el content-type macro y agrega el botón (no necesitas tocarlo). |

## Cómo descargarlo

1. Arriba, en la lista de archivos de GitHub, haz clic en `Nomina_por_proyecto.xlsm`.
2. Pulsa el botón **Download raw file** (icono de descarga ⤓, arriba a la derecha).

Cae en tu carpeta **Descargas / Downloads**.

## Puesta en marcha (solo la primera vez)

La macro **ya viene incrustada y activa**; solo tienes que permitir que Excel
la ejecute:

1. Abre `Nomina_por_proyecto.xlsm`. Si sale la barra amarilla de seguridad,
   pulsa **Habilitar contenido** (Enable Content). Eso es todo.
2. Si descargaste el archivo de internet y Excel lo abre en **Vista protegida**
   o muestra *“Las macros se han deshabilitado”* sin botón para habilitarlas:
   cierra Excel, haz clic derecho sobre el archivo → **Propiedades** → marca
   **Desbloquear** (Unblock) abajo → **Aceptar**, y vuelve a abrirlo.
3. **Ya no necesitas importar `CargarGasto.bas` ni entrar al editor (`Alt+F11`).**

## Uso diario

1. Ve a la hoja **Captura** y llena: Fecha, Proyecto, Empleado, Concepto,
   Sueldo, Bonos y Otras prestaciones. La celda `C16` muestra el total como
   vista previa (no es un botón).
2. Pulsa el botón azul **CARGAR GASTO** de la hoja Captura (o, si prefieres,
   `Alt+F8` → **CargarGasto** → **Ejecutar**). El botón ya trae la macro
   asignada y la macro ya está dentro del libro: funciona desde la primera vez.
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
