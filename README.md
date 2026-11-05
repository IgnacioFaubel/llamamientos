# Importación mensual de llamamientos

Aplicación de escritorio para sustituir de una vez los datos del mes cerrado en la Herramienta Llamamientos.

## Ejecución

Abre `LlamamientosMensuales.exe`. No necesita instalar Python ni ejecutar archivos `.bat`. El equipo sí debe tener Microsoft Excel de escritorio.

## Uso

1. Abre el Excel mensual de llamamientos en Excel, deja que recalcule y guárdalo después de aplicar todas las correcciones.
2. Cierra todas las ventanas de Microsoft Excel. La aplicación bloquea la importación mientras exista un proceso de Excel para impedir copias de conflicto de OneDrive.
3. Abre `LlamamientosMensuales.exe` y selecciona:
   - El Excel mensual de origen (`.xlsm` o `.xlsx`).
   - La Herramienta Llamamientos de destino (`.xlsx`).
4. Pulsa **Analizar archivos** y revisa el total diario, los avisos y los errores.
5. Si no hay errores, pulsa **Importar y reemplazar el mes**.

La aplicación recuerda las últimas rutas utilizadas, por lo que funciona con carpetas de SharePoint sincronizadas mediante OneDrive.

## Reglas aplicadas

- Lee las hojas diarias cuyo nombre tenga el formato `01_10`, `02_10`, etc.
- Toma los datos desde la fila 10.
- Solo importa filas con un estado informado en G.
- `ND` es un estado válido y se conserva.
- Los errores de Excel (`#N/D`, `#REF!`, etc.) bloquean la importación.
- La fecha se obtiene del nombre de la hoja; E9 se usa como comprobación.
- No admite duplicados por fecha y número de operario.
- Reemplaza el mes completo para poder repetir una importación sin acumular duplicados.
- Escribe fecha más D:G en A:E y conserva las fórmulas F:K.
- Redimensiona la tabla mensual al número exacto de registros.
- Justo antes de escribir, guarda una copia del estado anterior en la carpeta `Copias de seguridad llamamientos`, junto al destino.
- Abre directamente la Herramienta Llamamientos original seleccionada, modifica la hoja mensual y guarda los cambios en ese mismo archivo.
- Cancela la operación antes del respaldo si Excel está abierto, y cancela el guardado si el original está bloqueado, es de solo lectura o Excel abre una ruta diferente.
- Valida después los datos, duplicados, tabla y fórmulas del archivo original actualizado.

## Importante

La aplicación lee los valores calculados que Excel guardó en D:F. Por eso el origen debe haberse abierto, recalculado y guardado después de los cambios del cierre mensual. La escritura se hace mediante una instancia oculta de Microsoft Excel para conservar las demás hojas y funciones nativas del libro.
