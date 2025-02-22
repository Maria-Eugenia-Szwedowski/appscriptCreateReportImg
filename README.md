# appscriptCreateReportImg

📌 Generador de Informes con Imágenes desde Google Sheets a Google Docs y PDF

Este proyecto, desarrollado por María Eugenia Szwedowski, automatiza la generación de informes a partir de datos en Google Sheets, utilizando Google Apps Script. Ahora, también incluye la funcionalidad de insertar imágenes almacenadas en Google Drive en los documentos generados.


🛠️ Funcionalidad Principal

✔️ Toma datos de una hoja de cálculo en Google Sheets.
✔️ Usa una plantilla de Google Docs con marcadores para reemplazar información.
✔️ Reemplaza los marcadores con los datos de cada fila.
✔️ Inserta imágenes almacenadas en Google Drive en la ubicación deseada del documento.
✔️ Agrega la fecha de generación del informe al final del documento.
✔️ Genera automáticamente archivos .docx y .pdf.
✔️ Guarda los archivos en una carpeta específica de Google Drive.


📂 Organización del Código

El código está estructurado en un solo archivo de Google Apps Script.


📎 Requisitos

🔹 Google Sheets con una estructura de datos organizada en columnas.
🔹 Google Docs con una plantilla que contenga marcadores (<<nombre>>, <<fecha>>, etc.).
🔹 Google Drive con permisos de acceso adecuados.
🔹 Columna con enlaces de imágenes en Google Sheets (LinkImg). Ver pasos previos.
🔹 Configuración de los IDs de la plantilla y la carpeta de destino.


📌 Paso a Paso

1️⃣ Definir variables clave:

TEMPLATE_FILE_ID: ID del archivo de Google Docs que se usará como plantilla.
DESTINATION_FOLDER_ID: ID de la carpeta de Google Drive donde se guardarán los archivos generados.

2️⃣ Obtener y procesar datos:

sheet.getActiveSheet(): Obtiene la hoja activa.
sheet.getDataRange().getValues(): Extrae todos los datos de la hoja en una matriz (formValues).
docTemplate: Abre la plantilla en Google Docs.
folder: Accede a la carpeta de destino.
headers: Guarda la primera fila de la hoja de cálculo (encabezados de las columnas).

3️⃣ Generar archivos para cada fila de datos:

Se empieza desde rowIndex = 1 para omitir la fila de encabezados.
fileName: Obtiene un valor único (columna 3) para nombrar el archivo.
Se crea una copia de la plantilla y se abre como documento editable.
Se reemplazan los marcadores con los valores correspondientes.

4️⃣ Inserción de imágenes desde Drive:

Se obtiene la URL de la imagen desde la columna LinkImg.
Se descarga la imagen y se inserta en el documento.
Se ajusta el tamaño de la imagen (600x550 px).
Se maneja cualquier error en la inserción sin detener la ejecución del código.

5️⃣ Agregar la fecha de generación:

Se inserta un párrafo al final del documento con la fecha actual.

6️⃣ Guardar y generar archivos:

Se guarda y cierra el documento.
Se crea un archivo .docx y un .pdf.
Se guardan en la carpeta de destino en Google Drive.


📌 Resumen del Flujo

✅ Lee los datos desde Google Sheets.
✅ Abre la plantilla en Google Docs.
✅ Crea una copia del documento.
✅ Reemplaza los marcadores con los datos de la fila.
✅ Inserta imágenes desde Google Drive en la plantilla.
✅ Añade la fecha de generación al final del informe.
✅ Guarda los archivos generados en la carpeta de Drive.
✅ Convierte el documento a PDF.
✅ Guarda el PDF en la carpeta de destino.


📌 PASOS PREVIOS PARA LinkImg: Transformación de Enlace de Google Drive para Insertar Imágenes en Documentos

En la columna LinkImg de Google Sheets, se usa la siguiente fórmula para convertir enlaces de archivos de Google Drive en URLs de acceso directo para su inserción en los informes generados en Google Docs.


📌 Fórmula Utilizada:

excel
CopiarEditar

=IZQUIERDA("https://drive.google.com/uc?id=" & DERECHA(AR2; LARGO(AR2) - LARGO("https://drive.google.com/file/d/"));
LARGO("https://drive.google.com/uc?id=" & DERECHA(AR2; LARGO(AR2) - LARGO("https://drive.google.com/file/d/")-18)))



📌 Explicación Paso a Paso

Cuando subes una imagen a Google Drive y la compartes, el enlace generado tiene la siguiente estructura:
bash
CopiarEditar
https://drive.google.com/file/d/XXXXXXXXXXXXXXXXXX/view?usp=sharing


donde XXXXXXXXXXXXXXXXXX es el ID del archivo.
Para insertar imágenes en Google Docs desde Google Sheets, es necesario obtener una URL directa de descarga, que tiene esta estructura:
bash
CopiarEditar
https://drive.google.com/uc?id=XXXXXXXXXXXXXXXXXX


La fórmula extrae correctamente el ID del archivo y lo agrega a la URL de descarga directa.


📌 Desglose de la Fórmula

DERECHA(AR2; LARGO(AR2) - LARGO("https://drive.google.com/file/d/"))
Se extrae la parte derecha del enlace de Google Drive, eliminando la parte inicial https://drive.google.com/file/d/.
Esto deja el **ID del archivo + "/view?usp=sharing"`.
DERECHA(AR2; LARGO(AR2) - LARGO("https://drive.google.com/file/d/") - 18)
Se usa -18 para eliminar el texto "/view?usp=sharing" del final y quedarse solo con el ID del archivo.
El número -18 varía si Google cambia la estructura del enlace (es importante verificar).
Concatenación para obtener el enlace directo:
excel
CopiarEditar
"https://drive.google.com/uc?id=" & [ID del archivo]


Se añade la URL base https://drive.google.com/uc?id= al ID del archivo extraído.


📌 Ejemplo Práctico

✅ Enlace original de Google Drive:

bash
CopiarEditar
https://drive.google.com/file/d/11111111111111111111/view?usp=sharing


✅ Salida esperada en LinkImg:

bash
CopiarEditar
https://drive.google.com/uc?id=11111111111111111111


Este nuevo enlace permite descargar la imagen directamente, lo que hace posible insertarla en el informe generado en Google Docs.


📌 Resumen y Consideraciones

✔ La fórmula automatiza la conversión del enlace compartido en una URL directa de descarga.
✔ El -18 es crucial para eliminar la parte final del enlace de Google Drive y obtener solo el ID.
✔ Si Google cambia la estructura del enlace, el número -18 podría necesitar ajustes.
✔ Este proceso permite que Google Apps Script obtenga la imagen y la inserte en el documento sin errores.
