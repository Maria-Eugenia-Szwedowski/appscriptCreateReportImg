

const TEMPLATE_FILE_ID = '________________FILE_ID________';
const DESTINATION_FOLDER_ID = '____________FOLDER_ID____________';


function generateFilesFromTemplate() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var formValues = dataRange.getValues();
  var docTemplate = DocumentApp.openById(TEMPLATE_FILE_ID);
  var folder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
  var headers = formValues[0];
 
  for (var rowIndex = 1; rowIndex < formValues.length; rowIndex++) {
    var formRow = formValues[rowIndex];
    var fileName = formRow[2]; // Ubicar el Unique ID para que no se repitan los nombres de los archivos. Acá esta en columna 3
   
    // Crear una copia del documento de plantilla
    var copy = DriveApp.getFileById(docTemplate.getId()).makeCopy();
    var copyDoc = DocumentApp.openById(copy.getId());
    var copyDocBody = copyDoc.getBody();
   
    // Reemplazar los marcadores en el documento con los datos de la fila
    for (var columnIndex = 0; columnIndex < formRow.length; columnIndex++) {
      var placeholder = "<<" + headers[columnIndex] + ">>";
      var value = formRow[columnIndex];
      copyDocBody.replaceText(placeholder, value);
    }
   
    try {
    // Insertar la imagen desde la celda "LinkImg"
    var linkImgCell = sheet.getRange(rowIndex + 1, headers.indexOf("LinkImg") + 1); // Ajusta el encabezado según tu hoja
    var linkImg = linkImgCell.getValue();
   
    if (linkImg) {
      var imageBlob = UrlFetchApp.fetch(linkImg).getBlob();
      var image = copyDocBody.appendImage(imageBlob);
     
      // Ajusta el tamaño de la imagen si es necesario
      image.setWidth(600).setHeight(550);
      }
    } catch (error) {
      Logger.log("Error al insertar imagen: " + error.message);
      // Continuar la ejecución del código incluso si hay un error con la imagen
    }


    // Insertar el texto al final del documento
    var today = new Date();
    var formattedDate = Utilities.formatDate(today, "GMT", "dd-MM-yyyy");
    copyDocBody.appendParagraph("\n\nInforme realizado el día de la fecha " + formattedDate);
   
   
    // Guardar y cerrar el documento de copia
    copyDoc.saveAndClose();
   
    // Crear un archivo Docs a partir de la copia
    var newFileDoc = copy.setName(fileName + ".doc");
   
    // Crear un archivo PDF a partir de la copia y convertirlo a PDF
    var newFilePdf = copy.getAs("application/pdf").setName(fileName + ".pdf");
   
    // Agregar los archivos a la carpeta de destino
    //folder.createFile(newFileDoc);
    folder.createFile(newFilePdf);
   
    // Mover la copia del documento a la papelera
    //DriveApp.getFileById(copy.getId()).setTrashed(true);
  }
}
