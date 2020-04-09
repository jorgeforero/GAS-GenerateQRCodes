//
// QR Generator
// Generación de códigos QR a partir de los datos registrads en hoja de calculo actual
// jorge.forero@entrenoaldia.com
//

// Definiciones
// Tamaño Imagen QR
QRSIZE = '400x400';
// URL Google Charts API
URLCHRTAPI = 'https://chart.googleapis.com/chart?';
// Número de columnas de la hoja en donde se generarán los códigos
// Se generan n-1 columnas
COLUMNS = 5;
// Abajo de cada código se copia el correo correspondiente para referencia. Tamaño del texto
SIZEFONTLABEL = 5;
// Nombre de la hoja en donde se generan los códigos
SHEETNAMEGEN = 'QR Generados';

/**
* GeneraQR
* Genera los códigos QR a partir de la lista de valores encontrada en la hoja "Datos". La imagen es generada
* usando el Char API y es registrada en la hoja usando la fórmula =IMAGE() en cada celda en la hoja SHEETNAMEGEN.
* Una vez terminada la generación, se marcan los registros con "OK" en la columna de control D de la hoja de datos
*
* @param {void} void -
* @rerurn {void} - Códigos generados en la hoja NOMBREHOJAGEN
**/
function generateQRs() {

  // Apertura de la hoja donde se encuentran los datos
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = book.getSheetByName('Datos');
  var lastRow = dataSheet.getLastRow();
  var dataRange = dataSheet.getDataRange().getValues();

  // Formula generación del código QR al api de Chart
  var Url = URLCHRTAPI + 'chs=' + QRSIZE + '&cht=qr&chl=';

  // Crea una nueva hoja para almacenar los codigos generados
  var genSheet = book.insertSheet().setName(SHEETNAMEGEN);

  // Contadores que controlan la disposición de los codigos en la nueva hoja
  var rowGen = 1;
  var colGen = 1;
  var fonSizeArray = [[SIZEFONTLABEL]];
  var codesGenerated = 0;

  // Se recorre el arreglo de datos a generar
  for (var index=1; index<dataRange.length; index++) {
    var record = dataRange[index];
    var name = record[0];
    var lastname = record[1];
    var email = record[2];
    var flag = record[3];

    // Controla el estado del registro evaluado en la hoja de datos
    var row = index + 1;

    if (flag == '') {
     // Generación de la formula en la hoja de cálculo para generar código
     var qr = 'image(\"' + Url + name + '+' + lastname + '+' + email + '\")';
     // Registro de información en la nueva hoja y en la hoja de datos
     genSheet.setRowHeight(rowGen, 250).setColumnWidth(colGen, 250);
     genSheet.getRange(rowGen, colGen).setFormula(qr).setHorizontalAlignment('center').setVerticalAlignment('middle');
     genSheet.getRange(rowGen + 1, colGen).setValue(email).setHorizontalAlignment('center').setVerticalAlignment('top').setFontSizes(fonSizeArray);
     dataSheet.getRange('D' + row).setValue('OK').setHorizontalAlignment('center');
     colGen++;
     codesGenerated++;

     // determina el cambio de fila para tener una distribución de los codigo en matriz
     if (colGen == COLUMNS) {
         rowGen += 2;
         colGen = 1;
      };//if

    };//if
  };//for

  Browser.msgBox('Fueron generados ' + codesGenerated);
};

/**
* clearFlags
* Limpia la columna de 'Generado' de la hoja de datos y elimina la hoja SHEETNAMEGEN
*
* @param {void} void -
* @rerurn {void} -
**/
function clearFlags() {
  // Apertura de la hoja donde se encuentran los datos
  var book = SpreadsheetApp.getActiveSpreadsheet()
  var dataSheet = book.getSheetByName('Datos');
  // Limpia los contenidos de la columna 4(D) correspondiente a 'Generado'
  dataSheet.getRange(2, 4, dataSheet.getLastRow(), 1).clearContent();
  // Borra la hoja SHEETNAMEGEN donde se generan los códigos QR
  var genSheet = book.getSheetByName(SHEETNAMEGEN);
  book.deleteSheet(genSheet);
};

/**
* onOpen
* Adiciona la opcion Acciones al menu de la hoja
**/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Se adiciona el menu 'Acciones'
  ui.createMenu('Acciones')
      .addItem('Generar QRs', 'generateQRs')
      .addItem('Borrar Flags', 'clearFlags')
      .addToUi();
};
