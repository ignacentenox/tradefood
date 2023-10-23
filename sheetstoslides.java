// Version: 1.0
function copiarDatosASlides() {
    // ID de la hoja de Google Sheets
    var sheetId = '1S_xmGExfkZflR9KHwK26la1ti9lmFicDO7EBeGQDPQY';
    
    // ID de la presentación de Google Slides
    var slidesId = '1mE0wUBFCRElIw6d_DNppuje8_ubzuClffrHRt3O6oVk';
  
    // Abre el archivo de Google Sheets
    var spreadsheet = SpreadsheetApp.openById(sheetId);
    
    // Obtiene las hojas específicas por su nombre
    var sheetE = spreadsheet.getSheetByName('PRESENTACION EN LINEA'); 
    var sheetU = spreadsheet.getSheetByName('PRESENTACION EN LINEA');
    var sheetL = spreadsheet.getSheetByName('PRESENTACION EN LINEA');
    var sheetI = spreadsheet.getSheetByName('PRESENTACION EN LINEA');
    var sheetV = spreadsheet.getSheetByName('PRESENTACION EN LINEA');
    var sheetQ = spreadsheet.getSheetByName('PRESENTACION EN LINEA');
    var sheetP = spreadsheet.getSheetByName('PRESENTACION EN LINEA');
    var sheetS = spreadsheet.getSheetByName('PRESENTACION EN LINEA');
    var sheetR = spreadsheet.getSheetByName('PRESENTACION EN LINEA');
    
    // Abre la presentación de Google Slides
    var presentation = SlidesApp.openById(slidesId);
    
//valores de las celdas E2, U2, L2, I2, V2, Q2, P2 y S2 de las diferentes hojas
var cellValues = [
  sheetE.getRange("E2").getValue(),
  sheetU.getRange("U2").getValue(),
  sheetL.getRange("L2").getValue(),
  sheetI.getRange("I2").getValue(),
  sheetV.getRange("V2").getValue(),
  sheetQ.getRange("Q2").getValue(),
  sheetP.getRange("P2").getValue(),
  sheetR.getRange("R2").getValue()
  sheetS.getRange("S2").getValue()
];

    //C2 LOTE (NUMERO)
    //E2 UBICACION
    //U2 CANTIDAD
    //L2 PESO PROM
    //I2 ESTADO
    //V2 SANIDAD
    //Q2 DESBASTE
    //P2 PLAZO
    //S2 REVISO
    //R2 OBSERVACIONES
    
    //diapositivas de la presentación
    var slides = presentation.getSlides();
    
    // se necesita de dos diapositivas antes de continuar
    if (slides.length >= 2) {
      // Obtiene la segunda diapositiva en la presentación de Google Slides
      var slide = slides[1];
      
      // Obtiene los elementos de la página de la diapositiva
      var pageElements = slide.getPageElements();
      
      // Para cada valor de celda, actualiza el cuadro de texto correspondiente
      for (var i = 0; i < cellValues.length; i++) {
        var text = cellValues[i];
  
        // Obtiene el cuadro de texto correspondiente
        var textBox = pageElements[i];
  
        // Verifica si el elemento de la página es una forma
        if (textBox && textBox.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
          // Actualiza el texto del cuadro de texto
          textBox.asShape().getText().setText(text);
    
          // Cambia el tamaño de letra y la tipografía del cuadro de texto
          var textRange = textBox.asShape().getText();
          var textStyle = textRange.getTextStyle();
          textStyle.setFontSize(40); // Cambia 40 al tamaño de letra que desees (en puntos)
        } else {
          // Si el cuadro de texto no existe, lo crea
          var titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 100, 100 * i, 300, 50);
          titleShape.getText().setText(text);
          var textRange = titleShape.getText();
          var textStyle = textRange.getTextStyle();
          textStyle.setFontSize(40); // Cambia 40 al tamaño de letra que desees (en puntos)
        }
      }
    }
}
