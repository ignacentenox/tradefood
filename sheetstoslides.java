// Version: 1.0
function copiarDatosASlides() {
    // ID de la hoja de Google Sheets
    var sheetId = '1S_xmGExfkZflR9KHwK26la1ti9lmFicDO7EBeGQDPQY';
    
    // ID de la presentación de Google Slides
    var slidesId = '1mE0wUBFCRElIw6d_DNppuje8_ubzuClffrHRt3O6oVk';
  
    // Abre el archivo de Google Sheets
    var spreadsheet = SpreadsheetApp.openById(sheetId);
    
    // Obtiene las hojas específicas por su nombre
    var sheetB = spreadsheet.getSheetByName('PRESENTACION EN LINEA'); 
    var sheetC = spreadsheet.getSheetByName('PRESENTACION EN LINEA');
    var sheetD = spreadsheet.getSheetByName('PRESENTACION EN LINEA');
    var sheetE = spreadsheet.getSheetByName('PRESENTACION EN LINEA');
    var sheetL = spreadsheet.getSheetByName('PRESENTACION EN LINEA');
    
    // Abre la presentación de Google Slides
    var presentation = SlidesApp.openById(slidesId);
    
    //valores de las celdas B2, C2, D2, E2 y L2 de las diferentes hojas
    var cellValues = [
      sheetB.getRange("B2").getValue(),
      sheetC.getRange("C2").getValue(),
      sheetD.getRange("D2").getValue(),
      sheetE.getRange("E2").getValue(),
      sheetL.getRange("L2").getValue()
    ];
    
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
