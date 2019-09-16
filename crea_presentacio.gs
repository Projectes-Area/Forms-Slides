function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Presentació')
    .addItem('Crea presentació', 'crea_presentacio')
    .addToUi();
}

function crea_presentacio() {
  var templatePresentationId = "168PtJxeYYJTpNZsRvHvxOvKtPRSEfuFOIwCLM69s5hI";
  var values = SpreadsheetApp.getActive().getSheets()[0].getDataRange().getValues();
  var copyFile = {
    title: 'prova Presentació',
    parents: [{id: 'root'}]
  };
  
  // Copiar la plantilla de la presentació
  
  copyFile = Drive.Files.copy(copyFile, templatePresentationId);
  var presentationCopyId = copyFile.id;
  var diapositiva = Slides.Presentations.get(presentationCopyId).slides[0].objectId;
  
  for (var i=1;i<values.length;i++) {   
    // Copiar la primera diapositiva
    
    requests = [{
      duplicateObject: {
        "objectId": diapositiva
      }
    }]   
    var result = Slides.Presentations.batchUpdate({
      requests: requests
    }, presentationCopyId);    
    var diapo = result.replies[0].duplicateObject.objectId;
    
    // Fer les substitucions
    
    requests = [];    
    for (var j=1;j<values[0].length;j++) {      
      requests.push({
        replaceAllText: {
          containsText: {
            text: '{{'+values[0][j]+'}}',
            matchCase: true
          },
          replaceText: values[i][j],
          pageObjectIds: diapo
      }});
    }    
    
    // Afegir imatge
    
    var imageUrl = 'https://www.google.com/images/branding/googlelogo/2x/' +
        'googlelogo_color_272x92dp.png';
    var cm1 = 360000 // 1 cm EMU
    requests.push({
      createImage: {
        url: imageUrl,
        elementProperties: {
          pageObjectId: diapo,
          size: {
            height: {magnitude: 2 * cm1,
                    unit: 'EMU'},
            width: {magnitude: 3 * cm1,
                    unit: 'EMU'}
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: 20 * cm1,
            translateY: 5 * cm1,
            unit: 'EMU'
          }
        }
      }
    });    
    result = Slides.Presentations.batchUpdate({
        requests: requests
    }, presentationCopyId);    
  }
  
  // Esborrar la primera diapositiva
  
  requests = [{
    deleteObject: {
      "objectId": diapositiva
    }
  }];    
  result = Slides.Presentations.batchUpdate({
    requests: requests
  }, presentationCopyId); 
}