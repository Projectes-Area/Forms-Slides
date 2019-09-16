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
      if(Date.parse(values[i][j]).toString() != "NaN"){
        values[i][j]=values[i][j].toLocaleDateString("ca-ES").replace(/\/ /g, "");
      }      
      requests.push({
        replaceAllText: {
          containsText: {
            text: '{{'+values[0][j]+'}}',
            matchCase: true
          },
          replaceText: values[i][j],
          pageObjectIds: diapo
        }
      });      
    }         
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
