function crea_presentacio() {
  var templatePresentationId = "1d6DXRKDIEiZlmkDG_6X5p8-MFGeUlhWMaU7jMw5L7co";
  var values = SpreadsheetApp.getActive().getSheets()[0].getRange(2,2,3,2).getValues();
  var copyTitle = 'presentation';
  var copyFile = {
    title: copyTitle,
    parents: [{id: 'root'}]
  };
  copyFile = Drive.Files.copy(copyFile, templatePresentationId);
  var presentationCopyId = copyFile.id;
  var presentation = Slides.Presentations.get(presentationCopyId);
  var slides = presentation.slides;
  var diapositiva = slides[0].objectId;
  for (var i = 0; i < values.length; ++i) {
    var row = values[i];
    var nom = row[0]; 
    var professio = row[1];
    requests = [{
      duplicateObject: {
        "objectId": diapositiva
      }
    }]
    
    var result = Slides.Presentations.batchUpdate({
      requests: requests
    }, presentationCopyId);
    
    var diapo = result.replies[0].duplicateObject.objectId;

    requests = [{
      replaceAllText: {
        containsText: {
          text: '{{Nom}}',
          matchCase: true
        },
        replaceText: nom,
        pageObjectIds: diapo
      }
    }, {
      replaceAllText: {
        containsText: {
          text: '{{Professio}}',
          matchCase: true
        },
        replaceText: professio,
        pageObjectIds: diapo
      }
    }];

    result = Slides.Presentations.batchUpdate({
      requests: requests
    }, presentationCopyId);
  }
}