// Cal habilitar a Recursos-Serveis avançats de Google: Drive API i Google Slides API

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Presentació')
    .addItem('Crea presentació', 'crea_presentacio')
    .addToUi();
}

function crea_presentacio() {
  var cm1 = 360000 // 1 cm EMU
  var templatePresentationId = "1d6DXRKDIEiZlmkDG_6X5p8-MFGeUlhWMaU7jMw5L7co";
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
      if(values[0][j] == "Material" || values[0][j] == "Enllaços" ) { // Columnes amb valors separats per comes
        if(values[0][j] == "Material") {
          var offset = 11.11;
        }
        if(values[0][j] == "Enllaços") {
          var offset = 7.93;
        }
        var valor = values[i][j].toString().split(",");
        for(var k=0;k<valor.length;k++){
          requests.push({
            createShape: {
              objectId: values[0][j].slice(0,3) + i.toString() + k.toString(),
              shapeType:'TEXT_BOX',
              elementProperties:{
                pageObjectId: diapo,
                size: {
                  width: {magnitude: 3.20 * cm1,
                          unit: 'EMU'},
                  height: {magnitude: 0.37 * cm1,
                          unit: 'EMU'}
                },
                transform: {
                  scaleX: 1,
                  scaleY: 1,
                  translateX: 21.07 * cm1,
                  translateY: (offset + 0.37 * k) * cm1,
                  unit: 'EMU'
                }
              }       
            }
          });
          requests.push({
            insertText: {
              objectId: values[0][j].slice(0,3) + i.toString() + k.toString(),
              text:'Enllaç ' + (k + 1).toString(),         
            }
          });
          requests.push({
            updateTextStyle: {
              objectId: values[0][j].slice(0,3) + i.toString() + k.toString(),
              fields:'*',
              style:{
                fontSize:{
                  magnitude: 10,
                  unit: 'PT'
                },
                underline: true,
                link:{
                  url:valor[k].toString()
                }
              }
            }
          });
        }
      } else {            
        if(Object.prototype.toString.call(values[i][j]) === "[object Date]"){ // Columnes amb dates
          values[i][j]=values[i][j].toLocaleDateString("ca-ES").replace(/\/ /g, "");
        }      
        requests.push({
          replaceAllText: {
            containsText: {
              text: '{{'+values[0][j]+'}}',
              matchCase: true
            },
            replaceText: values[i][j].toString(),
            pageObjectIds: diapo
          }
        });      
      }      
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