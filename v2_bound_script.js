// Cal habilitar a Recursos-Serveis avançats de Google: Gmail API

function doGet() {
  // Log the email address of the person running the script.
    var usuari = Session.getActiveUser().getEmail()
    return ContentService.createTextOutput(usuari)
  }
  
  function onFormSubmit(e){
    var numColEditURL = 3
    var numColNumRev = 2
    
    // Registrar les vegades que s'ha enviat el formaulari d'aquesta resposta
    var response = e.range;
    var rowIndex = response.getRow() + 1;
    var sheet = SpreadsheetApp.getActive().getSheetByName("Dades");
    var data = sheet.getDataRange().getValues();
    var rev = 0;
    if (data[rowIndex][1] != "") {
      rev = parseInt(data[rowIndex][1]);
    }  
    sheet.getRange(rowIndex + 1, numColNumRev).setValue(rev + 1);
     
    if (rev > 0) {
        var strAccio = "editat una"
    } else {
        var strAccio = "enviat una nova"
        
        // Desar la URL per editar la resposta al formulari     
        var ssFrm = SpreadsheetApp.getActiveSpreadsheet();
        var formUrl = ssFrm.getFormUrl();
        var form = FormApp.openByUrl(formUrl); 
        var resForm = form.getResponses();         
        var url = resForm[resForm.length - 1].getEditResponseUrl();
        var col = sheet.getRange("A1:A").getValues();
        var fila = col.filter(String).length;
        sheet.getRange(fila, numColEditURL).setValue(url);         
    }
    
    // Enviar un e-mail a la persona responsable
    var id = rowIndex - 2;
    var emailAddress = 'recursos@xtec.cat';
    var subject = 'Informe de formació';
    var message = "S'ha " + strAccio + " resposta al formulari d'informe de formació. Pots consultar la <a href='https://edumet.cat/areatac/p2/index.php?ID=1pPH_fvZ9Jh75P6SRMaN4xEl6s4k1BSl0gvFOANtnlE4&config=Config&fltr=[id=" + id + "]'>fitxa de formació</a> corresponent.";
    MailApp.sendEmail({
      to: emailAddress,
      subject: subject,
      htmlBody: message
    });  
  }
  
  /*function ompleURLs() {
      var numColEditURL = 3
      var sheet = SpreadsheetApp.getActive().getSheetByName("Dades");
      var ssFrm = SpreadsheetApp.getActiveSpreadsheet();
      var formUrl = ssFrm.getFormUrl();
      var form = FormApp.openByUrl(formUrl); 
      var resForm = form.getResponses();         
      for (i=0;i<resForm.length;i++) {   
        var url = resForm[i].getEditResponseUrl();
        sheet.getRange(i + 20, numColEditURL).setValue(url);   
      }
  }*/  