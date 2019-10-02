// Cal habilitar a Recursos-Serveis avançats de Google: Gmail API

function sendMail() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Dades");
  var Avals = sheet.getRange("A1:A").getValues();
  var id = Avals.filter(String).length - 3;
  var emailAddress = 'recursos@xtec.cat';
  var subject = 'Nou informe de formació';
  var message = "S'ha enviat una nova resposta al formulari d'informe de formació. Pots consultar la <a href='https://edumet.cat/areatac/p2/index.php?ID=1pPH_fvZ9Jh75P6SRMaN4xEl6s4k1BSl0gvFOANtnlE4&config=Config&fltr=[id=" + id + "]'>fitxa de formació</a> corresponent.";
  MailApp.sendEmail({
    to: emailAddress,
    subject: subject,
    htmlBody: message
  });
}