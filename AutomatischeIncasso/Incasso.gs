this.deleteDocByName = function (fileName){
  var docs = DocsList.find(fileName);
  for (var n in docs)
    if (docs[n].getName() == fileName)
      DocsList.getFileById(docs[n].getId()).setTrashed(true);
}

this.getValues = function(sheetName) {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getSheetByName(sheetName);
  var rows = sheet.getDataRange();
  var values = rows.getValues();
  var headers = values.shift();
  return values;
}

this.getDebtors = function() {
  var debtors = [];
  var values = getValues('Debiteuren');
  for (var i = 0, row; row = values[i]; i++)
	debtors.push({
      instdamt: row[0], 
      mndtid: row[1], 
      seqtp: row[2], 
      dtofsgntr: Utilities.formatDate(row[3], "CET", "yyyy-MM-dd"), 
      amdmntind: row[4], 
      dbtrnm: row[5], 
      dbtriban: row[6], 
      dbtrustrd: row[7]
    });
  return debtors;
}

this.getMembers = function() {
  var members = [];
  var values = getValues('AspirantLeden');
  for (var i = 0, row; row = values[i]; i++)
	members.push({
      dbtrnm: row[0], 
      dbtriban: row[1], 
      adres: row[2],
      postcode: row[3],
      plaats: row[4],
      email: row[5]
    });

  return members;
}

function createIncassoXml() {
  this.saveXml = function (data) {
    var fileName = 'incasso.xml';
    deleteDocByName(fileName);
    var file = DriveApp.createFile(Utilities.newBlob(data, MimeType.PLAIN_TEXT, fileName));
    file.setName(fileName);
  }
  
  saveXml(CreatePain00800102.createPainMessage(getDebtors(), event));
};

this.moveFileToFolder = function(file, folder) {
  var docsFile = DocsList.getFileById(file.getId());
  var folders = docsFile.getParents();
  for (var n in folders)
    docsFile.removeFromFolder(folders[n]);
  docsFile.addToFolder(folder);
}

function sendDirectDebitPermissionForms()
{
  this.createSepaDdMandate = function () {
    this.createPdfForm = function(settings) {
      var docId = settings.document.getId();
      var file = DocsList.getFileById(docId);
      var pdfForm = DocsList.createFile(file.getAs('application/pdf'));
      moveFileToFolder(pdfForm, settings.sepaFolder);
      DocsList.getFileById(docId).setTrashed(true);
      return pdfForm;
    }
    
    this.createPersonalizedForm = function() {
      this.verenigingsnaam = sheet.getRange("B1").getValue();
      this.adresClub = sheet.getRange("B8").getValue();
      this.postcodeClub = sheet.getRange("B9").getValue();
      this.plaatsClub = sheet.getRange("B10").getValue();
      this.incassantId = sheet.getRange("B12").getValue();
      this.kenmerk = sheet.getRange("B13").getValue();
      this.redenBetaling = sheet.getRange("B14").getValue();
      this.tenNameVan = member.dbtrnm;
      this.adres = member.adres;
      this.postcode = member.postcode;
      this.plaats = member.plaats;
      this.iban = member.dbtriban;
      
      this.sepaFolder = DocsList.getFolder('Sepa');
      this.fileName = "Machtiging automatische incasso";
      this.document = DocumentApp.create(this.fileName);
      
      CreatePain00800102.createSignupForm(this);
      
      this.document.saveAndClose();

      moveFileToFolder(this.document, this.sepaFolder);
    }
    
    this.settings = new this.createPersonalizedForm();
    this.pdfForm = this.createPdfForm(this.settings);
  }
  
  this.sendDirectDebitPermissionForm = function() {
    var info = new createSepaDdMandate();
      
    var advancedArgs = {name:info.settings.verenigingsnaam, htmlBody: '', attachments: [info.pdfForm.getAs(MimeType.PDF)]}; 
    var emailText = 'Beste ' + member.dbtrnm + ',\n\n';
    emailText += 'Bij deze het formulier voor automatische machtiging voor het incasseren van de contributie.\n';
    emailText += 'U kunt dit formulier uitprinten, invullen en inleveren bij de penningmeester.\n\n';
    emailText += 'Met vriendelijke groeten,\n';
    emailText += 'De penningmeester.'
    MailApp.sendEmail(member.dbtrnm + " <" + member.email + ">", 'Machtiging automatische incasso.', emailText, advancedArgs);
    DocsList.getFileById(info.pdfForm.getId()).setTrashed(true);
  }
 
  this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Gegevens');
  var members = getMembers();
  for (var i = 0; this.member = members[i]; i++)
    this.sendDirectDebitPermissionForm();
}

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    { name : "Maak Incassobestand", functionName : "createIncassoXml" },
    { name : "Email automatische incasso formulieren", functionName : "sendDirectDebitPermissionForms" },
  ];
  spreadsheet.addMenu("Incasso", entries);
};
