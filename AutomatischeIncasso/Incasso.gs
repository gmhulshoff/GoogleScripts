this.deleteDocByName = function (fileName){
  var docs = DocsList.find(fileName);
  for (var n in docs)
    if (docs[n].getName() == fileName)
      DocsList.getFileById(docs[n].getId()).setTrashed(true);
}

this.getDebtors = function() {
  var debtors = [];
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var debtorSheet = spreadSheet.getSheetByName('Debiteuren');
  var settingsSheet = spreadSheet.getSheetByName('Gegevens');
  var debtorRows = debtorSheet.getDataRange();
  var debtorValues = debtorRows.getValues();
  var debtorHeaders = debtorValues.shift();
  for (var i = 0, row; row = debtorValues[i]; i++)
	debtors.push({
      instdamt: row[0], 
      mndtid: row[1], 
      seqtp: row[2], 
      dtofsgntr: Utilities.formatDate(row[3], "CET", "yyyy-MM-dd"), 
      amdmntind: row[4], 
      dbtrnm: row[5], 
      dbtriban: row[6], 
      dbtrustrd: row[7],
      adres: row[8],
      postcode: row[9],
      plaats: row[10],
      tennamevan: row[11],
      email: row[12]
    });
  return debtors;
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

function createPDF()
{
  this.createSepaDdMandate = function (){
    this.createPdfForm = function(settings) {
      var docId = settings.document.getId();
      var file = DocsList.getFileById(docId);
      deleteDocByName(settings.fileName + ".pdf");
      var pdfForm = DocsList.createFile(file.getAs('application/pdf'));
      DocsList.getFileById(docId).setTrashed(true);
      moveFileToFolder(pdfForm, settings.sepaFolder);
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
      this.tenNameVan = member["tennamevan"];
      this.adres = member["adres"];
      this.postcode = member["postcode"];
      this.plaats = member["plaats"];
      this.iban = member["dbtriban"];
      
      this.sepaFolder = DocsList.getFolder('Sepa');
      this.fileName = "Machtiging " + this.tenNameVan;
      this.document = DocumentApp.create(this.fileName);
      
      CreatePain00800102.createSignupForm(this);
      
      this.document.saveAndClose();

      moveFileToFolder(this.document, this.sepaFolder);
    }
    
    return createPdfForm(new createPersonalizedForm());
  }
 
  this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Gegevens');
  var members = getDebtors();
  for (var i = 0; this.member = members[i]; i++)
    createSepaDdMandate();
}

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    { name : "Maak Incassobestand", functionName : "createIncassoXml" },
    { name : "Maak Machtigings PDF", functionName : "createPDF" },
  ];
  spreadsheet.addMenu("Incasso", entries);
};
