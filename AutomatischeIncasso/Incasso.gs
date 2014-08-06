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
      lidnummer: row[12]
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

function createPDF()
{
  this.createSepaDdMandate = function (){
    this.moveFileToFolder = function(file, folder) {
      var folders = file.getParents();
      for (var n in folders)
        file.removeFromFolder(folders[n]);
      file.addToFolder(folder);
    }
    
    this.enterFormFields = function() {
      var document = DocumentApp.openById(personalizedForm.getId());
      var body = document.getBody();
      
      body.replaceText('{verenigingsnaam}', sheet.getRange("B1").getValue());
      body.replaceText('{adres club}', sheet.getRange("B8").getValue());
      body.replaceText('{postcode club}', sheet.getRange("B9").getValue());
      body.replaceText('{plaats club}', sheet.getRange("B10").getValue());
      body.replaceText('{incassantid}', sheet.getRange("B12").getValue());
      body.replaceText('{kenmerk}', sheet.getRange("B13").getValue());
      
      body.replaceText('{tennamevan}', member["tennamevan"]);
      body.replaceText('{adres}', member["adres"]);
      body.replaceText('{postcode}', member["postcode"]);
      body.replaceText('{plaats}', member["plaats"]);
      body.replaceText('{iban}', member["dbtriban"]);

      document.saveAndClose();

      return this.createPdfForm("machtiging-" + member["lidnummer"] + ".pdf");
    }
    
    this.createPdfForm = function(name) {
      var pdfForm = DocsList.createFile(personalizedForm.getAs('application/pdf'));
      personalizedForm.setTrashed(true);
      deleteDocByName(name);
      pdfForm.rename(name);
      moveFileToFolder(pdfForm, DocsList.getFolder('Sepa'));
      return pdfForm;
    }
    
    this.moveToFolder = function(file, folder) {
      this.personalizedForm = file.makeCopy('cloneSepaMachtiging');
      moveFileToFolder(personalizedForm, folder);
    }
    
    this.createPersonalizedForm = function() {
      var folder = DocsList.getFolder('Sepa');
      var files = folder.find('Sjabloon SEPA machtiging');
      for (var n in files)
        moveToFolder(files[n], folder);
    }
    
    createPersonalizedForm();
    return enterFormFields();
  }
 
  this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Gegevens');
  this.member = getDebtors()[0];
  
  return createSepaDdMandate();
}

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    { name : "Maak Incassobestand", functionName : "createIncassoXml" },
    { name : "Maak Machtigings PDF", functionName : "createPDF" },
  ];
  spreadsheet.addMenu("Incasso", entries);
};
