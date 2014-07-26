function createIncassoXml() {
  this.deleteDocByName = function (fileName){
    var docs = DocsList.find(fileName);
    for (n in docs)
      if (docs[n].getName() == fileName)
        DocsList.getFileById(docs[n].getId()).setTrashed(true);
  }
  this.saveXml = function (data) {
    var fileName = 'incasso.xml';
    deleteDocByName(fileName);
    var file = DriveApp.createFile(Utilities.newBlob(data, MimeType.PLAIN_TEXT, fileName));
    file.setName(fileName);
  }
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var debtorSheet = spreadSheet.getSheetByName('Debiteuren');
  var settingsSheet = spreadSheet.getSheetByName('Gegevens');
  var debtorRows = debtorSheet.getDataRange();
  var debtorValues = debtorRows.getValues();
  var event = {
    parameter : 
    {
      initiatingparty : settingsSheet.getRange("B1").getValue(),
      initiatingiban : settingsSheet.getRange("B2").getValue(),
      initiatingbic : settingsSheet.getRange("B3").getValue(),
      initiatingcreditoridentifier : settingsSheet.getRange("B4").getValue(),
      incassodate : Utilities.formatDate(settingsSheet.getRange("B5").getValue(), "CET", "yyyy-MM-dd'T'HH:mm:ss"),
	  creationdatetime : Utilities.formatDate(new Date(), "CET", "yyyy-MM-dd'T'HH:mm:ss"),
      endtoendid : settingsSheet.getRange("B6").getValue()
    }
  }
  var debtors = [];
  var debtorHeaders = debtorValues.shift();
  for (var i = 0, row; row = debtorValues[i]; i++)
	debtors.push({instdamt: row[0], mndtid: row[1], seqtp: row[2], 
	dtofsgntr: Utilities.formatDate(row[3], "CET", "yyyy-MM-dd"), 
	amdmntind: row[4], dbtrnm: row[5], dbtriban: row[6], dbtrustrd: row[7]});
  
  saveXml(CreatePain00800102.createPainMessage(debtors, event));
};

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Maak Incassobestand",
    functionName : "createIncassoXml"
  }];
  spreadsheet.addMenu("Incasso", entries);
};
