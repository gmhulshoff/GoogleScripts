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
  this.values = rows.getValues();
  this.headers = this.values.shift();
}

this.getValue = function(val, key) {
  if (key == 'incassodate' || key == 'dtofsgntr')
    return Utilities.formatDate(val, 'CET', 'yyyy-MM-dd');
  else if (key == 'creationdatetime')
    return Utilities.formatDate(val, 'CET', "yyyy-MM-dd'T'HH:mm:ss'Z'");
  else
    return val;
}

this.readSettings = function(values, map) {
  var setting = {};
  for (var i = 0, pair; pair = values[i]; i++)
    if (pair[0] && pair[0] != '' && map[pair[0]])
      setting[map[pair[0]]] = getValue(pair[1], map[pair[0]]);
  return setting;
}

this.getSettingValues = function(sheetName) {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getSheetByName(sheetName);
  var rows = sheet.getDataRange();
  var values = rows.getValues();
  var map = {
    'Verenigingsnaam' : 'initiatingparty',
    'IBAN' : 'initiatingiban',
    'BIC' : 'initiatingbic',
    'Kenmerk' : 'endtoendid',
    'Incassant ID' : 'initiatingcreditoridentifier',
    'Datum afschrijving' : 'incassodate',
    'Datum aangemaakt' : 'creationdatetime',
    'Adres' : 'adresClub',
    'Postcode' : 'postcodeClub',
    'Plaats' : 'plaatsClub',
    'Reden betaling' : 'redenBetaling',
    'Email' : 'email'
  };
  return readSettings(values, map);
}

this.readRecord = function(map, headers, row) {
  var record = {};
  for (var i = 0, key; key = map[headers[i]]; i++)
    if (key)
      record[key] = getValue(row[i], key);
  return record;
}

this.readRows = function(info, map) {
  var rows = [];
  for (var i = 0, row; row = info.values[i]; i++)
    rows.push(this.readRecord(map, info.headers, row));
  Logger.log(rows);
  return rows;
}

this.getDebtors = function() {
  var map = {
    'Bedrag' : 'instdamt',
    'Mandaat ID' : 'mndtid',
    'Eerste of herhaling' : 'seqtp',
    'Mandaatdatum' : 'dtofsgntr',
    'Betaling namens' : 'amdmntind',
    'Naam' : 'dbtrnm',
    'IBAN' : 'dbtriban',
    'Bericht' : 'dbtrustrd'
  };
  return readRows(new getValues('Debiteuren'), map);
}

this.getMembers = function() {
  var map = {
    'Naam' : 'dbtrnm',
    'IBAN' : 'dbtriban',
    'Adres' : 'adres',
    'Postcode' : 'postcode',
    'Plaats' : 'plaats',
    'Email' : 'email',
  };
  return readRows(new getValues('AspirantLeden'), map);
}

this.moveFileToFolder = function(file, folder) {
  var docsFile = DocsList.getFileById(file.getId());
  var folders = docsFile.getParents();
  for (var n in folders)
    docsFile.removeFromFolder(folders[n]);
  docsFile.addToFolder(folder);
}

function createIncassoXml() {
  this.saveXml = function (data) {
    var fileName = 'incasso.xml';
    var file = DriveApp.createFile(Utilities.newBlob(data, MimeType.PLAIN_TEXT, fileName));
    file.setName(fileName);
    moveFileToFolder(file, DocsList.getFolder('Sepa'));
    return file;
  }
  this.mailXmlToMe = function(xmlFile) {
    var settings = getSettingValues('Gegevens');
    var advancedArgs = {name:settings.initiatingparty, htmlBody: '', attachments: [xmlFile]}; 
    var emailText = 'Beste penningmeester,\n\n';
    emailText += 'Bij deze de incasso - xml die u naar uw bank kunt versturen.\n';
    emailText += 'U kunt dit bestand valideren bij de volgende url: https://ing-fvt.liaison.com/welcome.do\n';
    emailText += 'Hier kunt u inloggen met als gebruikersnaam ING10 en als wachtwootd Format10.\n';
    emailText += 'Vervolgens kunt u de optie kiezen onder - European direct debit - pain.008.001.02 - The Netherlands SEPA Direct Debit (BvN).\n';
    emailText += 'En daar het bijgevoegde bestand uploaden en laten valideren.\n\n';
    emailText += 'Met vriendelijke groeten,\n';
    emailText += 'De penningmeester.'
    MailApp.sendEmail(settings.initiatingparty + " <" + settings.email + ">", 'Automatisch incasso bestand.', emailText, advancedArgs);
    DocsList.getFileById(xmlFile.getId()).setTrashed(true);
  }

  this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Gegevens');
  var event = {};
  var map = {
    initiatingparty : 'Verenigingsnaam',
    initiatingiban : 'IBAN',
    initiatingbic : 'BIC',
    endtoendid : 'Kenmerk',
  };
  var settings = getSettingValues('Gegevens');
  event.parameter = {
    initiatingparty : settings.initiatingparty,
    initiatingiban : settings.initiatingiban,
    initiatingbic : settings.initiatingbic,
    endtoendid : settings.endtoendid,
    initiatingcreditoridentifier : settings.initiatingcreditoridentifier,
    incassodate : settings.incassodate,
    creationdatetime : settings.creationdatetime
  };
  var xmlFile = saveXml(CreatePain00800102.createPainMessage(getDebtors(), event));
  mailXmlToMe(xmlFile);
};

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
      var settings = getSettingValues('Gegevens');
      this.verenigingsnaam = settings.initiatingparty;
      this.adresClub = settings.adresClub;
      this.postcodeClub = settings.postcodeClub;
      this.plaatsClub = settings.plaatsClub;
      this.incassantId = settings.initiatingcreditoridentifier;
      this.kenmerk = settings.endtoendid;
      this.redenBetaling = settings.redenBetaling;
      this.tenNameVan = member.dbtrnm;
      this.adres = member.adres;
      this.postcode = member.postcode;
      this.plaats = member.plaats;
      this.iban = member.dbtriban;
      
      this.sepaFolder = DocsList.getFolder('Sepa');
      this.fileName = "Machtiging automatische incasso";
      this.document = DocumentApp.create(this.fileName);
      
      Logger.log(this);
      
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
    { name : "Stuur Incassobestand naar mezelf", functionName : "createIncassoXml" },
    { name : "Email automatische incasso formulieren", functionName : "sendDirectDebitPermissionForms" },
  ];
  spreadsheet.addMenu("Incasso", entries);
};
