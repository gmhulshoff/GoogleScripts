this.logInSpreadSheet = function(msg) {
  var sheet = spreadSheet.getSheetByName('Gegevens');
  sheet.getRange("A15").setValue(msg);
}

this.deleteDocByName = function (fileName){
  var docs = DocsList.find(fileName);
  for (var n in docs)
    if (docs[n].getName() == fileName)
      DocsList.getFileById(docs[n].getId()).setTrashed(true);
}

this.getValues = function(sheetName) {
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
  else if (key == 'dbtrnm')
    return val.replace(/[^a-zA-Z0-9\/-?:().,'+ ]/g, '?');
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
  this.spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
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

  this.sheet = spreadSheet.getSheetByName('Gegevens');
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
    this.spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    this.startNewBatch = function(i) { return (i % 50) == 0; }
    this.formatNumber = function(i) { return (i < 10 ? '0' : '') + (i < 100 ? '0' : '') + i; }
    this.addPersonalizedForm = function(i, settings, member) {
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
    
    this.document = addNewDocumentToBatch();
    
    logInSpreadSheet(i);
    
    CreatePain00800102.createSignupForm(this);
  }
  
  this.sendDirectDebitPermissionForm = function(settings) {
    var advancedArgs = {name:settings.verenigingsnaam, htmlBody: '', attachments: blobs}; 
    var emailText = 'Beste penningmeeser,\n\n';
    emailText += 'Bij deze de formulieren voor automatische machtiging voor het incasseren van de contributie.\n\n';
    emailText += 'Met vriendelijke groeten,\n';
    emailText += 'De penningmeester.'
    MailApp.sendEmail(settings.initiatingparty + " <" + settings.email + ">", "Machtiging automatische incasso", emailText, advancedArgs);
  }
  
  this.createMemberSignupFormsDocument = function(settings) {
    this.saveDocument = function() {
      if (document == null) return;
      document.saveAndClose();
      blobs.push(document.getAs(MimeType.PDF));
      DocsList.getFileById(document.getId()).setTrashed(true);
    }
    
    this.addPageBreakToCurrenDocument = function() {
      document.appendPageBreak();
      return document;
    }
    
    this.addNewDocumentToBatch = function() {
      if (!startNewBatch(i)) 
        return addPageBreakToCurrenDocument();
      saveDocument();
      return DocumentApp.create("Machtiging automatische incasso (" + formatNumber(i + 1) + '-' + formatNumber(Math.min(totalCount, i + 50)) + ")");
    }
    
    var members = getMembers();
    var blobs = [];
    this.document = null;
    var totalCount = members.length;

    for (var i = 0, member; member = members[i]; i++)
      addPersonalizedForm(i, settings, member);

    saveDocument();
    return blobs;
  }

  var settings = getSettingValues('Gegevens');
  var blobs = this.createMemberSignupFormsDocument(settings);  
  sendDirectDebitPermissionForm(settings);
}

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    { name : "Stuur Incassobestand naar mezelf", functionName : "createIncassoXml" },
    { name : "Email automatische incasso formulieren", functionName : "sendDirectDebitPermissionForms" },
  ];
  spreadsheet.addMenu("Incasso", entries);
};
