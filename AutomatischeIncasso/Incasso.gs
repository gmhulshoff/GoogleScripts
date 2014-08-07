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
    this.createPdfForm = function() {
      var pdfForm = DocsList.createFile(personalizedForm.getAs('application/pdf'));
      DocsList.getFileById(personalizedForm.getId()).setTrashed(true);
      deleteDocByName(fileName);
      pdfForm.rename(fileName + ".pdf");
      moveFileToFolder(pdfForm, DocsList.getFolder('Sepa'));
      return pdfForm;
    }
    
    this.createPersonalizedForm = function() {
      var folder = DocsList.getFolder('Sepa');
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
      
      this.fileName = "Machtiging " + member["tennamevan"];
      this.personalizedForm = createTemplate();

      moveFileToFolder(personalizedForm, folder);
    }
    
    createPersonalizedForm();
    return createPdfForm();
  }
 
  this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Gegevens');
  var members = getDebtors();
  for (var i = 0; this.member = members[i]; i++)
    createSepaDdMandate();
}

this.createTemplate = function() {
  this.setDocProperties = function(body) {
    body.setMarginBottom(margin);
    body.setMarginTop(margin);
    body.setMarginLeft(margin);
    body.setMarginRight(margin);
    var style = { };
    style[DocumentApp.Attribute.FONT_SIZE] = 11;
    body.setAttributes(style);
  }
  
  this.addHeaderTable = function(body) {
    this.rows = [ ['Doorlopende machtiging', '', 'SEPA'] ],
    this.columns = [
        {cellWidth: 450, cellColor : '#93c47d', fontColor : '#ffffff', fontSize : 24},
        {cellWidth: 25, cellColor : '#ffffff', fontColor : '#ffffff', fontSize : 24},
        {cellWidth: 100, cellColor : '#351c75', fontColor : '#ffffff', fontSize : 24}
      ];
    this.settings = { };
    
    setFormFieldProperties();
  }
  
  this.addMainFormFieldsTable = function(body) {
    this.rows = [ 
        ['Incassant:', ' '],
        ['Naam', verenigingsnaam],
        ['Adres', adresClub],
        ['Postcode, plaats', postcodeClub + ' ' + plaatsClub],
        ['Incassant ID', incassantId],
        ['Kenmerk machtiging', kenmerk],
        ['Reden betaling', redenBetaling]
      ];
    this.columns = [
        {cellWidth : 150, fontSize : 11},
        {cellWidth : 300, fontSize : 11, fontFamily : DocumentApp.FontFamily.COURIER_NEW}
      ];
    this.settings = { };
    
    setFormFieldProperties();
    setCellProperties(table.getCell(0, 0), {bold : true});
  }
  
  this.addTextBoxTable = function(body) {
    this.rows = [
      ['Door ondertekening van dit formulier geeft u toestemming aan ' + verenigingsnaam + ' om doorlopende ' + 
      'incasso-opdrachten te sturen naar uw bank om een bedrag van uw rekening af te schijven en aan uw bank om ' +
      'doorlopend een bedrag van uw rekening af te schrijven overeenkomstig de opdracht van ' + verenigingsnaam + '.\n\n' + 
      'Als u het niet eens bent met deze afschrijving kunt u deze laten terugboeken. Neem hiervoor binnen 8 weken na ' +
      'afschrijving contact op met uw bank. Vraag uw bank naar de voorwaarden.']
    ];
    this.columns = [ { cellWidth : 575 } ];
    this.settings = { borderWidth : 1, borderColor : '#93c47d' };
    
    setFormFieldProperties();
  }
  
  this.addMemberFormFieldsTable = function () {
    this.rows = [
      ['Naam', tenNameVan],
      ['Adres', adres],
      ['Postcode, plaats', postcode + ' ' + plaats],
      ['IBAN', iban]
    ];
    this.columns = [
        {cellWidth : 150},
        {cellWidth : 300, fontFamily : DocumentApp.FontFamily.COURIER_NEW}
      ];
    this.settings = { };
    
    setFormFieldProperties();
  }
  
  this.addSignupTable = function(body) {
    this.rows = [
      ['Plaats en datum', '\n\n'],
      [' ', ''],
      ['Handtekekening', '\n\n'],
    ];
    this.columns = [
        {cellWidth : 150},
        {cellWidth : 300, underline : true, linespacingafter : 10}
    ];
    this.settings = { };
    
    setFormFieldProperties();
  }

  this.setFormFieldProperties = function() {
    this.table = body.appendTable(rows);
    table.setBorderWidth(settings.borderWidth ? settings.borderWidth : 0);
    table.setBorderColor(settings.borderColor ? settings.borderColor : '#000000');
    for (var row = 0; row < rows.length; row++)
      for (var column = 0; column < columns.length; column++)
        setCellProperties(table.getCell(row, column), columns[column]);
  }

  this.setCellProperties = function (cell, properties) {
    cell.setBackgroundColor(properties.cellColor ? properties.cellColor : '#ffffff');
    if (properties.cellWidth)
      cell.setWidth(properties.cellWidth);
    setTextProperties(cell, properties);
  }
  
  this.setTextProperties = function(cell, properties) {
    var text = cell.editAsText();
    if (text.getText().length == 0)
      return;
    setLineSpacing(cell, properties);
    text.setForegroundColor(0, text.getText().length  - 1, properties.fontColor ? properties.fontColor : '#000000' );
    text.setFontSize(properties.fontSize ? properties.fontSize : 11);
    text.setBold(properties.bold ? true : false);
    if (properties.fontFamily)
      text.setFontFamily(properties.fontFamily);
  }

  this.setLineSpacing = function(cell, properties) {
     var searchResult = null;
     var lastParagraph = null;
     while (searchResult = cell.findElement(DocumentApp.ElementType.PARAGRAPH, searchResult)) {
       var paragraph = searchResult.getElement().asParagraph();
       paragraph.setSpacingBefore(0);
       paragraph.setSpacingAfter(settings.linespacingafter ? settings.linespacingafter : 0);
       paragraph.setLineSpacing(settings.linespacing ? settings.linespacing : 1);
       lastParagraph = paragraph;
     }
     if (lastParagraph && properties.underline)
       lastParagraph.appendHorizontalRule();
  }
  
  deleteDocByName(fileName);
  this.margin = 20;
  var document = DocumentApp.create(fileName);
  var body = document.getBody();
  setDocProperties(body);
  addHeaderTable(body);
  addMainFormFieldsTable(body);
  addTextBoxTable(body);
  addMemberFormFieldsTable(body);
  addSignupTable(body);
  
  document.saveAndClose();
  
  return document;
}

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    { name : "Maak Incassobestand", functionName : "createIncassoXml" },
    { name : "Maak Machtigings PDF", functionName : "createPDF" },
  ];
  spreadsheet.addMenu("Incasso", entries);
};
