function createSignupForm(settings) {
  this.setDocProperties = function() {
    body.setMarginBottom(margin);
    body.setMarginTop(margin);
    body.setMarginLeft(margin);
    body.setMarginRight(margin);
  }
  
  this.addHeader = function() {
    this.rows = [ ['Doorlopende machtiging', '', 'SEPA'] ],
    this.columns = [
        {cellWidth: 450, cellColor : '#93c47d', fontColor : '#ffffff', fontSize : 24},
        {cellWidth: 25, cellColor : '#ffffff', fontColor : '#ffffff', fontSize : 24},
        {cellWidth: 100, cellColor : '#351c75', fontColor : '#ffffff', fontSize : 24}
      ];
    this.settings = { };
    
    setFormFieldProperties();
  }
  
  this.addMainFormFields = function() {
    this.rows = [ 
        ['Incassant:', ' '],
        ['Naam', settings.verenigingsnaam],
        ['Adres', settings.adresClub],
        ['Postcode, plaats', settings.postcodeClub + ' ' + settings.plaatsClub],
        ['Incassant ID', settings.incassantId],
        ['Kenmerk machtiging', settings.kenmerk],
        ['Reden betaling', settings.redenBetaling]
      ];
    this.columns = [
        {cellWidth : 150, fontSize : 11},
        {cellWidth : 300, fontSize : 11, fontFamily : DocumentApp.FontFamily.COURIER_NEW}
      ];
    this.settings = { };
    
    setFormFieldProperties();
    setCellProperties(table.getCell(0, 0), {bold : true});
  }

  this.addInfoBox = function() {
    this.rows = [
      ['Door ondertekening van dit formulier geeft u toestemming aan ' + settings.verenigingsnaam + ' om doorlopende ' + 
      'incasso-opdrachten te sturen naar uw bank om een bedrag van uw rekening af te schijven en aan uw bank om ' +
      'doorlopend een bedrag van uw rekening af te schrijven overeenkomstig de opdracht van ' + settings.verenigingsnaam + '.\n\n' + 
      'Als u het niet eens bent met deze afschrijving kunt u deze laten terugboeken. Neem hiervoor binnen 8 weken na ' +
      'afschrijving contact op met uw bank. Vraag uw bank naar de voorwaarden.']
    ];
    this.columns = [ { cellWidth : 575 } ];
    this.settings = { borderWidth : 1, borderColor : '#93c47d' };
    
    setFormFieldProperties();
  }
  
  this.addMemberFormFields = function () {
    this.rows = [
      ['Naam', settings.tenNameVan],
      ['Adres', settings.adres],
      ['Postcode, plaats', settings.postcode + ' ' + settings.plaats],
      ['IBAN', settings.iban]
    ];
    this.columns = [
        {cellWidth : 150},
        {cellWidth : 300, fontFamily : DocumentApp.FontFamily.COURIER_NEW}
      ];
    this.settings = { };
    
    setFormFieldProperties();
  }
  
  this.addSignup = function() {
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

  this.body = settings.document.getBody();
  this.margin = 20;
  setDocProperties();
  addHeader();
  addMainFormFields();
  addInfoBox();
  addMemberFormFields();
  addSignup();
}
