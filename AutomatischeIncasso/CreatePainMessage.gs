function doGet() {
  this.createFormPanel = function () {
    var form = app.createFormPanel().setId('form').setEncoding('multipart/form-data');
    var formContent = app.createGrid().resize(10, 3);
    app.add(form);
    form.add(formContent);
    return formContent;
  }

  var app = UiApp.createApplication().setTitle('AutomatischeIncasso');
  createForm(app, createFormPanel(), 1);
  return app;
}

function doPost(event) {
  this.renderXml = function (xml, xmlMatches) {
    app.remove(0);
    app.add(app.createHTML('<pre>\n\n' + xml.replace(/</g, '&lt;') + '</pre>').setStyleAttribute("color", xmlMatches ? "green" : "red"));
  }

  var blob = event.parameter.file;
  var content = blob.getDataAsString();
  var app = UiApp.getActiveApplication();
  
  var data = Utilities.parseCsv(content);
  var header = data.shift();
  var debtors = [];
  for (var i = 0, record; record = data[i]; i++)
	debtors.push({ instdamt : record[0], mndtid : record[1], seqtp : record[2], 
	dtofsgntr : record[3], amdmntind : record[4],
	dbtrnm : record[5], dbtriban : record[6], dbtrustrd : record[7]});

  event.parameter.creationdatetime = '2014-07-12T19:07:11';
  var xml = CreatePain00800102.createPainMessage(debtors, event);
  var checkXml = getCheckXml(event);
  
  renderXml(xml, checkXml.replace(/\s/g, '') == xml.replace(/\s/g, ''));
  
  return app;
}

function createForm(app, formContent, currentRow) {
  this.createFileUploadWidget = function () { return app.createFileUpload().setName("file"); }
  this.createSubmitButton = function () { return app.createSubmitButton('Maak incassobestand').setId('convert').setEnabled(false); }
  this.createIncassoButton = function () { return app.createButton('Incasso XML', handler).setId('incassoButton'); }
  this.createTextBox = function (name, text, label, row) {
    formContent.setWidget(row, 1, app.createTextBox().setName(name).setText(text));
    formContent.setWidget(row, 2, app.createLabel(label));
  }
  this.createUploadEnableWhenFileSelected = function () {
    var submitEnabler = app
      .createClientHandler()
      .validateLength(fileUploadBox, 1, 200)
      .forTargets(submitButton)
      .setEnabled(true);
    fileUploadBox.addChangeHandler(submitEnabler);
  }

  var fileUploadBox = createFileUploadWidget();
  var submitButton = createSubmitButton();
  formContent.setWidget(currentRow, 1, fileUploadBox);
  formContent.setWidget(currentRow++, 2, submitButton);

  var handler = app.createServerHandler('getIncassoXml').addCallbackElement(formContent);
  formContent.setWidget(currentRow++, 1, createIncassoButton());

  createTextBox('initiatingparty', 'Oer \'t net', '(Verenigings-naam)', currentRow++);
  createTextBox('initiatingiban', 'NL18INGB0000825457', '(Verenigins-iban)', currentRow++);
  createTextBox('initiatingbic', 'INGBNL2A', '(Verenigins-bic)', currentRow++);
  createTextBox('initiatingcreditoridentifier', 'NL97ZZZ401029480000', '(Verenigins creditor identifier)', currentRow++);
  createTextBox('incassodate', '2014-01-06', '(Incassodatum)', currentRow++);
  createTextBox('endtoendid', 'D026', '(Betalingskenmerk)', currentRow++);
  
  createUploadEnableWhenFileSelected();
}

function getIncassoXml(event) {
  this.showPopup = function (testData) {
    var popup = app.createPopupPanel(true, true);
    popup.add(app.createHTML('<pre>\n\n' + testData.replace(/</g, '&lt;') + '</pre>').setStyleAttribute("color", xmlMatches ? "green" : "red"));
    popup.setGlassEnabled(true);
    popup.setPopupPosition(150, 120);
    popup.show();
  }

  var app = UiApp.getActiveApplication();
  var debtors = [
	{instdamt: "1.00", mndtid: "ETT00010v1", seqtp: "FRST", dtofsgntr: "2013-12-11", amdmntind: "false", 
		dbtrnm: "D. Biteur", dbtriban: "NL13TEST0123456789", dbtrustrd: "Dit is testbericht 1"}, 
	{instdamt: "2.00", mndtid: "ETT00011v1", seqtp: "RCUR", dtofsgntr: "2013-12-12", amdmntind: "false",
		dbtrnm: "D. Biteur", dbtriban: "NL13TEST0123456789", dbtrustrd: "Dit is testbericht 2"},
	{instdamt: "3.00", mndtid: "ETT00011v1", seqtp: "RCUR", dtofsgntr: "2013-12-12", amdmntind: "false",
		dbtrnm: "QreaCom", dbtriban: "NL13TEST0123456789", dbtrustrd: "Dit is testbericht 3"} 
  ];

  event.parameter.creationdatetime = '2014-07-12T19:07:11';
  var xml = CreatePain00800102.createPainMessage(debtors, event);
  var xmlMatches = getCheckXml(event).replace(/\s/g, '') == xml.replace(/\s/g, '');
  var button = app.getElementById('incassoButton');
  
  showPopup(xml);
  
  return app;
}

function getCheckXml(event) {
  return '<?xml version="1.0" encoding="UTF-8"?>' +
	'<Document xmlns="urn:iso:std:iso:20022:tech:xsd:pain.008.001.02" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
	'  <CstmrDrctDbtInitn>' +
	'    <GrpHdr>' +
	'      <MsgId>D0262014-07-12T19:07:11</MsgId>' +
	'      <CreDtTm>2014-07-12T19:07:11</CreDtTm>' +
	'      <NbOfTxs>3</NbOfTxs>' +
	'      <InitgPty>' +
	'        <Nm>' + event.parameter.initiatingparty + '</Nm>' +
	'      </InitgPty>' +
	'    </GrpHdr>' +
	'    <PmtInf>' +
	'      <PmtInfId>' + event.parameter.endtoendid + '-FRST</PmtInfId>' +
	'      <PmtMtd>DD</PmtMtd>' +
	'      <PmtTpInf>' +
	'        <SvcLvl>' +
	'          <Cd>SEPA</Cd>' +
	'        </SvcLvl>' +
	'        <LclInstrm>' +
	'          <Cd>CORE</Cd>' +
	'        </LclInstrm>' +
	'        <SeqTp>FRST</SeqTp>' +
	'      </PmtTpInf>' +
	'      <ReqdColltnDt>2014-01-06</ReqdColltnDt>' +
	'      <Cdtr>' +
	'        <Nm>' + event.parameter.initiatingparty + '</Nm>' +
	'        <PstlAdr>' +
	'          <Ctry>NL</Ctry>' +
	'        </PstlAdr>' +
	'      </Cdtr>' +
	'      <CdtrAcct>' +
	'        <Id>' +
	'          <IBAN>' + event.parameter.initiatingiban + '</IBAN>' +
	'        </Id>' +
	'      </CdtrAcct>' +
	'      <CdtrAgt>' +
	'        <FinInstnId>' +
	'          <BIC>' + event.parameter.initiatingbic + '</BIC>' +
	'        </FinInstnId>' +
	'      </CdtrAgt>' +
	'      <UltmtCdtr>' +
	'        <Nm>' + event.parameter.initiatingparty + '</Nm>' +
	'      </UltmtCdtr>' +
	'      <ChrgBr>SLEV</ChrgBr>' +
	'      <CdtrSchmeId>' +
	'        <Id>' +
	'          <PrvtId>' +
	'            <Othr>' +
	'              <Id>' + event.parameter.initiatingcreditoridentifier + '</Id>' +
	'              <SchmeNm>' +
	'                <Prtry>SEPA</Prtry>' +
	'              </SchmeNm>' +
	'            </Othr>' +
	'          </PrvtId>' +
	'        </Id>' +
	'      </CdtrSchmeId>' +
	'      <DrctDbtTxInf>' +
	'        <PmtId>' +
	'          <EndToEndId>D026-ETT00010v1</EndToEndId>' +
	'        </PmtId>' +
	'        <InstdAmt Ccy="EUR">1.00</InstdAmt>' +
	'        <DrctDbtTx>' +
	'          <MndtRltdInf>' +
	'            <MndtId>ETT00010v1</MndtId>' +
	'            <DtOfSgntr>2013-12-11</DtOfSgntr>' +
	'            <AmdmntInd>false</AmdmntInd>' +
	'          </MndtRltdInf>' +
	'        </DrctDbtTx>' +
	'        <DbtrAgt>' +
	'          <FinInstnId />' +
	'        </DbtrAgt>' +
	'        <Dbtr>' +
	'          <Nm>D. Biteur</Nm>' +
	'          <PstlAdr>' +
	'            <Ctry>NL</Ctry>' +
	'          </PstlAdr>' +
	'        </Dbtr>' +
	'        <DbtrAcct>' +
	'          <Id>' +
	'            <IBAN>NL13TEST0123456789</IBAN>' +
	'          </Id>' +
	'        </DbtrAcct>' +
	'        <UltmtDbtr>' +
	'          <Nm>D. Biteur</Nm>' +
	'        </UltmtDbtr>' +
	'        <RmtInf>' +
	'          <Ustrd>Dit is testbericht 1</Ustrd>' +
	'        </RmtInf>' +
	'      </DrctDbtTxInf>' +
	'    </PmtInf>' +
	'    <PmtInf>' +
	'      <PmtInfId>' + event.parameter.endtoendid + '-RCUR</PmtInfId>' +
	'      <PmtMtd>DD</PmtMtd>' +
	'      <PmtTpInf>' +
	'        <SvcLvl>' +
	'          <Cd>SEPA</Cd>' +
	'        </SvcLvl>' +
	'        <LclInstrm>' +
	'          <Cd>CORE</Cd>' +
	'        </LclInstrm>' +
	'        <SeqTp>RCUR</SeqTp>' +
	'      </PmtTpInf>' +
	'      <ReqdColltnDt>2014-01-06</ReqdColltnDt>' +
	'      <Cdtr>' +
	'        <Nm>' + event.parameter.initiatingparty + '</Nm>' +
	'        <PstlAdr>' +
	'          <Ctry>NL</Ctry>' +
	'        </PstlAdr>' +
	'      </Cdtr>' +
	'      <CdtrAcct>' +
	'        <Id>' +
	'          <IBAN>' + event.parameter.initiatingiban + '</IBAN>' +
	'        </Id>' +
	'      </CdtrAcct>' +
	'      <CdtrAgt>' +
	'        <FinInstnId>' +
	'          <BIC>' + event.parameter.initiatingbic + '</BIC>' +
	'        </FinInstnId>' +
	'      </CdtrAgt>' +
	'      <UltmtCdtr>' +
	'        <Nm>' + event.parameter.initiatingparty + '</Nm>' +
	'      </UltmtCdtr>' +
	'      <ChrgBr>SLEV</ChrgBr>' +
	'      <CdtrSchmeId>' +
	'        <Id>' +
	'          <PrvtId>' +
	'            <Othr>' +
	'              <Id>' + event.parameter.initiatingcreditoridentifier + '</Id>' +
	'              <SchmeNm>' +
	'                <Prtry>SEPA</Prtry>' +
	'              </SchmeNm>' +
	'            </Othr>' +
	'          </PrvtId>' +
	'        </Id>' +
	'      </CdtrSchmeId>' +
	'      <DrctDbtTxInf>' +
	'        <PmtId>' +
	'          <EndToEndId>D026-ETT00011v1</EndToEndId>' +
	'        </PmtId>' +
	'        <InstdAmt Ccy="EUR">2.00</InstdAmt>' +
	'        <DrctDbtTx>' +
	'          <MndtRltdInf>' +
	'            <MndtId>ETT00011v1</MndtId>' +
	'            <DtOfSgntr>2013-12-12</DtOfSgntr>' +
	'            <AmdmntInd>false</AmdmntInd>' +
	'          </MndtRltdInf>' +
	'        </DrctDbtTx>' +
	'        <DbtrAgt>' +
	'          <FinInstnId />' +
	'        </DbtrAgt>' +
	'        <Dbtr>' +
	'          <Nm>D. Biteur</Nm>' +
	'          <PstlAdr>' +
	'            <Ctry>NL</Ctry>' +
	'          </PstlAdr>' +
	'        </Dbtr>' +
	'        <DbtrAcct>' +
	'          <Id>' +
	'            <IBAN>NL13TEST0123456789</IBAN>' +
	'          </Id>' +
	'        </DbtrAcct>' +
	'        <UltmtDbtr>' +
	'          <Nm>D. Biteur</Nm>' +
	'        </UltmtDbtr>' +
	'        <RmtInf>' +
	'          <Ustrd>Dit is testbericht 2</Ustrd>' +
	'        </RmtInf>' +
	'      </DrctDbtTxInf>' +
	'      <DrctDbtTxInf>' +
	'        <PmtId>' +
	'          <EndToEndId>D026-ETT00011v1</EndToEndId>' +
	'        </PmtId>' +
	'        <InstdAmt Ccy="EUR">3.00</InstdAmt>' +
	'        <DrctDbtTx>' +
	'          <MndtRltdInf>' +
	'            <MndtId>ETT00011v1</MndtId>' +
	'            <DtOfSgntr>2013-12-12</DtOfSgntr>' +
	'            <AmdmntInd>false</AmdmntInd>' +
	'          </MndtRltdInf>' +
	'        </DrctDbtTx>' +
	'        <DbtrAgt>' +
	'          <FinInstnId />' +
	'        </DbtrAgt>' +
	'        <Dbtr>' +
	'          <Nm>QreaCom</Nm>' +
	'          <PstlAdr>' +
	'            <Ctry>NL</Ctry>' +
	'          </PstlAdr>' +
	'        </Dbtr>' +
	'        <DbtrAcct>' +
	'          <Id>' +
	'            <IBAN>NL13TEST0123456789</IBAN>' +
	'          </Id>' +
	'        </DbtrAcct>' +
	'        <UltmtDbtr>' +
	'          <Nm>QreaCom</Nm>' +
	'        </UltmtDbtr>' +
	'        <RmtInf>' +
	'          <Ustrd>Dit is testbericht 3</Ustrd>' +
	'        </RmtInf>' +
	'      </DrctDbtTxInf>' +
	'    </PmtInf>' +
	'  </CstmrDrctDbtInitn>' +
	'</Document>';
}
