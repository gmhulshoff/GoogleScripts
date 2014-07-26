function doGet() {
  var app = UiApp.createApplication().setTitle('EDI Contest conversie');
  createForm(app, createFormPanelForUpload(app), 1);
  return app;
}

function createFormPanelForUpload(app) {
  var form = app.createFormPanel().setId('form').setEncoding('multipart/form-data');
  app.add(form);
  return createFormContentGrid(app, form, 5, 3);
}

function createForm(app, formContent, currentRow) {
  var fileUploadBox = createFileUploadWidget(app, formContent, currentRow);
  var submitButton = createSubmitButton(app, formContent, currentRow++);
  createTestButton(app, formContent, currentRow++);
  createUploadValidation(app, fileUploadBox, submitButton); 
}

function createUploadValidation(app, fileUploadBox, submitButton) {
  var submitEnabler = app.createClientHandler().validateLength(fileUploadBox, 1, 200)
    .forTargets(submitButton).setEnabled(true);
  fileUploadBox.addChangeHandler(submitEnabler);
}

function createFormContentGrid(app, form, height, width) {
  var formContent = app.createGrid().resize(height, width);
  form.add(formContent);
  return formContent;
}

function createFileUploadWidget(app, formContent, row) {
  fileUploadBox = app.createFileUpload().setName("file");
  formContent.setWidget(row, 1, fileUploadBox);
  return fileUploadBox;
}

function createSubmitButton(app, formContent, row) {
  var submitButton = app.createSubmitButton('Converteer').setId('convert').setEnabled(false);
  formContent.setWidget(row, 2, submitButton);
  return submitButton;
}

function createTestButton(app, formContent, row) {
  var buttonShowExample = app.createButton('Voorbeeld', app.createServerHandler('showExample'));
  formContent.setWidget(row, 1, buttonShowExample);
}

function showExample(event) {
  var app = UiApp.getActiveApplication();
  var testData = '[REG1TEST;1]\n' + 
  'TName=2011 IARU R1 VHF/UHF Contest\n' + 
  'TDate=20111016;20111016\n\n' +
  '[QSORecords;2]\n111016;0901;SK7MW;2;;500;55;56;;JO65MJ;550;;N;N;\n' +
  '111016;1914;PE1RLF;2;;500;58;57;;JO32CG;92;;N;N;';

  showPopup(app, testData);
 
  return app;
}

function showPopup(app, testData) {
  var popup = app.createPopupPanel(true, true);
  var html = app.createHTML(
    '<pre>Voor conversie:\n\n' + 
    testData + 
	'\n\nNa conversie:\n\n' +
	parseContent(testData) + 
	'\n\n(klik buiten dit gebied om deze popup te sluiten)' +
	'</pre>');
  popup.add(html);
  popup.setGlassEnabled(true);
  popup.setPopupPosition(100, 100);
  popup.show();
}
 
function doPost(e) {
  var blob = e.parameter.file;
  var content = blob.getDataAsString();
  var app = UiApp.getActiveApplication();
  
  app.add(app.createHTML('<pre>' + parseContent(content) + '</pre>'));
  
  return app;
}

function parseContent(content) {
  return content
	.replace(/(([^;]*?;){3})(2;)(([^;]*?;){10})/g, "$1<replaceme>$4")
	.replace(/<replaceme>/g,"1;");
}
