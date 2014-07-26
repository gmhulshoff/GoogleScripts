function importSpreadsheet() {
  this.loadZippedJson = function () {
    var spreadSheetName = spreadSheet.getName() + "Export.zip";
    var files = DriveApp.getFilesByName(spreadSheetName);
    if (!files.hasNext()) return null;
    var file = files.next();
    var blob = file.getBlob();
    var unzippedFiles = Utilities.unzip(blob);
    if (unzippedFiles.length == 0) return null;
    var json = unzippedFiles[0].getAs("application/json").getDataAsString();
    return JSON.parse(json);
  }
  
  this.getOrCreateImport = function () {
    var rootFolder = DocsList.getRootFolder();
    var name = spreadSheet.getName() + "Import";
    var files = rootFolder.find(name);
    for (var i in files)
      return SpreadsheetApp.open(files[i]);
    return SpreadsheetApp.create(name);
  }

  this.importCell = function (cell) {
    var range = newSheet.getRange(cell.adress)
      .setNumberFormat(cell.numberFormat)
      .setNote(cell.note)
      .setHorizontalAlignment(cell.horizontalAlign)
      .setVerticalAlignment(cell.verticalAlign)
      .setFontSize(cell.fontSize)
      .setFontWeight(cell.fontWeight)
      .setFontFamily(cell.fontFamily)
      .setFontColor(cell.color)
      .setBackground(cell.backGround)
      .setWrap(cell.wrap == "true");
    if (cell.formula != "")
      range.setFormula(cell.formula);
    else
      range.setValue(cell.value);
  }
  
  this.getCleanSheet = function (name) {
    var newSheet = spreadSheet.getSheetByName(name);
    if (newSheet == null)
      return spreadSheet.insertSheet(name);
    newSheet.clearContents();
    newSheet.clearFormats();
    newSheet.clearNotes();
    return newSheet;
  }
  
  this.importNamedRanges = function () {
    for (var rangeIndex in sheetContents.namedRanges) {
      var newRange = sheetContents.namedRanges[rangeIndex];
      spreadSheet.setNamedRange(newRange.name, spreadSheet.getRange(newRange.value));
    }
  }
  
  this.importCells = function() {
    for (var cellIndex in sheet.cells)
      importCell(sheet.cells[cellIndex]);
  }
  
  this.setColumnWidths = function () {
    for (var colIndex = 1; colIndex <= newSheet.getLastColumn(); colIndex++)
      newSheet.setColumnWidth(colIndex, sheet.columnWidths[colIndex - 1]);
  }
  
  this.setColumnHeights = function () {
    for (var rowIndex = 1; rowIndex <= newSheet.getLastRow(); rowIndex++)
      newSheet.setRowHeight(rowIndex, sheet.rowHeights[rowIndex - 1]);
  }
  
  this.importSheets = function() {
    for (var sheetIndex in sheetContents.sheets) {
      this.sheet = sheetContents.sheets[sheetIndex];
      this.newSheet = getCleanSheet(sheet.name);
      newSheet.activate();
      newSheet.setFrozenRows(sheet.frozenRows);
      newSheet.setFrozenColumns(sheet.frozenColumns);
      importCells();
      setColumnWidths();
      setColumnHeights();
    }
  }
  
  this.importSpreadsheet = function () {
    this.spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    this.sheetContents = loadZippedJson();
    if (sheetContents == null) return;
    this.spreadSheet = getOrCreateImport();
    var removeMe = spreadSheet.getActiveSheet().setName('RemoveMe');
  
    spreadSheet.setSpreadsheetTimeZone(sheetContents.timezone);
    spreadSheet.setSpreadsheetLocale(sheetContents.locale);
  
    importSheets();
    importNamedRanges();

    spreadSheet.deleteSheet(removeMe);
  }
  
  importSpreadsheet();
}

function exportSpreadsheet() {
  this.saveZippedJson = function (data) {
    var spreadSheetName = spreadSheet.getName();
    var blob = Utilities.newBlob(data, "application/json", spreadSheetName + ".json");
    var zip = Utilities.zip([blob]);
    var file = DriveApp.createFile(zip);
    file.setName(spreadSheetName + "Export.zip");
    Logger.log(file);
  }
  
  this.localizeDates = function (i, j) {
    if (typeof(sheetContents.values[i][j]) == "object")
      sheetContents.values[i][j] = Utilities.formatDate(new Date(sheetContents.values[i][j]), spreadSheet.getSpreadsheetTimeZone(), "yyyy-MM-dd");
  }

  this.localizeNumbers = function (i, j) {
    var nr;
    if (export.locale == "nl_NL" && (nr = /(\d{3},)*\d{0,3}\.\d+/.exec(sheetContents.values[i][j])))
      sheetContents.values[i][j] = nr[0].replace(".","!").replace(",", ".").replace("!",",");
  }

  this.resolveNamedRangesInFormula = function (i, j) {
    if (sheetContents.formulas[i][j].length > 0)
      getNamedRanges(sheetContents.formulas[i][j]);
  }

  this.getNamedRanges = function (formula) {
    var regex = new RegExp("(([a-zA-Z]+)(\\)|;|>|=|<))", "g");
    var match, range;
    while (match = regex.exec(formula))
      if (rangeNames[match[2]] == null && (range = spreadSheet.getRangeByName(match[2])) != null)
        rangeNames[match[2]] = range.getSheet().getName() + "!" + range.getA1Notation();
  }
  
  this.getSheetContents = function (rows) {
    return { 
	  values : rows.getValues(), 
	  formulas : rows.getFormulas(),
	  formats : rows.getNumberFormats(),
	  colors : rows.getFontColors(),
	  families : rows.getFontFamilies(),
	  wraps : rows.getWraps(),
	  horAlignemnts : rows.getHorizontalAlignments(),
	  verAlignments : rows.getVerticalAlignments(),
	  fontSizes : rows.getFontSizes(),
	  fontStyles : rows.getFontStyles(),
	  fontWeights : rows.getFontWeights(),
	  fontlines : rows.getFontLines(),
	  notes : rows.getNotes(),
	  backGrounds : rows.getBackgrounds(),
    };
  }

  this.getCells = function () {
    this.sheetContents = getSheetContents(sheet.getDataRange())
    this.cells = [];
    for (var i in sheetContents.values)
      for (var j in sheetContents.values[i])
        cells.push(getCellData(i, j));
    return cells;
  }
  
  this.getColumnWidths = function () {
    var widths = []
    for (var col = 1; col <= sheet.getLastColumn(); col++)
      widths.push(sheet.getColumnWidth(col));
    return widths;
  }

  this.getRowHeights = function () {
    var heights = [];
    for (var row = 1; row <= sheet.getLastRow(); row++)
      heights.push(sheet.getRowHeight(row));
    return heights;
  }

  this.getCellData = function (i, j) {
    localizeDates(i, j);
    localizeNumbers(i, j);
    resolveNamedRangesInFormula(i, j);
  
    return { 
	  adress : sheet.getRange(eval(i) + 1, eval(j) + 1).getA1Notation(), 
	  value : sheetContents.values[i][j].toString(),
	  formula : sheetContents.formulas[i][j],
	  numberFormat : sheetContents.formats[i][j],
	  horizontalAlign : sheetContents.horAlignemnts[i][j],
	  verticalAlign : sheetContents.verAlignments[i][j],
	  fontSize : sheetContents.fontSizes[i][j],
	  fontWeight : sheetContents.fontWeights[i][j],
	  fontStyle : sheetContents.fontStyles[i][j],
	  fontFamily : sheetContents.families[i][j],
	  wrap : sheetContents.wraps[i][j].toString(),
	  color : sheetContents.colors[i][j],
	  backGround : sheetContents.backGrounds[i][j],
	  note : sheetContents.notes[i][j]
    };
  }
  
  this.exportSheets = function() {
    for (var i in sheets) {
      this.sheet = sheets[i];
      export.sheets.push({ 
        name : sheet.getName(), 
        columnWidths : getColumnWidths(sheet), 
        rowHeights : getRowHeights(sheet), 
        frozenRows : sheet.getFrozenRows(), 
        frozenColumns : sheet.getFrozenColumns(), 
        cells : getCells()
      });
    }
  }
  
  this.exportNamedRanges = function() {
    for (var i in namedRanges)
      export.namedRanges.push({ name : i, value : namedRanges[i]} );
  }

  this.exportSpreadsheet = function () {
    this.spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    this.sheets = spreadSheet.getSheets();
    this.export = { locale : spreadSheet.getSpreadsheetLocale(), timezone : spreadSheet.getSpreadsheetTimeZone(), namedRanges : [], sheets : [] };
    this.namedRanges = [];

	exportSheets();
	exportNamedRanges();
    
    saveZippedJson(JSON.stringify(export,undefined,"\t"));
  }
  
  exportSpreadsheet();
}
