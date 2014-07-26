function makeNotes() {
  this.readDates = function (sheetName, dateCol, noteCol) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    var rows = sheet.getDataRange();
    var values = rows.getValues();
    var backgrounds = rows.getBackgrounds();
    var dates = new Array();
  
    for (var i in values)
      dates.push({ date: new Date(values[i][dateCol]), note : values[i][noteCol], background : backgrounds[i][noteCol] });

    return dates;
  }
  
  this.markCell = function (cell, contents, data) {
    if (data.note == "") return contents;
    contents = contents + (contents == "" ? "" : ", ") + data.note;
    cell.setComment(data.note);
    cell.setBackground(data.background);
    return contents;
  }

  this.addDays = function (date, days) {
    return new Date(date.getFullYear(),date.getMonth(),date.getDate() + days);
  }

  this.getAppointment = function (appointments, date) {
    for (var i in appointments)
      if (appointments[i].date.toLocaleDateString() == date.toLocaleDateString())
        return appointments[i];
    return { note : "", backgronund : "white", date : date };
  }

  this.makeNotes = function () {
    var sheet = SpreadsheetApp.getActiveSheet();
    var rows = sheet.getDataRange();
    var values = rows.getValues();
  
    var appointments = readDates("Afspraken", 0, 5);
    var holidays = readDates("Feestdagen", 0, 3);
    var birthdays = readDates("Verjaardagen", 0, 3);

    var startDate = new Date(values[0][0]);
    for (var i in values) {
      if (i > 1 && i <= 53) {
        var weekNo = values[i - 1][1];
        var day = addDays(startDate, 7 * (weekNo - 1));
        for (var j in values[i]) {
          if (j > 2 && j < 10) {
            var cell = sheet.getRange(i, j);
            var weekDay = addDays(day, j - 3);
            var contents = markCell(cell, "", getAppointment(appointments, weekDay));
            contents = markCell(cell, contents, getAppointment(birthdays, weekDay));
            contents = markCell(cell, contents, getAppointment(holidays, weekDay));
          }
        }
      }
    }
  }
  
  makeNotes();
}

