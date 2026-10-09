function doGet(e) {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return ContentService
      .createTextOutput(JSON.stringify({
        records: []
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();

  var records = data.map(function(row) {

    var age = row[1];

    var ageGroup = "";

    if (age < 20) {
      ageGroup = "Under 20";
    } else if (age >= 20 && age <= 29) {
      ageGroup = "20-29";
    } else if (age >= 30 && age <= 39) {
      ageGroup = "30-39";
    } else if (age >= 40 && age <= 49) {
      ageGroup = "40-49";
    } else if (age >= 50 && age <= 59) {
      ageGroup = "50-59";
    } else if (age >= 60) {
      ageGroup = "60+";
    }

    return {
      name: row[0],
      age: row[1],
      ageGroup: ageGroup,
      gender: row[3],
      profession: row[2]
    };

  }).filter(function(record) {

    return record.name &&
           record.name.toString().trim() !== "";

  });

  return ContentService
    .createTextOutput(JSON.stringify({
      records: records
    }))
    .setMimeType(ContentService.MimeType.JSON);
}