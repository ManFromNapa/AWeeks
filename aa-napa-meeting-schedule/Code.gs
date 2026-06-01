function createGoogleDocFromSheet() {
  var templateDocId = "1169KEN7t7yB9bsk6INrL6rwI1N17fFetWSypfO6Ri1w";

  // Prep regex
  var caRegex = /, CA \d{5}, USA$/;
  var napaRegex = /Napa, CA, USA$/;

  // Create a copy of the template Doc
  var templateDocFile = DriveApp.getFileById(templateDocId);
  var newDocFile = templateDocFile.makeCopy('AA Napa Meeting Schedule');
  var newDocId = newDocFile.getId();

  // Open the new Doc
  var newDoc = DocumentApp.openById(newDocId);
  var newBody = newDoc.getBody();

  // Insert the front page of the document
  var currentDate = new Date();
  var formattedDate = Utilities.formatDate(currentDate, "GMT-07:00", "MM/dd/yyyy");

  // Find the position of the specific text and insert the revision date before it
  var searchText = "For most current schedule";
  var searchResult = newBody.findText(searchText);

  if (searchResult) {
    var searchElement = searchResult.getElement();
    if (searchElement.getType() === DocumentApp.ElementType.TEXT) {
      var searchParagraph = searchElement.getParent();
      var index = newBody.getChildIndex(searchParagraph);
      var revisionParagraph = newBody.insertParagraph(index, 'Current as of ' + formattedDate);
      revisionParagraph.setFontSize(14);
      revisionParagraph.editAsText().setBold(true);
      revisionParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    } else {
      var revisionParagraph = newBody.appendParagraph('Revision - ' + formattedDate);
      revisionParagraph.setFontSize(8);
      revisionParagraph.editAsText().setBold(0, ('Revision - ' + formattedDate).length - 1, true);
    }
  }

  // Fetch the data from the Google Sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Meetings');
  var data = sheet.getDataRange().getValues();

  // Create a map using data from the Types sheet
  var typeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Types');
  var typeData = typeSheet.getRange('A:B').getValues();
  var typeMap = Object.fromEntries(typeData.map(row => [row[1], row[0]]));
  var uniqueTypesUsed = new Set();

  // Define the indices of the columns you want to include and their order
  var columnIndices = [1, 3, 5, 6, 15, 8];

  // Group data by days of the week
  var groups = {};
  for (var i = 1; i < data.length; i++) {
    var day = data[i][0];
    if (!groups[day]) groups[day] = [];
    groups[day].push(data[i]);
  }

  // Define days of the week
  var daysOfWeek = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

  // Append data day by day
  for (var day of daysOfWeek) {
    if (groups[day]) {
      var dayParagraph = newBody.appendParagraph(day);
      dayParagraph.setFontSize(10);
      dayParagraph.editAsText().setBold(0, day.length - 1, true);
      for (var i = 0; i < groups[day].length; i++) {
        var rowData = groups[day][i];
        var typesArray = rowData[8].split(', ');
        typesArray.forEach(type => uniqueTypesUsed.add(type));
        var types = typesArray.map(type => typeMap[type] || type).join(', ');
        rowData[8] = types;

        var address = rowData[6];
        address = address.replace(caRegex, "");
        address = address.replace(napaRegex, "");
        rowData[6] = address;

        var lineData = columnIndices.map(index => rowData[index]);

        if (lineData[2] === "Online Meeting") {
          lineData[2] = "Zoom";
        }

        lineData.splice(5, 0, "Types: ");

        var lineStart = lineData.slice(0, 2).join(" - ");
        var lineEnd = lineData.slice(2).join(" ");
        var line = lineStart + " - " + lineEnd;
        line = line.replace(/\s{2,}/g, ' ');

        var paragraph = newBody.appendParagraph(line).setFontSize(8);
        paragraph.editAsText().setBold(false);
        paragraph.editAsText().setBold(0, lineStart.length, true);
        paragraph.editAsText().setBold(0, 0, true);

        var column3StartIndex = line.indexOf(lineData[2]);
        var column3EndIndex = column3StartIndex + lineData[2].length;
        if (column3StartIndex >= 0 && column3EndIndex < line.length) {
          paragraph.editAsText().setBold(column3StartIndex, column3EndIndex, false);
        }

        var typesStartIndex = line.indexOf("Types: ");
        var typesEndIndex = typesStartIndex + "Types: ".length + lineData[6].length;
        if (typesStartIndex >= 0 && typesEndIndex > typesStartIndex && typesEndIndex <= line.length) {
          paragraph.editAsText().setBold(typesStartIndex, typesEndIndex - 1, true);
        }
      }
      newBody.appendParagraph('');
    }
  }

  var typeDescriptions = [];
  uniqueTypesUsed.forEach(typeCode => {
    if (typeMap[typeCode]) {
      typeDescriptions.push(typeMap[typeCode] + ": " + typeCode);
    }
  });
  typeDescriptions.sort();

  var rowsForTable = [];
  for (var i = 0; i < typeDescriptions.length; i += 2) {
    var row = [typeDescriptions[i], typeDescriptions[i+1] || ''];
    rowsForTable.push(row);
  }

  var legendTitle = newBody.appendParagraph('Meeting Legend');
  legendTitle.setFontSize(10);
  legendTitle.editAsText().setBold(true);

  var table = newBody.appendTable(rowsForTable);
  for (var r = 0; r < table.getNumRows(); r++) {
    for (var c = 0; c < table.getRow(r).getNumCells(); c++) {
      var cell = table.getCell(r, c);
      cell.editAsText().setFontSize(7);
      cell.setPaddingTop(2);
      cell.setPaddingBottom(2);
      cell.setPaddingLeft(2);
      cell.setPaddingRight(2);
    }
  }

  var body = newDoc.getBody();
  body.setPageWidth(11 * 72);
  body.setPageHeight(8.5 * 72);

  // Google Docs auto-adds an empty paragraph after a table — reuse it for the revision line
  // and zero out its spacing to avoid a blank line gap
  var currentDate = new Date();
  var formattedDate = Utilities.formatDate(currentDate, "GMT-07:00", "MM/dd/yyyy");
  var lastElement = newBody.getChild(newBody.getNumChildren() - 1);
  if (lastElement.getType() === DocumentApp.ElementType.PARAGRAPH && lastElement.asText().getText() === '') {
    var revisionParagraph = lastElement.asParagraph();
    revisionParagraph.setText('Revision - ' + formattedDate);
    revisionParagraph.setFontSize(8);
    revisionParagraph.setSpacingBefore(0);
    revisionParagraph.setSpacingAfter(0);
    revisionParagraph.editAsText().setBold(0, ('Revision - ' + formattedDate).length - 1, true);
  } else {
    var revisionParagraph = newBody.appendParagraph('Revision - ' + formattedDate);
    revisionParagraph.setFontSize(8);
    revisionParagraph.setSpacingBefore(0);
    revisionParagraph.setSpacingAfter(0);
    revisionParagraph.editAsText().setBold(0, ('Revision - ' + formattedDate).length - 1, true);
  }

  newDoc.saveAndClose();

  var newDoc = DocumentApp.openById(newDoc.getId());
  var blob = newDoc.getAs('application/pdf');

  var fileId = DriveApp.createFile(blob).getId();
  var url = 'https://docs.google.com/uc?export=download&id=' + fileId;
  Logger.log(url);

  var docFile = DriveApp.getFileById(newDoc.getId());
  var pdfFile = DriveApp.getFileById(fileId);

  docFile.setTrashed(true);
  pdfFile.setTrashed(true);
  return blob;
}

function doGet() {
  var blob = createGoogleDocFromSheet();
  var base64data = Utilities.base64Encode(blob.getBytes());
  var pdfUrl = 'data:application/pdf;base64,' + base64data;

  var htmlContent = `
    <html>
    <body>
    <p>This page was generated by the AA Print Schedule Script.  Here is the link to 
    <a href="${pdfUrl}" download="AA Napa Meeting Schedule.pdf" id="downloadLink">Download the PDF Print Schedule</a>
    <script>
      document.getElementById('downloadLink').click();
    </script>
    </body>
    </html>`;

  return HtmlService.createHtmlOutput(htmlContent);
}
