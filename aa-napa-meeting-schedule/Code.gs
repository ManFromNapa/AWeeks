/**
 * AA Napa Meeting Schedule Generator
 *
 * Reads meeting data from the 'Meetings' sheet and meeting type codes from
 * the 'Types' sheet in the active Google Spreadsheet, then generates a
 * formatted PDF meeting schedule based on a Google Doc template.
 *
 * Expected sheets:
 *   - 'Meetings': One row per meeting, with day, time, name, address, type codes, etc.
 *   - 'Types': Two columns — col A: full type name, col B: type code abbreviation
 *
 * Output: A PDF blob served as a download via the doGet() web app endpoint.
 */
function createGoogleDocFromSheet() {
  var templateDocId = "1169KEN7t7yB9bsk6INrL6rwI1N17fFetWSypfO6Ri1w";

  // Regex patterns to strip redundant location suffixes from addresses
  // e.g. ", CA 94559, USA" and "Napa, CA, USA" are removed for brevity
  var caRegex = /, CA \d{5}, USA$/;
  var napaRegex = /Napa, CA, USA$/;

  // Copy the template Doc so the original is never modified
  var templateDocFile = DriveApp.getFileById(templateDocId);
  var newDocFile = templateDocFile.makeCopy('AA Napa Meeting Schedule');
  var newDocId = newDocFile.getId();

  // Open the new Doc and get its body for editing
  var newDoc = DocumentApp.openById(newDocId);
  var newBody = newDoc.getBody();

  // Format today's date in Pacific time for the "Current as of" header
  var currentDate = new Date();
  var formattedDate = Utilities.formatDate(currentDate, "GMT-07:00", "MM/dd/yyyy");

  // Find the "For most current schedule" placeholder in the template and insert
  // the current date just before it as a bold centered heading
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

  // Read all meeting rows from the 'Meetings' sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Meetings');
  var data = sheet.getDataRange().getValues();

  // Read type code mappings from the 'Types' sheet
  // col B (index 1) = code abbreviation (e.g. "DR"), col A (index 0) = full name (e.g. "Daily Reflections")
  // typeMap is keyed by code: { "DR": "Daily Reflections", ... }
  var typeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Types');
  var typeData = typeSheet.getRange('A:B').getValues();
  var typeMap = Object.fromEntries(typeData.map(row => [row[1], row[0]]));
  var uniqueTypesUsed = new Set(); // Tracks which type codes appear in this schedule for the legend

  // Column indices from the Meetings sheet to include in each printed line:
  // 1=Time, 3=Meeting Name, 5=Location Name, 6=Address, 15=Zoom/Notes, 8=Type Codes
  var columnIndices = [1, 3, 5, 6, 15, 8];

  // Group meeting rows by day of week (column 0)
  var groups = {};
  for (var i = 1; i < data.length; i++) {
    var day = data[i][0];
    if (!groups[day]) groups[day] = [];
    groups[day].push(data[i]);
  }

  // Process days in calendar order
  var daysOfWeek = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

  for (var day of daysOfWeek) {
    if (groups[day]) {
      // Print the day name as a bold 10pt heading
      var dayParagraph = newBody.appendParagraph(day);
      dayParagraph.setFontSize(10);
      dayParagraph.editAsText().setBold(0, day.length - 1, true);

      for (var i = 0; i < groups[day].length; i++) {
        var rowData = groups[day][i];

        // Collect type codes used across all meetings for the legend at the end
        var typesArray = rowData[8].split(', ');
        typesArray.forEach(type => uniqueTypesUsed.add(type));

        // Replace type code abbreviations with their full names for display
        var types = typesArray.map(type => typeMap[type] || type).join(', ');
        rowData[8] = types;

        // Strip redundant location suffixes from the address field
        var address = rowData[6];
        address = address.replace(caRegex, "");
        address = address.replace(napaRegex, "");
        rowData[6] = address;

        // Extract only the columns we want to print, in the defined order
        var lineData = columnIndices.map(index => rowData[index]);

        // Normalize online meeting location label
        if (lineData[2] === "Online Meeting") {
          lineData[2] = "Zoom";
        }

        // Insert the "Types: " label before the type codes in the line
        lineData.splice(5, 0, "Types: ");

        // lineStart = "Time - Meeting Name" (printed in bold)
        // lineEnd = everything else (location, address, notes, types)
        var lineStart = lineData.slice(0, 2).join(" - ");
        var lineEnd = lineData.slice(2).join(" ");
        var line = lineStart + " - " + lineEnd;
        line = line.replace(/\s{2,}/g, ' '); // Collapse any double spaces

        // Append the meeting line, then bold only the time and name portion
        var text = newBody.appendParagraph(line).setFontSize(8).editAsText();
        text.setBold(false);
        text.setBold(0, lineStart.length, true);
      }
      // Add a blank line between days
      newBody.appendParagraph('');
    }
  }

  // Build the sorted list of type descriptions used in this schedule
  var typeDescriptions = [];
  uniqueTypesUsed.forEach(typeCode => {
    if (typeMap[typeCode]) {
      typeDescriptions.push(typeMap[typeCode] + ": " + typeCode);
    }
  });
  typeDescriptions.sort();

  // Pair descriptions into 2-column rows to save vertical space in the legend table
  var rowsForTable = [];
  for (var i = 0; i < typeDescriptions.length; i += 2) {
    var row = [typeDescriptions[i], typeDescriptions[i+1] || ''];
    rowsForTable.push(row);
  }

  // Append the legend title and 2-column table
  var legendTitle = newBody.appendParagraph('Meeting Legend');
  legendTitle.setFontSize(10);
  legendTitle.editAsText().setBold(true);

  var table = newBody.appendTable(rowsForTable);

  // Set font size and tight padding on all table cells to minimize space
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

  // Set page to landscape orientation (11" x 8.5")
  var body = newDoc.getBody();
  body.setPageWidth(11 * 72);
  body.setPageHeight(8.5 * 72);

  // Google Docs automatically appends an empty paragraph after every table.
  // Rather than adding another paragraph, reuse that one for the revision line
  // and zero out its spacing to keep it tight against the table.
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

  // Export the Doc as a PDF blob
  var newDoc = DocumentApp.openById(newDoc.getId());
  var blob = newDoc.getAs('application/pdf');

  // Create a temporary Drive file to get a download URL (logged for debugging)
  var fileId = DriveApp.createFile(blob).getId();
  var url = 'https://docs.google.com/uc?export=download&id=' + fileId;
  Logger.log(url);

  // Trash both the temporary Doc and PDF file — the blob is returned directly
  var docFile = DriveApp.getFileById(newDoc.getId());
  var pdfFile = DriveApp.getFileById(fileId);
  docFile.setTrashed(true);
  pdfFile.setTrashed(true);

  return blob;
}

/**
 * Web app entry point. Calls createGoogleDocFromSheet(), encodes the resulting
 * PDF as base64, and serves an HTML page that auto-triggers a file download.
 */
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
