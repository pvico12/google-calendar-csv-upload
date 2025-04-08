// Color options mapped to their IDs and HTML color codes, using the provided Enum names.
const colorOptions = {
  'Peacock': { id: '1', htmlColor: '#00CED1' }, // PALE_BLUE
  'Sage': { id: '2', htmlColor: '#BCB88A' }, // PALE_GREEN
  'Grape': { id: '3', htmlColor: '#6A5ACD' }, // MAUVE
  'Flamingo': { id: '4', htmlColor: '#FC8EAC' }, // PALE_RED
  'Banana': { id: '5', htmlColor: '#FFE135' }, // YELLOW
  'Tangerine': { id: '6', htmlColor: '#FF974D' }, // ORANGE
  'Lavender': { id: '7', htmlColor: '#E6E6FA' }, // CYAN
  'Graphite': { id: '8', htmlColor: '#808080' }, // GRAY
  'Blueberry': { id: '9', htmlColor: '#4169E1' }, // BLUE
  'Basil': { id: '10', htmlColor: '#8FBC8F' }, // GREEN
  'Tomato': { id: '11', htmlColor: '#FF6347' }, // RED
};

function importOutlookCSVToCalendarWithColor() {
  // Create HTML for color buttons
  let html = '<html><body><p>Select Calendar Event Color:</p>';
  for (const colorName in colorOptions) {
    const color = colorOptions[colorName];
    html += `<button style="width: 30px; height: 30px; border-radius: 15px; background-color: ${color.htmlColor}; border: none; margin: 5px;" onclick="google.script.run.withSuccessHandler(selectColor('${colorName}'))"></button>`;
  }
  html += `<script>
    function selectColor(colorName){
      google.script.run.importOutlookCSVToCalendarWithColorContinue(colorName);
      google.script.host.close();
    }
    </script></body></html>`;

  // Display the HTML dialog
  const ui = SpreadsheetApp.getUi();
  const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(300).setHeight(200);
  ui.showModalDialog(htmlOutput, 'Select Color');
}

function selectColor(colorName) {
  PropertiesService.getScriptProperties().setProperty('selectedColor', colorName);
}

function importOutlookCSVToCalendarWithColorContinue(colorName) {  
  const selectedColor = colorName;
  const COLOR_ID = colorOptions[selectedColor].id;

  // CONFIGURATION:
  const CALENDAR_ID = 'some-calendar-id'; // Replace with your calendar ID
  const CSV_SHEET_NAME = 'Data'; // Replace with the sheet name containing your CSV data

  // Get the calendar and sheet
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CSV_SHEET_NAME);

  if (!calendar) {
    SpreadsheetApp.getUi().alert('Calendar not found. Please check your Calendar ID.');
    return;
  }

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet not found. Please check your sheet name.');
    return;
  }

  // Get the data from the sheet
  const data = sheet.getDataRange().getDisplayValues();

  // Assuming the first row is a header row, process the data from the second row onwards
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Read data from the Outlook CSV columns
    const title = row[0]; // Subject
    const startDate = row[1]; // Start Date
    const startTime = row[2]; // Start Time
    const endDate = row[3]; // End Date
    const endTime = row[4]; // End Time
    const location = row[16] || ''; // Location
    const description = row[15] || ''; // Description

    // Combine date and time to create Date objects
    const startDateTime = new Date(startDate + ' ' + startTime);
    const endDateTime = new Date(endDate + ' ' + endTime);

    // Create the event with the specified color
    try {
      targetEvent = calendar.createEvent(title, startDateTime, endDateTime, {
        description: description,
        location: location,
      });
      targetEvent.setColor(COLOR_ID);
    } catch (e) {
      SpreadsheetApp.getUi().alert('Error creating event for row ' + (i + 1) + ': ' + e.toString());
      return;
    }
  }

  SpreadsheetApp.getUi().alert('Events created successfully!');
}

// Function to create a menu for the script
function onOpen() {
  Logger.log('Creating menu option');
  try {
    SpreadsheetApp.getUi()
      .createMenu('Calendar Import')
      .addItem('Import Outlook CSV to Calendar', 'importOutlookCSVToCalendarWithColor')
      .addToUi();
  } catch (e) {
    Logger.log('Error creating menu: ' + e.toString());
  }
}
