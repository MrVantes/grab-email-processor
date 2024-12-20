function getEmailsFromSpecificSenderToday() {
  const senderEmail = "no-reply@example.com"; // Replace with generic email
  const today = new Date();
  const timezone = "GMT+7";

  const formattedDate = Utilities.formatDate(
    new Date(today.getFullYear(), today.getMonth(), today.getDate()),
    timezone,
    "yyyy/MM/dd"
  );

  const todayBandung = new Date(formattedDate);

  const months = [
    "JAN",
    "FEB",
    "MAR",
    "APR",
    "MAY",
    "JUN",
    "JUL",
    "AUG",
    "SEP",
    "OCT",
    "NOV",
    "DEC",
  ];
  const month = months[todayBandung.getMonth()];

  const query = `from:${senderEmail} after:` + formattedDate;

  const threads = GmailApp.search(query);
  Logger.log("Threads found: " + threads.length);

  let addRow = 0;

  const sheet = SpreadsheetApp.openById(
    "your-spreadsheet-id" // Replace with actual ID
  ).getSheetByName(`${month} ${todayBandung.getFullYear()}`);

  if (threads.length === 0) {
    const date = Utilities.formatDate(
      new Date(formattedDate),
      timezone,
      "MM/dd/yyyy"
    );
    findAndEnterData(
      sheet,
      date,
      0,
      1,
      "Work",
      "GenericLocation1",
      "GenericLocation2",
      "N/A",
      "N/A",
      "N/A"
    );
    findAndEnterData(
      sheet,
      date,
      0,
      0,
      "Work",
      "GenericLocation2",
      "GenericLocation1",
      "N/A",
      "N/A",
      "N/A"
    );
  } else {
    threads.forEach((thread) => {
      // Get the messages and reverse the order
      const messages = thread.getMessages().reverse();

      messages.forEach((message) => {
        var sender = message.getFrom();
        var subject = message.getSubject();
        var date = message.getDate();
        var body = message.getPlainBody();
        const formattedEmailDate = Utilities.formatDate(
          date,
          timezone,
          "MM/dd/yyyy"
        );
        const formattedEmailTime = Utilities.formatDate(
          date,
          timezone,
          "HH:mm:ss"
        );
        Logger.log(formattedEmailTime);

        if (body.startsWith("Receipt")) {
          const money = extractTotalPaidAmountFromEmail(body);
          const location = extractLocationFromEmail(body);
          const loc1 = simplifyText(location[0]);
          const loc2 = simplifyText(location[1]);
          const car = simplifyText(extractCarFromEmail(body));
          let type;
          if (
            [
              "GenericLocation1",
              "GenericLocation2",
              "GenericLocation3",
            ].includes(loc1) &&
            [
              "GenericLocation1",
              "GenericLocation2",
              "GenericLocation3",
            ].includes(loc2)
          ) {
            type = "Work";
          } else {
            addRow += 1;
            type = "Leisure";
          }
          const isMorning = formattedEmailTime < "12:00:00";
          const add = type === "Leisure" ? 1 : isMorning ? 0 : 1;

          Logger.log(isMorning ? "Morning" : "Afternoon");
          Logger.log(add + addRow);

          // Find and enter the data
          findAndEnterData(
            sheet,
            formattedEmailDate,
            money,
            add + addRow,
            type,
            loc1,
            loc2,
            car,
            "PaymentMethod",
            "Person"
          );
        }
      });

      thread.markRead();
    });
  }

  Logger.log("Processed " + threads.length + " threads.");
}

const extractFromEmail = (body, regex, processMatch = (match) => match[1]) =>
  body.match(regex) ? processMatch(body.match(regex)) : -1;

const extractCarFromEmail = (body) =>
  extractFromEmail(body, /Receipt[\s\S]*?\n([^\n]+)/);

const extractTotalPaidAmountFromEmail = (body) =>
  extractFromEmail(
    body,
    /Total Paid[\s\S]*?Amount (\d+)\.(\d+)/,
    (match) => match[1] + match[2]
  );

const extractLocationFromEmail = (body) =>
  extractFromEmail(
    body,
    /([^\d\[\]:][^\n]+)(?=\s*\d{1,2}:\d{2}[APM]+|\[image: drop-off]|$)/g,
    (match) => match.map((m) => m.trim())
  );

const map = new Map([
  [12, 11],
  [1, 20],
  [2, 29],
  [3, 38],
  [4, 47],
]);

function findAndEnterData(
  sheet,
  targetDate,
  money,
  add,
  type,
  place1,
  place2,
  car,
  card,
  people
) {
  const today = new Date(targetDate);
  today.setHours(0, 0, 0, 0);
  const col = 2; // map.get(today.getMonth() + 1); Get the column based on the month
  const addParsed = Math.round(parseFloat(add)); // Parse and round add once

  Logger.log("Finding...");

  const data = sheet.getRange(1, col, sheet.getLastRow()).getValues();

  // Loop through the data and compare dates
  for (let i = 0; i < data.length; i++) {
    const cellValue = data[i][0];
    if (cellValue instanceof Date) {
      const cellDate = new Date(cellValue).setHours(0, 0, 0, 0);
      if (cellDate === today.getTime()) {
        const row = i + 1 + addParsed;
        Logger.log(`Found today's date in row ${row}`);
        if (type === "Work") {
          enterDataInColumns(sheet, row, col, [
            money,
            type,
            place1,
            place2,
            car,
            card,
            people,
          ]);
          Logger.log("Inserted data");
          return;
        } else {
          copyAndInsertRow(sheet, row, col);
          enterDataInColumns(sheet, row, col, [
            money,
            type,
            place1,
            place2,
            car,
            card,
            people,
          ]);
          Logger.log("Inserted data");
          return;
        }
      }
    }
  }
}

function enterDataInColumns(sheet, row, dateColumn, dataToEnter) {
  sheet
    .getRange(row, dateColumn + 1, 1, dataToEnter.length)
    .setValues([dataToEnter]);
}

function copyAndInsertRow(sheet, targetRow, startColumn) {
  const endColumn = startColumn + 7;

  if (targetRow <= 1) {
    throw new Error("No row above the target row to copy.");
  }

  const rowToCopy = sheet
    .getRange(targetRow - 1, startColumn, 1, endColumn - startColumn + 1)
    .getValues()[0];
  sheet.insertRowBefore(targetRow);
  sheet
    .getRange(targetRow, startColumn, 1, endColumn - startColumn + 1)
    .setValues([rowToCopy]);

  const formatRange = sheet.getRange(
    targetRow - 1,
    startColumn,
    1,
    endColumn - startColumn + 1
  );
  const newRowRange = sheet.getRange(
    targetRow,
    startColumn,
    1,
    endColumn - startColumn + 1
  );

  formatRange.copyFormatToRange(
    sheet,
    startColumn,
    endColumn,
    targetRow,
    targetRow
  );
}

function simplifyText(text) {
  const simplifications = {
    "Location A": "GenericLocation1",
    "Location B": "GenericLocation2",
    "Location C": "GenericLocation3",
    "CarService X": "CarService",
    "CarService Y": "CarService",
  };
  return simplifications[text] || "New Place";
}
