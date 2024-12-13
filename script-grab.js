function getEmailsFromSpecificSenderToday() {
  const senderEmail = "your-email@example.com";
  const today = new Date();
  const timezone = "GMT+7";
  const formattedDate = Utilities.formatDate(
    new Date(today.getFullYear(), today.getMonth(), today.getDate()),
    timezone,
    "yyyy/MM/dd"
  );
  Logger.log("Today's Date: " + formattedDate);

  const query = `from:${senderEmail} after:` + formattedDate;
  const threads = GmailApp.search(query);
  Logger.log("Threads found: " + threads.length);

  const sheet = SpreadsheetApp.openById("YOUR_SPREADSHEET_ID").getSheetByName(
    "YOUR_SHEET_NAME"
  );

  if (threads.length === 0) {
    const date = Utilities.formatDate(
      new Date(formattedDate),
      timezone,
      "MM/dd/yyyy"
    );
    findAndEnterData(
      date,
      0,
      1,
      "Work",
      "Location1",
      "Location2",
      "N/A",
      "N/A",
      "N/A"
    );
    findAndEnterData(
      date,
      0,
      0,
      "Work",
      "Location2",
      "Location1",
      "N/A",
      "N/A",
      "N/A"
    );
  } else {
    threads.forEach((thread) => {
      const messages = thread.getMessages();

      messages.forEach((message) => {
        const sender = message.getFrom();
        const subject = message.getSubject();
        const date = message.getDate();
        const body = message.getPlainBody();
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

        if (body.startsWith(" Grab E-Receipt")) {
          const money = extractTotalPaidAmountFromEmail(body);
          const location = extractLocationFromEmail(body);
          const loc1 = simplifyText(location[0]);
          const loc2 = simplifyText(location[1]);
          const car = simplifyText(extractGrabFromEmail(body));
          const type =
            ["Location1", "Location2", "Location3"].includes(loc1) &&
            ["Location1", "Location2", "Location3"].includes(loc2)
              ? "Work"
              : "Leisure";

          const isMorning = formattedEmailTime < "12:00:00";
          const add = isMorning ? 0 : 1;

          Logger.log(isMorning ? "Morning" : "Afternoon");

          // Find and enter the data
          findAndEnterData(
            formattedEmailDate,
            money,
            add,
            type,
            loc1,
            loc2,
            car,
            "CardName",
            "PersonName"
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

const extractGrabFromEmail = (body) =>
  extractFromEmail(body, /Grab E-Receipt[\s\S]*?\n([^\n]+)/);

const extractTotalPaidAmountFromEmail = (body) =>
  extractFromEmail(
    body,
    /Total Paid[\s\S]*?RP (\d+)\.(\d+)/,
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
  const col = map.get(today.getMonth() + 1);
  const addParsed = Math.round(parseFloat(add));

  Logger.log("Finding...");

  const sheet = SpreadsheetApp.openById("YOUR_SPREADSHEET_ID").getSheetByName(
    "YOUR_SHEET_NAME"
  );
  const data = sheet.getRange(1, col, sheet.getLastRow()).getValues();

  // Loop through the data and compare dates
  for (let i = 0; i < data.length; i++) {
    const cellValue = data[i][0];
    if (cellValue instanceof Date) {
      const cellDate = new Date(cellValue).setHours(0, 0, 0, 0);
      if (cellDate === today.getTime()) {
        const row = i + 1 + addParsed;
        Logger.log(`Found today's date in row ${row}`);
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

function enterDataInColumns(sheet, row, dateColumn, dataToEnter) {
  sheet
    .getRange(row, dateColumn + 1, 1, dataToEnter.length)
    .setValues([dataToEnter]);
}

function simplifyText(text) {
  const simplifications = {
    "Location1 Full Name": "Location1",
    "Location2 Full Name": "Location2",
    "Location3 Full Name": "Location3",
    ServiceA: "ServiceA",
    ServiceB: "ServiceB",
  };
  return simplifications[text] || "N/A";
}
