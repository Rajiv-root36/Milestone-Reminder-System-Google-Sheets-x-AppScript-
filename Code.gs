// ============================================================
// PROJECT REMINDER SYSTEM — Apps Script
// ============================================================
// HOW TO SET UP (one time only):
// 1. Open your Google Sheet → Extensions → Apps Script
// 2. Paste this entire file into Code.gs and save (Ctrl+S)
// 3. Reload the sheet — a "📋 Project Reminders" menu will appear
// 4. Click "📋 Project Reminders" → "Activate daily email trigger"
//    → Approve permissions when Google asks → Done!
// ============================================================

const DASHBOARD_SHEET_NAME = "Dashboard";

// ============================================================
// TASK DEFINITIONS PER COMPLEXITY
// type "offset"      = fixed days from start (negative = before)
// type "offsetEnd"   = fixed days from end date
// type "nextWeekday" = next Nth occurrence of a weekday after start+minDays
// type "lastWeekday" = last occurrence of a weekday in the end date's week
// type "everyWeekday"= every occurrence of a weekday from start until end
// ============================================================

const TASKS_FT = [
  { name: "Follow up on 1st draft of test cases",
    type: "offset", offsetDays: -45 },

  { name: "Follow up on final draft of test cases",
    type: "offset", offsetDays: -30 },

  { name: "Create the Google execution sheet",
    type: "offset", offsetDays: -10 },

  { name: "Send start of Step 2 mail",
    type: "offset", offsetDays: 0 },

  { name: "Message TAM to schedule client meeting",
    type: "offset", offsetDays: 2 },

  { name: "Update lead planning on Buyer Central",
    type: "offset", offsetDays: 4 },                         // First Friday (Mon+4)

  { name: "Send QA progress report",
    type: "nextWeekday", weekday: 1, minDays: 1, nth: 1 },  // Next Monday after start

  { name: "Send Go/No-Go + request UAT environment",
    type: "nextWeekday", weekday: 3, minDays: 1, nth: 1 },  // Next Wednesday after start

  { name: "Update lead planning + send end of Step 2 mail",
    type: "mondayAfterEnd" },                                // Monday after end date
];

const TASKS_C = [
  { name: "Follow up on 1st draft of test cases",
    type: "offset", offsetDays: -60 },                      // 2 months before start

  { name: "Follow up on final draft of test cases",
    type: "offset", offsetDays: -45 },                      // 1.5 months before start

  { name: "Create the Google execution sheet",
    type: "offset", offsetDays: -20 },                      // 20 days before start

  { name: "Send start of Step 2 mail",
    type: "offset", offsetDays: 0 },                        // Start date

  { name: "Message TAM to schedule client meeting",
    type: "offset", offsetDays: 2 },                        // +2 days after start

  { name: "Update lead planning on Buyer Central",
    type: "everyWeekday", weekday: 5 },                     // Every Friday start→end

  { name: "Send QA progress report",
    type: "everyWeekday", weekday: 1 },                     // Every Monday start→end

  { name: "Send Go/No-Go + request UAT environment",
    type: "lastWeekday", weekday: 3 },                      // Last Wednesday of end week

  { name: "Update lead planning + send end of Step 2 mail",
    type: "mondayAfterEnd" },                               // Monday after end date
];

// ============================================================
// 1. SETUP — Run once to create the Dashboard sheet
// ============================================================
function setupDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);

  if (sheet) {
    const otherSheets = ss.getSheets().filter(s => s.getName() !== DASHBOARD_SHEET_NAME);
    if (otherSheets.length === 0) {
      SpreadsheetApp.getUi().alert("⚠️ Cannot reset — Dashboard is the only sheet. Add another sheet first, then try again.");
      return;
    }
    const ui = SpreadsheetApp.getUi();
    const res = ui.alert("Dashboard sheet already exists. Reset it?", ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) return;
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet(DASHBOARD_SHEET_NAME);

  // Headers — all columns including task names and Notes
  // Task headers starting at column F
  const allTaskNames = [
    "Follow up on 1st draft of test cases",
    "Follow up on final draft of test cases",
    "Create the Google execution sheet",
    "Send start of Step 2 mail",
    "Message TAM to schedule client meeting",
    "Update lead planning on Buyer Central",
    "Send QA progress report",
    "Send Go/No-Go + request UAT environment",
    "Update lead planning + send end of Step 2 mail"
  ];
  const allHeaders = ["Project Name", "Start Date", "End Date", "Complexity", "Status"].concat(allTaskNames).concat(["Notes"]);
  sheet.getRange(1, 1, 1, allHeaders.length).setValues([allHeaders]);
  sheet.getRange(1, 1, 1, allHeaders.length)
    .setBackground("#1a1a2e")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(11);

  // Column widths
  sheet.setColumnWidth(1, 220); // Project Name
  sheet.setColumnWidth(2, 110); // Start Date
  sheet.setColumnWidth(3, 110); // End Date
  sheet.setColumnWidth(4, 100); // Complexity
  sheet.setColumnWidth(5, 110); // Status
  for (let i = 6; i <= 5 + allTaskNames.length; i++) sheet.setColumnWidth(i, 160);
  sheet.setColumnWidth(6 + allTaskNames.length, 200); // Notes

  // Date picker for Start/End
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setHelpText("Pick a date using the calendar (DD/MM/YYYY)")
    .build();
  sheet.getRange(2, 2, 100, 2).setDataValidation(dateRule);
  sheet.getRange(2, 2, 100, 2).setNumberFormat("dd/mm/yyyy");

  // Complexity dropdown
  const complexityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["FT", "C1", "C2", "C3"], true)
    .build();
  sheet.getRange(2, 4, 100, 1).setDataValidation(complexityRule);

  // Status dropdown
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Upcoming", "Active", "Completed", "On Hold"], true)
    .build();
  sheet.getRange(2, 5, 100, 1).setDataValidation(statusRule);

  // Tooltip on task date columns to warn about manual overrides
  const taskNoteRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setHelpText("Auto-calculated based on your timeline. You can override this date manually — but editing Project Name, Start Date, End Date or Complexity will reset it back to the calculated date.")
    .build();
  sheet.getRange(2, 6, 100, allTaskNames.length).setDataValidation(taskNoteRule);
  sheet.getRange(2, 6, 100, allTaskNames.length).setNumberFormat("dd/mm/yyyy");

  sheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert(
    "✅ Dashboard created!\n\n" +
    "How to add a project:\n" +
    "1. Enter Project Name, Start Date, End Date, Complexity\n" +
    "2. Task dates calculate automatically\n\n" +
    "Next step:\n" +
    "Click '📋 Project Reminders' → 'Activate daily email trigger'"
  );
}

// ============================================================
// 2. ACTIVATE TRIGGER
// ============================================================
function createDailyTrigger() {
  const userEmail = Session.getActiveUser().getEmail();

  if (!userEmail) {
    SpreadsheetApp.getUi().alert("⚠️ Could not detect your email. Please ensure you are logged into your Google account.");
    return;
  }

  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "checkRemindersAndSendEmail") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("checkRemindersAndSendEmail")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 620px; margin: 0 auto; color: #1a1a2e;">
      <div style="background: #1a1a2e; padding: 20px 24px; border-radius: 8px 8px 0 0;">
        <h2 style="color: #ffffff; margin: 0; font-size: 18px;">✅ You're all set!</h2>
        <p style="color: #9ca3af; margin: 4px 0 0; font-size: 13px;">Project Reminder System activated</p>
      </div>
      <div style="background: #f8f9fa; padding: 20px 24px; border-radius: 0 0 8px 8px; border: 1px solid #e5e7eb; border-top: none;">
        <p style="font-size: 14px; color: #374151;">Hi <strong>${userEmail}</strong>,</p>
        <p style="font-size: 14px; color: #374151; margin-top: 8px;">
          Your daily project reminders are now active. Every morning at <strong>8 AM</strong>,
          you will automatically receive an email listing all tasks due that day.
        </p>
        <div style="background: #ffffff; border: 1px solid #e5e7eb; border-radius: 8px; padding: 16px 20px; margin: 16px 0;">
          <p style="font-size: 13px; font-weight: bold; color: #1a1a2e; margin: 0 0 10px;">How to add a project:</p>
          <p style="font-size: 13px; color: #374151; margin: 6px 0;">1. Enter <strong>Project Name</strong>, <strong>Start Date</strong>, <strong>End Date</strong></p>
          <p style="font-size: 13px; color: #374151; margin: 6px 0;">2. Select <strong>Complexity</strong> — FT, C1, C2, or C3</p>
          <p style="font-size: 13px; color: #374151; margin: 6px 0;">3. Task dates calculate <strong>automatically</strong> based on complexity</p>
          <p style="font-size: 13px; color: #374151; margin: 6px 0;">4. Reminders arrive every morning at 8 AM ☕</p>
        </div>
        <p style="font-size: 12px; color: #9ca3af; margin-top: 16px; text-align: center;">
          Project Reminder System &middot; <a href="${SpreadsheetApp.getActiveSpreadsheet().getUrl()}" style="color: #4B6CB7;">Open Dashboard</a>
        </p>
      </div>
    </div>
  `;

  MailApp.sendEmail({
    to: userEmail,
    subject: "✅ Project Reminder System — You're activated!",
    htmlBody: htmlBody
  });

  SpreadsheetApp.getUi().alert(
    "✅ All done!\n\nA confirmation email has been sent to:\n" + userEmail +
    "\n\nYou will now receive daily reminders every morning at 8 AM automatically."
  );
}

// ============================================================
// 3. CORE — Runs daily at 8 AM
//    Step 1: Auto-update statuses
//    Step 2: Check tasks due today and send email
// ============================================================
function checkRemindersAndSendEmail() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const sheet     = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) return;

  const userEmail = Session.getActiveUser().getEmail();
  const today     = stripTime(new Date());
  const lastRow   = sheet.getLastRow();
  if (lastRow < 2) return;

  const allData = sheet.getDataRange().getValues();

  // --- Step 1: Auto-update statuses ---
  for (let row = 2; row <= lastRow; row++) {
    const r           = allData[row - 1];
    const projectName = r[0];
    const startDate   = r[1];
    const endDate     = r[2];
    const currStatus  = r[4];

    if (!projectName || !startDate || !endDate) continue;
    if (currStatus === "On Hold") continue;

    const start = stripTime(new Date(startDate));
    const end   = stripTime(new Date(endDate));
    let newStatus = "Upcoming";
    if (today >= start && today <= end) newStatus = "Active";
    if (today > end) newStatus = "Completed";

    if (newStatus !== currStatus) {
      sheet.getRange(row, 5).setValue(newStatus);
      colorRow(sheet, row, newStatus);
      allData[row - 1][4] = newStatus; // update in-memory too
    }
  }

  // --- Step 2: Check tasks due today ---
  const dueTasks = [];

  for (let row = 1; row < allData.length; row++) {
    const projectName = allData[row][0];
    const startDate   = allData[row][1];
    const endDate     = allData[row][2];
    const complexity  = allData[row][3];
    const status      = allData[row][4];

    if (!projectName || !startDate || !endDate || !complexity) continue;
    if (status === "Completed" || status === "On Hold") continue;

    const tasks = getTasks(complexity);
    const start = new Date(startDate);
    const end   = new Date(endDate);

    tasks.forEach(task => {
      const taskDates = resolveTaskDates(task, start, end);
      taskDates.forEach(taskDate => {
        if (stripTime(taskDate).getTime() === today.getTime()) {
          dueTasks.push({
            project: projectName,
            task: task.name,
            complexity: complexity,
            startDate: formatDate(startDate),
            endDate: formatDate(endDate)
          });
        }
      });
    });
  }

  Logger.log("Tasks due today: " + dueTasks.length);
  if (dueTasks.length === 0) return;

  sendReminderEmail(dueTasks, today, userEmail);
}

// ============================================================
// 4. ON EDIT — recalculates task dates when project row changes
// ============================================================
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== DASHBOARD_SHEET_NAME) return;

  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (row < 2) return;
  // React to changes in: Name(1), Start(2), End(3), Complexity(4)
  if (col > 4) return;

  const rowData     = sheet.getRange(row, 1, 1, 4).getValues()[0];
  const projectName = rowData[0];
  const startDate   = rowData[1];
  const endDate     = rowData[2];
  const complexity  = rowData[3];

  if (!projectName || !startDate || !endDate || !complexity) return;

  writeTaskDatesToRow(sheet, row, new Date(startDate), new Date(endDate), complexity);

  // Auto-set status
  const today = stripTime(new Date());
  const start = stripTime(new Date(startDate));
  const end   = stripTime(new Date(endDate));
  let status  = "Upcoming";
  if (today >= start && today <= end) status = "Active";
  if (today > end) status = "Completed";
  sheet.getRange(row, 5).setValue(status);
  colorRow(sheet, row, status);
}

// ============================================================
// 5. EMAIL BUILDER
// ============================================================
function sendReminderEmail(dueTasks, today, userEmail) {
  const recipient = userEmail || Session.getActiveUser().getEmail();
  const subject   = `📋 ${dueTasks.length} task${dueTasks.length > 1 ? "s" : ""} due today — ${formatDate(today)}`;

  const byProject = {};
  dueTasks.forEach(item => {
    if (!byProject[item.project]) byProject[item.project] = [];
    byProject[item.project].push(item);
  });

  let htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 620px; margin: 0 auto; color: #1a1a2e;">
      <div style="background: #1a1a2e; padding: 20px 24px; border-radius: 8px 8px 0 0;">
        <h2 style="color: #ffffff; margin: 0; font-size: 18px;">📋 Daily Project Reminders</h2>
        <p style="color: #9ca3af; margin: 4px 0 0; font-size: 13px;">${formatDate(today)}</p>
      </div>
      <div style="background: #f8f9fa; padding: 20px 24px; border-radius: 0 0 8px 8px; border: 1px solid #e5e7eb; border-top: none;">
  `;

  Object.keys(byProject).forEach(projectName => {
    const items = byProject[projectName];
    const badge = items[0].complexity;
    const badgeColor = badge === "FT" ? "#4B6CB7" : "#7C3AED";
    htmlBody += `
      <div style="background: #ffffff; border: 1px solid #e5e7eb; border-radius: 8px; padding: 16px 20px; margin-bottom: 16px;">
        <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px;">
          <div style="display: flex; align-items: center; gap: 8px;">
            <h3 style="margin: 0; font-size: 15px; color: #1a1a2e;">${projectName}</h3>
            <span style="font-size: 11px; font-weight: 600; color: #fff; background: ${badgeColor}; padding: 2px 8px; border-radius: 20px;">${badge}</span>
          </div>
          <span style="font-size: 11px; color: #6b7280; background: #f3f4f6; padding: 3px 10px; border-radius: 20px;">
            ${items[0].startDate} → ${items[0].endDate}
          </span>
        </div>
        <table style="width: 100%; border-collapse: collapse;">
    `;
    items.forEach(item => {
      htmlBody += `
          <tr>
            <td style="padding: 8px 0; border-bottom: 1px solid #f3f4f6;">
              <span style="display: inline-block; width: 8px; height: 8px; background: #E24B4A; border-radius: 50%; margin-right: 10px; vertical-align: middle;"></span>
              <span style="font-size: 14px; color: #374151;">${item.task}</span>
            </td>
            <td style="text-align: right; padding: 8px 0; border-bottom: 1px solid #f3f4f6;">
              <strong style="font-size: 13px; color: #E24B4A;">Due today</strong>
            </td>
          </tr>
      `;
    });
    htmlBody += `</table></div>`;
  });

  htmlBody += `
        <p style="font-size: 12px; color: #9ca3af; margin-top: 8px; text-align: center;">
          Sent by your Project Reminder System &middot;
          <a href="${SpreadsheetApp.getActiveSpreadsheet().getUrl()}" style="color: #4B6CB7;">Open Dashboard</a>
        </p>
      </div>
    </div>
  `;

  MailApp.sendEmail({ to: recipient, subject: subject, htmlBody: htmlBody });
}

// ============================================================
// 6. DEBUG
// ============================================================
function debugDates() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  const data  = sheet.getDataRange().getValues();
  const today = stripTime(new Date());

  let msg = "Today: " + today.toDateString() + "\n\n";

  for (let row = 1; row < data.length; row++) {
    const projectName = data[row][0];
    const startDate   = data[row][1];
    const endDate     = data[row][2];
    const complexity  = data[row][3];
    if (!projectName || !startDate || !endDate || !complexity) continue;

    msg += "Project: " + projectName + " [" + complexity + "]\n";
    const tasks = getTasks(complexity);
    tasks.forEach(task => {
      const dates = resolveTaskDates(task, new Date(startDate), new Date(endDate));
      dates.forEach(d => {
        const isToday = stripTime(d).getTime() === today.getTime();
        msg += "  [" + (isToday ? "TODAY ✅" : "      ") + "] " + task.name + ": " + d.toDateString() + "\n";
      });
    });
    msg += "\n";
  }

  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}

// ============================================================
// 7. MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📋 Project Reminders")
    .addItem("Activate daily email trigger", "createDailyTrigger")
    .addToUi();
}

// ============================================================
// HELPERS
// ============================================================

// Returns correct task list based on complexity
function getTasks(complexity) {
  return (complexity === "FT") ? TASKS_FT : TASKS_C;
}

// Resolves a task definition into one or more actual Date objects
function resolveTaskDates(task, startDate, endDate) {
  const start = new Date(startDate);
  const end   = new Date(endDate);

  switch (task.type) {
    case "offset": {
      const d = new Date(start);
      d.setDate(d.getDate() + task.offsetDays);
      return [d];
    }
    case "nextWeekday": {
      return [getNextWeekday(start, task.weekday, task.minDays, task.nth)];
    }
    case "lastWeekday": {
      return [getLastWeekdayOfWeek(end, task.weekday)];
    }
    case "mondayAfterEnd": {
      return [getMondayAfter(end)];
    }
    case "everyWeekday": {
      return getAllWeekdays(start, end, task.weekday);
    }
    default:
      return [];
  }
}

// Writes computed task dates to the row (col 6+)
// Skips cells that have been manually overridden by the user
function writeTaskDatesToRow(sheet, row, startDate, endDate, complexity) {
  const tasks    = getTasks(complexity);
  const startCol = 6;

  // Read existing values to detect manual overrides
  const existingVals = sheet.getRange(row, startCol, 1, tasks.length).getValues()[0];

  tasks.forEach((task, i) => {
    const dates     = resolveTaskDates(task, startDate, endDate);
    const formatted = dates.map(d => formatDate(d)).join(", ");
    const existing  = existingVals[i];

    // If cell is empty, always write the calculated date
    if (!existing || existing === "") {
      sheet.getRange(row, startCol + i).setValue(formatted);
      return;
    }

    // If cell has a value, check if it matches the calculated date
    // If it doesn't match → user has manually overridden it → leave it alone
    const existingStr = existing instanceof Date
      ? formatDate(existing)
      : existing.toString().trim();

    if (existingStr === formatted) {
      // Matches calculated — safe to overwrite (no manual override)
      sheet.getRange(row, startCol + i).setValue(formatted);
    }
    // else: manual override detected — skip this cell
  });
}

// Next Nth occurrence of a weekday (0=Sun...6=Sat) after start+minDays
function getNextWeekday(startDate, weekday, minDays, nth) {
  const base = new Date(startDate);
  base.setDate(base.getDate() + (minDays || 1));
  let count = 0;
  while (true) {
    if (base.getDay() === weekday) {
      count++;
      if (count === (nth || 1)) return new Date(base);
    }
    base.setDate(base.getDate() + 1);
  }
}

// Last occurrence of a weekday in the same calendar week as endDate (Mon–Sun)
function getLastWeekdayOfWeek(endDate, weekday) {
  const end       = new Date(endDate);
  const dayOfWeek = end.getDay(); // 0=Sun
  // Find Monday of that week
  const monday = new Date(end);
  monday.setDate(end.getDate() - ((dayOfWeek + 6) % 7));
  // Find the target weekday in that week
  const target = new Date(monday);
  target.setDate(monday.getDate() + ((weekday + 6) % 7)); // adjust: Mon=0 in our offset
  // If weekday is Wed(3), offset from Monday is 2
  const offset = (weekday === 0) ? 6 : weekday - 1;
  target.setDate(monday.getDate() + offset);
  return target;
}

// Monday immediately after endDate
function getMondayAfter(endDate) {
  const d = new Date(endDate);
  d.setDate(d.getDate() + 1);
  while (d.getDay() !== 1) {
    d.setDate(d.getDate() + 1);
  }
  return d;
}

// All occurrences of a weekday between startDate and endDate (inclusive)
function getAllWeekdays(startDate, endDate, weekday) {
  const dates = [];
  const d     = new Date(startDate);
  d.setDate(d.getDate() + 1); // start from day after project start
  while (d <= endDate) {
    if (d.getDay() === weekday) dates.push(new Date(d));
    d.setDate(d.getDate() + 1);
  }
  return dates;
}

function stripTime(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function formatDate(date) {
  return new Date(date).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}

function colorRow(sheet, row, status) {
  const totalCols = 25;
  const colors    = { "Active": "#e6f4ea", "Upcoming": "#e8f0fe", "Completed": "#f1f3f4", "On Hold": "#fff3e0" };
  sheet.getRange(row, 1, 1, totalCols).setBackground(colors[status] || "#ffffff");
}

function applyAlternatingColors() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) return;

  const lastCol = 15; // covers all columns
  
  for (let row = 2; row <= 100; row++) {
    const color = row % 2 === 0 ? "#f8f9fa" : "#ffffff";
    sheet.getRange(row, 1, 1, lastCol).setBackground(color);
  }

  SpreadsheetApp.getUi().alert("✅ Alternating row colors applied!");
}
