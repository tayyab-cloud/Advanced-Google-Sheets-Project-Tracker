// ===============================================================
// GLOBAL SETTINGS
// ===============================================================

const SHEET_NAME = "Tasks"; // The name of the sheet where tasks are stored.

// ===============================================================
// CORE UI & TRIGGERS
// ===============================================================

/**
 * Creates the main menu when the spreadsheet is opened.
 * Also runs formatRows() on open to ensure colors are correct.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('✅ Project Tracker')
    .addItem('➕ Add New Task', 'showAddTaskSidebar')
    .addItem('👥 Add New Assignee', 'showAddAssigneeDialog')
    .addItem('👤 View/Edit Team', 'showViewTeamSidebar') 
    .addItem('🔍 Search & Filter Tasks', 'showSearchSidebar')
    .addSeparator()
    .addItem('📊 Generate Dashboard', 'generateFullDashboard')
    .addItem('⚠️ View Overdue Tasks', 'goToOverdueTasks')
    .addSeparator()
    .addItem('❌ Delete Selected Tasks', 'deleteSelectedTasksWithConfirmation')
    .addItem('🗄️ Archive Completed Tasks', 'archiveTasks')
    .addSeparator() // A separator to keep dangerous actions separate
    .addItem('🔥 Purge Old Archived Tasks', 'purgeArchivedTasks') // <-- THIS IS THE NEW LINE
    .addToUi();

  formatRows(); 
}

/**
 * Runs when a user edits a cell. Used to trigger the edit task dialog.
 */
function handleEditTrigger(e) {
  const range = e.range;
  // The "Edit" checkbox is in Column J (column number 10).
  if (range.getColumn() === 10 && range.getValue() === true) {
    range.setValue(false); // Immediately uncheck the box.
    openEditDialog();
  }
}

// ===============================================================
// SIDEBAR & DIALOG FUNCTIONS
// ===============================================================

/**
 * Shows the sidebar to add a new task.
 */
function showAddTaskSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('AddTaskSidebar').setTitle('Add New Task');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Shows a dialog box to add a new team member.
 */
function showAddAssigneeDialog() {
  const html = HtmlService.createHtmlOutputFromFile('AddAssigneeDialog').setWidth(400).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Assignee');
}

/**
 * Opens the Edit Task dialog with data from the currently selected row.
 */
function openEditDialog() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const activeRow = sheet.getActiveRange().getRow();

  if (activeRow === 1) return; // Ignore the header row.

  // Get data from columns A to G (7 columns).
  const rowData = sheet.getRange(activeRow, 1, 1, 7).getValues()[0];
  
  // Map data using the correct new column indices.
  const taskData = {
    row: activeRow,
    id: rowData[0],       // Column A
    title: rowData[1],    // Column B
    assignee: rowData[2], // Column C (Name)
    priority: rowData[4], // Column E
    status: rowData[5],   // Column F
    dueDate: rowData[6]   // Column G
  };

  const htmlTemplate = HtmlService.createTemplateFromFile('EditForm');
  htmlTemplate.task = taskData;
  const htmlOutput = htmlTemplate.evaluate().setWidth(400).setHeight(550);
  ui.showModalDialog(htmlOutput, 'Edit Task Details');
}

// ===============================================================
// DATA MANIPULATION (ADD, EDIT, DELETE)
// ===============================================================

/**
 * Adds a new task to the 'Tasks' sheet using data from the sidebar.
 */
// DEBUGGING VERSION
function addTask(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const taskID = "TSK-" + ("000" + (lastRow)).slice(-3);
  const timestamp = new Date();

  const [assigneeName, assigneeEmail] = data.assignee.split('|||');

  sheet.appendRow([
    taskID, data.title, assigneeName, assigneeEmail, data.priority, data.status, 
    new Date(data.dueDate), timestamp, '', '' 
  ]);

  const newRow = sheet.getLastRow();
  sheet.getRange(newRow, 9, 1, 2).insertCheckboxes(); 
  sheet.getRange(newRow, 7).setNumberFormat("yyyy-mm-dd");

  // --- TEMPORARILY COMMENTED OUT FOR DEBUGGING ---
  formatRows();
  sortTasksByDate();
  generateFullDashboard();
  
  return "✅ Task '" + data.title + "' has been added successfully!";
}

/**
 * Updates an existing task on the sheet using data from the edit form.
 */
function updateTaskOnSheet(formData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const row = formData.row;

  // Set values in the correct new columns.
  // NOTE: This version does not update the assignee's email, only the name.
  // A more advanced edit form would be needed to change assignees.
  sheet.getRange(row, 2).setValue(formData.title);      // Col B
  sheet.getRange(row, 3).setValue(formData.assignee);   // Col C
  sheet.getRange(row, 5).setValue(formData.priority);   // Col E
  sheet.getRange(row, 6).setValue(formData.status);     // Col F
  sheet.getRange(row, 7).setValue(new Date(formData.dueDate)); // Col G
  sheet.getRange(row, 8).setValue(new Date());          // Col H (Timestamp)

  // Update formatting and dashboard.
  formatRows();
  generateFullDashboard(); 

  return "Task updated successfully!";
}

/**
 * Adds a new assignee to the 'Team' sheet. Allows duplicate names but enforces unique emails.
 */
function addNewAssignee(formData) {
  if (!formData.name || !formData.email || !formData.dob || !formData.address) {
    throw new Error("All fields are required.");
  }

  const name = formData.name.trim();
  const email = formData.email.trim();
  const teamSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Team");

  if (!teamSheet) { throw new Error("'Team' sheet not found. Please check the sheet name."); }

  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRegex.test(email)) { throw new Error("Please enter a valid email address."); }

  const existingData = teamSheet.getDataRange().getValues();
  for (let i = 1; i < existingData.length; i++) {
    // Only check for duplicate emails.
    if (existingData[i][2] && existingData[i][2].toString().toLowerCase() === email.toLowerCase()) {
      throw new Error(`Email '${email}' is already registered.`);
    }
  }
  
  const lastRow = teamSheet.getLastRow();
  const newId = "EMP-" + ("000" + (lastRow)).slice(-3);
  const dob = new Date(formData.dob);
  const age = calculateAge(dob);
  
  // Append row. Order must match Team sheet: ID, Name, Email, DOB, Address, Age.
  teamSheet.appendRow([newId, name, email, dob, formData.address, age]);
  teamSheet.getRange(teamSheet.getLastRow(), 4).setNumberFormat("yyyy-mm-dd");

  return `✅ Success! '${name}' has been added to the team.`;
}

/**
 * Asks for confirmation and then deletes all tasks with the 'Delete' checkbox ticked.
 */
function deleteSelectedTasksWithConfirmation() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  // 'Delete' checkbox is in Column I (index 8).
  const deleteColumnIndex = 8;
  let rowsToDelete = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][deleteColumnIndex] === true) {
      rowsToDelete.push(i + 1);
    }
  }

  if (rowsToDelete.length === 0) {
    SpreadsheetApp.getUi().alert("No tasks were selected for deletion.");
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Confirm Deletion',
    `Are you sure you want to delete ${rowsToDelete.length} selected task(s)?`,
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    // Delete rows from the bottom up to avoid messing up row indices.
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      sheet.deleteRow(rowsToDelete[i]);
    }
    SpreadsheetApp.getUi().alert("✅ Selected tasks deleted successfully.");
    generateFullDashboard();
  }
}

// ===============================================================
// DASHBOARD & REPORTING
// ===============================================================

/**
 * Generates or UPDATES a full dashboard with metrics and charts.
 * Handles duplicate names by using unique emails for grouping.
 */
function generateFullDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const taskSheet = ss.getSheetByName(SHEET_NAME);
  let dashboardSheet = ss.getSheetByName("Dashboard");

  if (!dashboardSheet) { dashboardSheet = ss.insertSheet("Dashboard"); }

  const data = taskSheet.getDataRange().getValues();
  let toDo = 0, inProgress = 0, done = 0, overdue = 0;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const assigneeCounts = {}; 

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    const assigneeName = row[2];
    const assigneeEmail = row[3];
    const status = row[5];
    const dueDate = new Date(row[6]);

    if (status === "To Do") toDo++;
    else if (status === "In Progress") inProgress++;
    else if (status === "Done") done++;

    if (status !== "Done" && dueDate < today) overdue++;

    if (assigneeEmail) {
      if (!assigneeCounts[assigneeEmail]) {
        assigneeCounts[assigneeEmail] = { name: assigneeName, count: 0 };
      }
      assigneeCounts[assigneeEmail].count++;
    }
  }

  dashboardSheet.getRange("A:B").clearContent();
  dashboardSheet.getRange("D:E").clearContent();
  dashboardSheet.getRange("G:H").clearContent();

  dashboardSheet.getRange("A1:B1").setValues([["Metric", "Value"]]).setFontWeight("bold");
  const metrics = [
    ["Total Tasks", data.length - 1], ["To Do", toDo], ["In Progress", inProgress], ["Done", done], ["Overdue Tasks", overdue]
  ];
  dashboardSheet.getRange(2, 1, metrics.length, 2).setValues(metrics);
  dashboardSheet.getRange("A1:B" + (metrics.length + 1)).setBorder(true, true, true, true, true, true);
  dashboardSheet.autoResizeColumns(1, 2);

  const statusSummaryData = [["Status", "Count"], ["To Do", toDo], ["In Progress", inProgress], ["Done", done]];
  const statusRange = dashboardSheet.getRange(1, 4, statusSummaryData.length, 2);
  statusRange.setValues(statusSummaryData).setFontWeight("bold");
  dashboardSheet.getRange(2, 4, statusSummaryData.length - 1, 2).setFontWeight("normal");

  const assigneeSummaryData = [["Assignee", "Task Count"]];
  for (const email in assigneeCounts) {
    const assigneeInfo = assigneeCounts[email];
    const uniqueLabel = `${assigneeInfo.name} (${email})`;
    assigneeSummaryData.push([uniqueLabel, assigneeInfo.count]);
  }
  
  let assigneeRange;
  if (assigneeSummaryData.length > 1) {
    assigneeRange = dashboardSheet.getRange(1, 7, assigneeSummaryData.length, 2);
    assigneeRange.setValues(assigneeSummaryData).setFontWeight("bold");
    dashboardSheet.getRange(2, 7, assigneeSummaryData.length - 1, 2).setFontWeight("normal");
  }

  const charts = dashboardSheet.getCharts();
  let pieChartFound = false;
  let columnChartFound = false;

  charts.forEach(chart => {
    const title = chart.getOptions().get('title');
    if (title === 'Tasks by Status') {
      pieChartFound = true;
      dashboardSheet.updateChart(chart.modify().clearRanges().addRange(statusRange).build());
    } else if (title === 'Tasks per Assignee') {
      columnChartFound = true;
      if (assigneeRange) {
        dashboardSheet.updateChart(chart.modify().clearRanges().addRange(assigneeRange).build());
      }
    }
  });

  if (!pieChartFound && (toDo > 0 || inProgress > 0 || done > 0)) {
    const pieChart = dashboardSheet.newChart().setChartType(Charts.ChartType.PIE).addRange(statusRange)
      .setOption('title', 'Tasks by Status').setOption('pieHole', 0.4).setOption('colors', ['#ef5350', '#ffca28', '#66bb6a'])
      .setPosition(2, 4, 0, 0).build();
    dashboardSheet.insertChart(pieChart);
  }

  if (!columnChartFound && assigneeRange) {
    const columnChart = dashboardSheet.newChart().setChartType(Charts.ChartType.COLUMN).addRange(assigneeRange)
      .setOption('title', 'Tasks per Assignee').setOption('legend', { position: 'none' })
      .setPosition(2, 7, 0, 0).build();
    dashboardSheet.insertChart(columnChart);
  }
  
  dashboardSheet.autoResizeColumns(4, 2);
  dashboardSheet.autoResizeColumns(7, 2);
}

/**
 * Finds all overdue tasks, highlights them, and navigates to the first one.
 */
function goToOverdueTasks() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  let overdueRows = [];

  for (let i = 1; i < data.length; i++) {
    const status = data[i][5];        // Status is in Column F
    const dueDate = new Date(data[i][6]); // Due Date is in Column G
    if (status !== "Done" && dueDate < today) {
      overdueRows.push(i + 1);
    }
  }
  
  formatRows(); // Reset colors first.

  if (overdueRows.length > 0) {
    SpreadsheetApp.setActiveSheet(sheet);
    sheet.setActiveRange(sheet.getRange(overdueRows[0], 1));
    overdueRows.forEach(row => {
      // Highlight main data columns (A to H)
      sheet.getRange(row, 1, 1, 8).setBackground("#ffcdd2");
    });
    ui.alert(`${overdueRows.length} overdue task(s) have been highlighted.`);
  } else {
    ui.alert("✅ No overdue tasks found.");
  }
}

/**
 * Finds all overdue tasks and sends a summary email to each assignee.
 */
function sendOverdueTaskEmails() {
  const taskSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = taskSheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const overdueTasksByAssignee = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const assigneeName = row[2];
    const assigneeEmail = row[3];
    const status = row[5];
    const dueDate = new Date(row[6]);

    if (status !== "Done" && dueDate < today) {
      const taskTitle = row[1];
      const formattedDueDate = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

      // Group by the unique email address.
      if (!overdueTasksByAssignee[assigneeEmail]) {
        overdueTasksByAssignee[assigneeEmail] = { name: assigneeName, tasks: [] };
      }
      overdueTasksByAssignee[assigneeEmail].tasks.push({
        title: taskTitle,
        dueDate: formattedDueDate
      });
    }
  }

  // Loop through the grouped tasks and send emails.
  for (const email in overdueTasksByAssignee) {
    const assigneeInfo = overdueTasksByAssignee[email];
    const subject = `⚠️ Overdue Task Reminder (${assigneeInfo.tasks.length} Tasks)`;
    
    let htmlBody = `<h3>Hello ${assigneeInfo.name},</h3>`;
    htmlBody += `<p>This is a reminder that you have ${assigneeInfo.tasks.length} overdue task(s):</p><ul>`;
    assigneeInfo.tasks.forEach(task => {
      htmlBody += `<li><b>${task.title}</b> (Due Date: ${task.dueDate})</li>`;
    });
    htmlBody += "</ul><p>Please take action as soon as possible. Thank you!</p>";
    
    MailApp.sendEmail({ to: email, subject: subject, htmlBody: htmlBody });
    Logger.log(`Sent reminder to ${assigneeInfo.name} at ${email}.`);
  }
}

// ===============================================================
// HELPER FUNCTIONS
// ===============================================================

/**
 * Gets a list of assignee objects {name, email} from the 'Team' sheet for dropdowns.
 */
function getAssigneeList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Team");
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 2).getValues(); // Get columns B (Name) and C (Email).
  const assigneeList = [];

  data.forEach(row => {
    if (row[0] && row[1]) {
      assigneeList.push({ name: row[0], email: row[1] });
    }
  });
  return assigneeList;
}

/**
 * Applies background colors to rows based on the task status.
 */
function formatRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (sheet.getLastRow() < 2) return; // Don't run if there are no tasks.

  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8);
  const data = dataRange.getValues();
  const colors = [];

  for (let i = 0; i < data.length; i++) {
    const status = data[i][5]; // Status is in Column F.
    let color = "#ffffff";
    if (status === "To Do") color = "#ffebee";
    else if (status === "In Progress") color = "#fff9c4";
    else if (status === "Done") color = "#e8f5e9";
    colors.push([color, color, color, color, color, color, color, color]);
  }
  dataRange.setBackgrounds(colors);
}

/**
 * Sorts the 'Tasks' sheet by Due Date (Column G) in ascending order.
 */
function sortTasksByDate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (sheet.getLastRow() < 2) return;
  // Sort by Column G (column 7).
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).sort({ column: 7, ascending: true });
}

/**
 * Calculates age based on a birth date.
 */
function calculateAge(birthDate) {
  const today = new Date();
  let age = today.getFullYear() - birthDate.getFullYear();
  const m = today.getMonth() - birthDate.getMonth();
  if (m < 0 || (m === 0 && today.getDate() < birthDate.getDate())) {
    age--;
  }
  return age;
}

// ===============================================================
// VIEW & EDIT TEAM MEMBER LOGIC
// ===============================================================

/**
 * Shows the sidebar to view all team members.
 * Called from the main menu.
 */
function showViewTeamSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('ViewTeamSidebar').setTitle('Team Members');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Fetches all details for every team member from the 'Team' sheet.
 * Called by the ViewTeamSidebar.
 * @returns {object[]} An array of team member objects.
 */
function getTeamDetails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Team");
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const members = data.map(row => ({
    id: row[0],
    name: row[1],
    email: row[2],
    dob: Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), "yyyy-MM-dd"),
    address: row[4],
    age: row[5]
  }));
  return members;
}

/**
 * Opens the Edit Assignee dialog pre-filled with a specific member's data.
 * @param {string} assigneeId The unique ID of the assignee to edit (e.g., "EMP-001").
 */
function openEditAssigneeDialog(assigneeId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Team");
  const data = sheet.getDataRange().getValues();
  let assigneeData = null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == assigneeId) {
      assigneeData = {
        id: data[i][0],
        name: data[i][1],
        email: data[i][2],
        dob: Utilities.formatDate(new Date(data[i][3]), Session.getScriptTimeZone(), "yyyy-MM-dd"),
        address: data[i][4]
      };
      break;
    }
  }

  if (!assigneeData) {
    SpreadsheetApp.getUi().alert("Could not find assignee with ID: " + assigneeId);
    return;
  }

  const htmlTemplate = HtmlService.createTemplateFromFile('EditAssigneeDialog');
  htmlTemplate.assignee = assigneeData;
  const htmlOutput = htmlTemplate.evaluate().setWidth(400).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Edit Details for ${assigneeData.name}`);
}

/**
 * Updates an assignee's details in the 'Team' sheet AND synchronizes the
 * name change across all relevant tasks and the dashboard.
 * [FINAL PROFESSIONAL VERSION WITH FULL SYNC]
 *
 * @param {object} formData The updated data from the form.
 * @returns {string} A success message.
 */
function updateAssigneeDetails(formData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Team");
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  let originalEmail = "";

  // Find the row number for the given ID
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == formData.id) {
      rowIndex = i + 1;
      originalEmail = data[i][2]; // Store the email, which is our unique key
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error("Could not update. Assignee not found.");
  }

  const newName = formData.name.trim();
  const newAge = calculateAge(new Date(formData.dob));

  // Step 1: Update the 'Team' sheet (the source of truth)
  sheet.getRange(rowIndex, 2, 1, 5).setValues([[
    newName,
    originalEmail, // Keep original email
    new Date(formData.dob),
    formData.address,
    newAge
  ]]);

  // Step 2: Synchronize the name change in the 'Tasks' sheet.
  synchronizeAssigneeNameInTasks(originalEmail, newName);
  
  // --- THIS IS THE NEW AND FINAL STEP ---
  // Step 3: Regenerate the dashboard to reflect the new name on the chart.
  generateFullDashboard();
  // --- END OF NEW STEP ---

  return "✅ Details updated successfully! System is fully synchronized.";
}
/**
 * Synchronizes the assignee's name across all tasks in the 'Tasks' sheet
 * after their name has been updated in the 'Team' sheet.
 * This is a critical function for maintaining data integrity.
 *
 * @param {string} assigneeEmail The unique email of the assignee.
 * @param {string} newName The new name of the assignee.
 */
function synchronizeAssigneeNameInTasks(assigneeEmail, newName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (sheet.getLastRow() < 2) return; // No tasks to update.

  const range = sheet.getRange(2, 3, sheet.getLastRow() - 1, 2); // Get Columns C (Name) and D (Email)
  const values = range.getValues();

  let hasChanges = false;
  // Loop through all tasks
  for (let i = 0; i < values.length; i++) {
    // Check if the email in the task row matches the updated assignee's email
    if (values[i][1] === assigneeEmail) {
      // If it matches and the name is different, update it
      if (values[i][0] !== newName) {
        values[i][0] = newName; // Update the name in our array
        hasChanges = true;
      }
    }
  }

  // For maximum efficiency, only write back to the sheet if there were changes.
  if (hasChanges) {
    range.setValues(values);
  }
}
/**
 * Checks if a given assignee has any tasks that are not marked as 'Done'.
 * [CORRECTED LOGIC]
 * @param {string} assigneeEmail The unique email of the assignee to check.
 * @returns {boolean} True if they have active tasks, false otherwise.
 */
function hasActiveTasks(assigneeEmail) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (sheet.getLastRow() < 2) return false; // No tasks exist at all.

  // --- THIS IS THE FIX ---
  // We need to get data starting from Column D (Email) up to Column F (Status).
  // That's a total of 3 columns (D, E, F).
  const range = sheet.getRange(2, 4, sheet.getLastRow() - 1, 3);
  const data = range.getValues();
  
  for (let i = 0; i < data.length; i++) {
    // Now the indices match the new data range we fetched:
    const emailInSheet = data[i][0];  // This is Column D
    const statusInSheet = data[i][2]; // This is Column F (index 2 in our new range)
    
    // If we find a task for this email that is NOT 'Done', they have active tasks.
    if (emailInSheet === assigneeEmail && statusInSheet !== 'Done') {
      return true; // Found an active task, stop immediately.
    }
  }
  
  return false; // Loop finished, no active tasks were found.
}

/**
 * Deletes an assignee after checking for active tasks, getting user confirmation,
 * and refreshing the dashboard.
 * [FINAL PROFESSIONAL VERSION WITH FULL SYNC]
 *
 * @param {string} assigneeId The unique ID of the assignee to delete.
 * @returns {string} A success message.
 * @throws {Error} If the assignee has active tasks or is not found.
 */
function deleteAssignee(assigneeId) {
  const teamSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Team");
  const data = teamSheet.getDataRange().getValues();
  let rowIndex = -1;
  let assigneeEmail = "";
  let assigneeName = "";

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == assigneeId) {
      rowIndex = i + 1;
      assigneeName = data[i][1];
      assigneeEmail = data[i][2];
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error("Assignee not found. They may have already been deleted.");
  }
  
  // Step 1: Check for active tasks.
  if (hasActiveTasks(assigneeEmail)) {
    throw new Error(`Cannot delete '${assigneeName}'. They have unfinished tasks.`);
  }

  // Step 2: Show a native confirmation dialog before deleting.
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Confirm Deletion',
    `Are you sure you want to permanently delete '${assigneeName}'? This action cannot be undone.`,
    ui.ButtonSet.YES_NO
  );

  // Step 3: Only proceed if the user clicks 'YES'.
  if (response == ui.Button.YES) {
    // Step 3a: Delete the row from the 'Team' sheet.
    teamSheet.deleteRow(rowIndex);
    
    // --- THIS IS THE NEW AND CRITICAL STEP ---
    // Step 3b: Regenerate the dashboard. This will remove the assignee from the chart
    // ONLY IF they have zero tasks (including 'Done' ones). If they have 'Done'
    // tasks, their record will remain on the dashboard, which is correct.
    generateFullDashboard();
    // --- END OF NEW STEP ---

    return `✅ '${assigneeName}' has been successfully deleted from the team list.`;
  } else {
    throw new Error("Deletion cancelled by user.");
  }
}

// ===============================================================
// ARCHIVING FUNCTION
// ===============================================================

/**
 * Finds all tasks marked as "Done", moves them to an "Archive" sheet,
 * and then deletes them from the main "Tasks" sheet.
 */
function archiveTasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName(SHEET_NAME);
  let archiveSheet = ss.getSheetByName("Archive");

  // Step 1: Create the 'Archive' sheet if it doesn't exist.
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet("Archive");
    // Copy the headers from the Tasks sheet to the new Archive sheet.
    tasksSheet.getRange(1, 1, 1, tasksSheet.getLastColumn()).copyTo(archiveSheet.getRange(1, 1));
  }

  const data = tasksSheet.getDataRange().getValues();
  let archivedCount = 0;
  
  // Step 2: Loop through the tasks from the BOTTOM to the TOP.
  // This is critical to ensure rows are not skipped when deleting.
  for (let i = data.length - 1; i >= 1; i--) {
    const rowData = data[i];
    const status = rowData[5]; // Status is in Column F (index 5)

    // Step 3: Check if the task's status is "Done".
    if (status === 'Done') {
      // Step 3a: Append the entire row to the Archive sheet.
      archiveSheet.appendRow(rowData);
      
      // Step 3b: Delete the original row from the Tasks sheet.
      tasksSheet.deleteRow(i + 1); // i + 1 because row numbers are 1-based.
      
      archivedCount++;
    }
  }

  // Step 4: Provide feedback to the user.
  if (archivedCount > 0) {
    // Step 4a: Refresh the dashboard to show updated counts.
    generateFullDashboard();
    SpreadsheetApp.getUi().alert(`✅ Success! ${archivedCount} completed task(s) have been archived.`);
  } else {
    SpreadsheetApp.getUi().alert("No completed tasks found to archive.");
  }
}

// ===============================================================
// SEARCH & FILTER LOGIC
// ===============================================================

/**
 * Shows the search & filter sidebar.
 * Called from the main menu.
 */
function showSearchSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('SearchSidebar').setTitle('Search & Filter Tasks');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Hides and shows rows in the Tasks and Archive sheets based on filter criteria.
 * @param {object} filters An object containing all the filter criteria from the sidebar.
 * @returns {string} A status message for the user.
 */
function filterTasks(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName(SHEET_NAME);
  const archiveSheet = ss.getSheetByName("Archive");

  // First, hide all rows in the relevant sheets to start fresh.
  if (tasksSheet.getLastRow() > 1) {
    tasksSheet.hideRows(2, tasksSheet.getLastRow() - 1);
  }
  if (filters.includeArchived && archiveSheet && archiveSheet.getLastRow() > 1) {
    archiveSheet.hideRows(2, archiveSheet.getLastRow() - 1);
  }

  let matchingRowCount = 0;
  
  // A helper function to process a sheet
  const processSheet = (sheet, isArchived) => {
    if (!sheet || sheet.getLastRow() < 2) return;
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowNum = i + 1;

      // Extract data from the row using correct indices
      const title = row[1];
      const assigneeName = row[2];
      const assigneeEmail = row[3];
      const priority = row[4];
      const status = row[5];
      
      let isMatch = true;

      // Apply each filter. If any filter fails, set isMatch to false.
      if (filters.keyword && !title.toLowerCase().includes(filters.keyword.toLowerCase())) isMatch = false;
      if (filters.status && status !== filters.status) isMatch = false;
      if (filters.priority && priority !== filters.priority) isMatch = false;
      if (filters.assignee && `${assigneeName}|||${assigneeEmail}` !== filters.assignee) isMatch = false;
      
      // If all filters passed, show the row.
      if (isMatch) {
        sheet.showRows(rowNum);
        matchingRowCount++;
      }
    }
  };

  // Process the main Tasks sheet
  processSheet(tasksSheet, false);
  
  // If requested, also process the Archive sheet
  if (filters.includeArchived) {
    processSheet(archiveSheet, true);
  }

  return `Found ${matchingRowCount} matching task(s).`;
}

/**
 * Resets all filters by showing all rows in the Tasks and Archive sheets.
 * @returns {string} A status message for the user.
 */
function resetFilters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName(SHEET_NAME);
  const archiveSheet = ss.getSheetByName("Archive");

  if (tasksSheet && tasksSheet.getLastRow() > 1) {
    tasksSheet.showRows(2, tasksSheet.getLastRow() - 1);
  }
  if (archiveSheet && archiveSheet.getLastRow() > 1) {
    archiveSheet.showRows(2, archiveSheet.getLastRow() - 1);
  }
  
  return "Filters have been reset.";
}
// ===============================================================
// PURGE FUNCTION
// ===============================================================

/**
 * Permanently deletes tasks from the "Archive" sheet that are older
 * than a user-specified number of days. Includes multiple safety checks.
 * [CORRECTED ALERT LOGIC]
 */
function purgeArchivedTasks() {
  const ui = SpreadsheetApp.getUi();

  const promptResponse = ui.prompt(
    'Purge Old Archives',
    'Enter the age in days for tasks to be deleted (e.g., 365 for tasks older than one year).',
    ui.ButtonSet.OK_CANCEL
  );

  if (promptResponse.getSelectedButton() != ui.Button.OK) {
    return;
  }
  
  const daysText = promptResponse.getResponseText();
  const days = parseInt(daysText);

  if (isNaN(days) || days <= 0) {
    ui.alert('Invalid Input', 'Please enter a positive number for the days.', ui.ButtonSet.OK);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archiveSheet = ss.getSheetByName("Archive");

  if (!archiveSheet || archiveSheet.getLastRow() < 2) {
    ui.alert("The 'Archive' sheet is empty or does not exist.");
    return;
  }
  
  const today = new Date();
  const cutoffDate = new Date();
  cutoffDate.setDate(today.getDate() - days);
  
  const data = archiveSheet.getDataRange().getValues();
  const rowsToDelete = [];
  
  for (let i = 1; i < data.length; i++) {
    const timestamp = new Date(data[i][7]); // Timestamp is in Column H (index 7)
    if (timestamp < cutoffDate) {
      rowsToDelete.push(i + 1);
    }
  }

  if (rowsToDelete.length === 0) {
    ui.alert(`No tasks found older than ${days} days.`);
    return;
  }

  const warningMessage = `WARNING:\n\nYou are about to PERMANENTLY delete ${rowsToDelete.length} task(s) older than ${days} days.\n\nThis action CANNOT be undone.\n\nAre you sure you want to proceed?`;
  const finalConfirmation = ui.alert('FINAL CONFIRMATION', warningMessage, ui.ButtonSet.YES_NO);

  if (finalConfirmation == ui.Button.YES) {
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      archiveSheet.deleteRow(rowsToDelete[i]);
    }
    
    // --- THIS IS THE FIX ---
    // We provide the title, the prompt, AND the button set.
    ui.alert('Success!', `✅ ${rowsToDelete.length} old archived task(s) have been permanently purged.`, ui.ButtonSet.OK);

  } else {
    // --- THIS IS ALSO FIXED ---
    ui.alert('Operation Cancelled', 'The purge operation was cancelled.', ui.ButtonSet.OK);
  }
}
