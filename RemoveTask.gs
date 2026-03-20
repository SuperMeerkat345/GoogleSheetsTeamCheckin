// Shows the form for removing tasks
// Adds the html from RemoveTaskForm.html to the modal
function CallRemoveTaskForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // get the form to put into the alert
  const html = HtmlService.createHtmlOutputFromFile('RemoveTaskForm')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  ui.showModalDialog(html, 'Remove Task');
}
// Physically removes the task from the form to the sheet
function RemoveTask(task_id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MAIN");
  const secrets = ss.getSheetByName("SENSITIVE")

  const num_tasks_cell = secrets.getRange("O2");

  const values = sheet.getRange("D:D").getValues().flat();
  const index = values.indexOf(Number(task_id)); 

  if (index === -1) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Task ID not found.");
    return;
  }

  const currentRow = index + 1;

  num_tasks_cell.setValue(num_tasks_cell.getValue() - 1);

  const task = sheet.getRange(`D${currentRow}:H${currentRow}`);
  // 2. Clear the dropdown chip from Column H (the 5th column in this range)
  sheet.getRange(currentRow, 8).setDataValidation(null)
  task.setValues([["", "", "", "", ""]]);
  ShiftUpTasks();

  SpreadsheetApp.getActiveSpreadsheet().toast("Task removed successfully.");
}
// Moves up all rows until they hit the top 
// Used after removing a row
function ShiftUpTasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName("MAIN")
  
  let data_range = sheet.getRange("D:H")
  let data = data_range.getValues() // returns 2d arr

  // filters out row with no elements
  let filtered_data = data.filter(arr => {
    return arr.some(cell => cell !== "")
  })

  data_range.clearContent();
  sheet.getRange(1, 4, filtered_data.length, 5).setValues(filtered_data);
}
