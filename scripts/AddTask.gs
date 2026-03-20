// Shows the form for adding tasks
// Adds the html from AddTaskForm.html to the modal
function CallAddTaskForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // get the form to put into the alert
  const html = HtmlService.createHtmlOutputFromFile('AddTaskForm')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  ui.showModalDialog(html, 'Add Task');
}
function GetValidTeams() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const secrets = ss.getSheetByName("SENSITIVE")

  let teams_range = secrets.getRange("E2:E")
  let teams = teams_range.getValues().flat().filter(elem => elem !== "")
  let teams_unique = new Set(teams) // get unique
  
  // turn back into array and sort
  return Array.from(teams_unique).sort((a, b) => a - b);
}
function GetValidScouters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const secrets = ss.getSheetByName("SENSITIVE")

  let scouters_range = secrets.getRange("B2:B")
  let scouters = scouters_range.getValues().flat().filter(elem => elem !== "")

  return scouters.sort()
}
// Physically adds the task from the form to the sheet
function AddTask(task_scouter, task_team_num, task_description, task_status, task_ping) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MAIN");
  const secrets = ss.getSheetByName("SENSITIVE");

  const num_tasks_cell = secrets.getRange("O2")
  const last_task_cell = secrets.getRange("P2")

  const currentRow = Number(num_tasks_cell.getValue()) + 2 // num_tasks + offset
  const task_id = Number(last_task_cell.getValue()) + 1 // next id
  const scouter = GetScouter(task_scouter, task_team_num)

  if (scouter === "queue_len_err") {
    ss.toast("Failed to add task: queue_len_err");
    return;
  }

  num_tasks_cell.setValue(num_tasks_cell.getValue()+1)
  last_task_cell.setValue(task_id)

  const task = sheet.getRange(`D${currentRow}:H${currentRow}`)
  task.setValues([[task_id, scouter, task_team_num, task_description, task_status]])
  
  const statuses = ["Pending", "In Progress", "Completed", "Blocked"];
  const colors = ["#grey", "#0000ff", "#00ff00", "#ff0000"]; // Optional visual reference

  // create the validation rule for statuses
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statuses)
    .setAllowInvalid(false)
    .build();

  // Apply to Column H (8) of the currentRow
  const statusCell = sheet.getRange(currentRow, 8);
  statusCell.setDataValidation(rule);
  
  // Set the initial value so it's not empty
  statusCell.setValue(task_status);

  ss.toast("Task #" + task_id + " assigned to " + scouter);

  // send slack msg
  if (task_ping) {
    PingScouter(task_scouter, task_team_num, task_description)
  }
}
// Processes the scouter
// does the queue if its a queue
function GetScouter(task_scouter, task_team_num) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName("MAIN")
  const secrets = ss.getSheetByName("SENSITIVE")

  if (task_scouter === "smart_queue") {
    let teams = secrets.getRange("M2:M").getValues().flat().filter(item => item !== "");
    let scouters = secrets.getRange("N2:N").getValues().flat().filter(item => item !== "");

    // if alr scouted a team, they get auto assigned to them
    let index_scouter = teams.map(String).indexOf(String(task_team_num));
    if (index_scouter !== -1) {
      return scouters[index_scouter]
    }

    // shift queue if that didnt work
    let q = ShiftQueue(sheet);
    if (q === "queue_len_err") return q;

    // append team: scouter DB
    let nextRow = teams.length + 2; // +2 because of 1-based indexing and header row
    secrets.getRange(nextRow, 13).setValue(task_team_num); 
    secrets.getRange(nextRow, 14).setValue(q); 

    return q;            
  }
  else if (task_scouter === "queue") {
    return ShiftQueue(sheet)
  }
  else {
    return task_scouter
  }
}
function ShiftQueue(sheet) {
  let queue = sheet.getRange("B2:B").getValues().flat().filter(item => item !== "");
  let working = sheet.getRange("C2:C").getValues().flat().filter(item => item !== "");
      
  if(queue.length == 0) {
    return "queue_len_err"
  }
    
  let newWorker = queue.shift(); 
  working.push(newWorker);

  sheet.getRange("B2:C" + sheet.getLastRow()).clearContent();

  if (queue.length > 0) {
    sheet.getRange(2, 2, queue.length, 1).setValues(queue.map(item => [item]));
  }
    
  if (working.length > 0) {
    sheet.getRange(2, 3, working.length, 1).setValues(working.map(item => [item]));
  }

  return newWorker
}
function SetupStatusColors() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MAIN");
  const range = sheet.getRange("H2:H"); // Apply to the whole status column
  
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Completed")
      .setBackground("#b7e1cd") // Light Green
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("In Progress")
      .setBackground("#c9daf8") // Light Blue
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Blocked")
      .setBackground("#f4cccc") // Light Red
      .setRanges([range])
      .build()
  ];
  
  sheet.setConditionalFormatRules(rules);
}
// on literally any edit don't need to run auto does it
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  
  // 1. Only run if we are on "MAIN" and editing Column H (Status)
  if (sheet.getName() !== "MAIN" || range.getColumn() !== 8) return;
  
  const newValue = e.value;
  if (newValue !== "Completed") return;

  const row = range.getRow();
  const scouterName = sheet.getRange(row, 5).getValue(); // Column E: Scouter
  if (!scouterName) return;

  // 2. Check all tasks for this scouter
  const data = sheet.getRange("E2:H" + sheet.getLastRow()).getValues();
  const hasIncompleteTasks = data.some(row => {
    return row[0] === scouterName && row[3] !== "Completed";
  });

  // 3. If they are finished with everything, move them
  if (!hasIncompleteTasks) {
    MoveWorkerToQueue(sheet, scouterName);
  }
}
function MoveWorkerToQueue(sheet, scouterName) {
  // Get current lists
  let queue = sheet.getRange("B2:B").getValues().flat().filter(item => item !== "");
  let working = sheet.getRange("C2:C").getValues().flat().filter(item => item !== "");

  // Remove from working, add to queue
  const workerIndex = working.indexOf(scouterName);
  if (workerIndex !== -1) {
    working.splice(workerIndex, 1);
    queue.push(scouterName);

    // Update the sheet
    sheet.getRange("B2:C" + Math.max(sheet.getLastRow(), 2)).clearContent();
    
    if (queue.length > 0) {
      sheet.getRange(2, 2, queue.length, 1).setValues(queue.map(item => [item]));
    }
    if (working.length > 0) {
      sheet.getRange(2, 3, working.length, 1).setValues(working.map(item => [item]));
    }
    
    SpreadsheetApp.getActiveSpreadsheet().toast(scouterName + " is back in the queue!", "Tasks Finished");
  }
}
function ResetStatusColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MAIN");
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return;

  // This clears content, formats, and data validation (the chips)
  sheet.getRange("H2:H" + lastRow).clear();
}

function test1() {
  console.log("test")
  console.log(GetValidTeams())
}
