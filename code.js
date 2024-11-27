var ui = SpreadsheetApp.getUi();
var ss = SpreadsheetApp.getActive();
var sheet = ss.getActiveSheet();

function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  ui.createMenu('Sensitivity Analysis')
    .addItem('Run.....', 'create')
    .addToUi();
    
  // Avoid accessing properties if the user hasn't yet run the add-on (via menu items)
  // https://developers.google.com/apps-script/reference/script/auth-mode
  if (e && e.authMode != ScriptApp.AuthMode.NONE) {
    // initialize document state for datatables
    PropertiesService.getDocumentProperties().setProperty(DATATABLE_KEY, PropertiesService.getDocumentProperties().getProperty(DATATABLE_KEY) || "{}");
  }
}

// Collect input from the user 
function create() {

  // Columns
  var prompt_1 = ui.prompt("Specify Variable",
    `1. Please specify the location of the first variable that will change in your sensitivity analysis.

    2. Please specify the increments that the value should increase by each time.

    3. Please specify the amount of times you want to increase the number by. 

    Seperate it all by commas "," please and don't use spaces. 

    For example, enter "A2,0.5,4" to use the number in A2 as a starting point, increasing it by 0.5 each time for a total of 4 times.

    Leave blank if you want to use your previous input: ` + getProperty_('input_1'), 
    ui.ButtonSet.OK_CANCEL
  ); 
    

  if ( prompt_1.getSelectedButton() == ui.Button.OK) {
    var input_1;
    var response = prompt_1.getResponseText();
    
    if (response === '') {
      input_1 = getProperty_('input_1').split(",")
    } else {
      input_1 = response.split(",");
      setProperty_('input_1', response);
    }
    
    validity_checks_(input_1)
    
    var factor1_location = input_1[0]
    var factor1_value = sheet.getRange(factor1_location).getValue()
    var cols_increment = Number(input_1[1])
    var cols_num = Number(input_1[2])

    Logger.log("Factor 1 is: " + factor1_value + ". It increases by: " + cols_increment +" and does it: " + cols_num + " times.")
  } else {
    return;
  };


  // Rows
  var prompt_2 = ui.prompt("Specify Variable",
    `1. Please specify the location of the second variable that will change in your sensitivity analysis.

    2. Please specify the increments that the value should increase by each time.

    3. Please specify the amount of times you want to increase the number by. 

    Seperate it all by commas "," please and don't use spaces. 

    For example, enter "A2,0.5,4" to use the number in A2 as a starting point, increasing it by 0.5 each time for a total of 4 times.
    
    Leave blank if you want to use your previous input: ` + getProperty_('input_2'),
    ui.ButtonSet.OK_CANCEL
  );

  if (prompt_2.getSelectedButton() == ui.Button.OK) {
    var input_2;
    var response = prompt_2.getResponseText();

    if (response === '') {
      input_2 = getProperty_('input_2').split(",")
    } else {
      input_2 = response.split(",")
      setProperty_('input_2', response)
    }

    validity_checks_(input_2)

    var factor2_location = input_2[0]
    var factor2_value = sheet.getRange(factor2_location).getValue()
    var rows_increment = Number(input_2[1])
    var rows_num = Number(input_2[2])

    Logger.log("Factor 2 is: " + factor2_value + ". It increases by: " + rows_increment +" and does it: " + rows_num + " times.")
  } else {
    return;
  };

  // Model location
  var prompt_3 = ui.prompt("Specify Model Output",
    `Please specify where the output of the model is calculated (e.g. by changing the first two variables specified, the total profit is calculated.)

    Format should be "A2". 
    
    Leave blank if you want to use your previous input: ` + getProperty_('input_3'),
    ui.ButtonSet.OK_CANCEL
  );

  var model_output_location = prompt_3.getResponseText();

  if (model_output_location === '') {
    model_output_location = getProperty_('input_3')
  } else {
    setProperty_('input_3', model_output_location)
  }
  
  // Ouput location
  var prompt_4 = ui.prompt("Output Range",
    `Please specify the cell where you would like to paste the output in the format "A2". 
    
    The range will be ` +  cols_num + ` wide and ` + rows_num + ` long. 
    Warning: this may overwrite data that is already there.

    Leave blank if you want to use your previous input: ` + getProperty_('input_4'), 
    ui.ButtonSet.OK_CANCEL
  );

  var output_range_start = prompt_4.getResponseText();

  if (output_range_start === '') {
    output_range_start = getProperty_('input_4')
  } else {
    setProperty_('input_4', output_range_start)
  }

  // Calculating the different values when changing the inputs
  var array = sensitivity(factor1_value, factor1_location, cols_num, cols_increment, factor2_value, factor2_location, rows_num, rows_increment, model_output_location)

  // Defining where to paste the output
  var starting_row = sheet.getRange(output_range_start).getRow();
  var starting_col = sheet.getRange(output_range_start).getColumn();
  var range = sheet.getRange(starting_row, starting_col, rows_num + 1, cols_num + 1)
  range.setValues(array)
  
  var numColumns = range.getNumColumns();
  var numRows = range.getNumRows();
  var firstColumn = range.getColumn();
  var firstRow = range.getRow();


  // Formatting the header rows 
  var formatRows = sheet.getRange(factor1_location).getNumberFormat();
  var headerRows = sheet.getRange(firstRow, firstColumn + 1, 1, numColumns - 1)
  headerRows.setNumberFormat(formatRows);
  headerRows.setBackground('#009646');
  headerRows.setFontColor('white');
  headerRows.setFontWeight('bold');

  // Formatting the header columns 
  var formatCols = sheet.getRange(factor2_location).getNumberFormat();
  var headerCols = sheet.getRange(firstRow + 1, firstColumn, numRows - 1, 1)
  headerCols.setNumberFormat(formatCols);
  headerCols.setBackground('#009646');
  headerCols.setFontColor('white');
  headerCols.setFontWeight('bold');
  
  // Formatting the data itself
  var formatModel = sheet.getRange(model_output_location).getNumberFormat();
  var modelData = sheet.getRange(firstRow + 1, firstColumn +1 , numRows - 1, numColumns - 1)
  modelData.setNumberFormat(formatModel);

  // Resets the input factors to their original values
  sheet.getRange(factor1_location).setValue(factor1_value)
  sheet.getRange(factor2_location).setValue(factor2_value)
}

// Run the sensitivity analysis 
function sensitivity(factor1_value, factor1_location, cols_num, cols_increment, factor2_value, factor2_location, rows_num, rows_increment, model_output_location) {
  // Create an empty 2D array
  var array = []
  
  // Loop through each row and calculate the model output 
  for (var i = 0; i < rows_num; i++) {
    var temp_factor2_value = factor2_value + rows_increment * i

    array[i] = []

    // Adding in the value of the factor as the value in the first column
    array[i].push(temp_factor2_value)

    // Loop through each column in a row and calculate the model output 
    for (var j = 0; j < cols_num; j++) {
      var temp_factor1_value = factor1_value + cols_increment * j

      array[i].push(run_model_ (factor1_location, temp_factor1_value, factor2_location, temp_factor2_value, model_output_location))
    }
  }

  // Defines the header row 
  var headers = ["Sensitivity Output"]

  for (j = 0; j < cols_num; j++) {
    var temp_factor1_value = factor1_value + cols_increment * j
    headers.push(temp_factor1_value)
  }

  // Add in the header row into the array and return it
  array.splice(0,0,headers)
  return array
}

// Run the model with the current values of the factors 
function run_model_ (factor1_location, temp_factor1_value, factor2_location, temp_factor2_value, model_output_location) {
  sheet.getRange(factor1_location).setValue(temp_factor1_value)
  sheet.getRange(factor2_location).setValue(temp_factor2_value)

  return sheet.getRange(model_output_location).getValue();
}

// Check if it is a valid cell reference 
function isValidCellReference_(cellRef) {
  var regex = /^[A-Za-z]{1,2}[1-9][0-9]{0,3}?$/;
  return regex.test(cellRef);
};

// Check if it is a valid number
function isValidNumber_(num) {
  var regex = /^-?\d+(\.\d+)?$/;
  return regex.test(num)
};

// Run the different kinds of checks on the input from the user
function validity_checks_(input) {
  Logger.log("Starting checks....")
  if (isValidCellReference_(input[0]) == false) {
    ui.alert("Error", "The cell reference is not a valid one. You put in '" + input[0] + "' and something like 'A2' was expected.", 
    ui.ButtonSet.OK_CANCEL);
    Logger.log("Cell check failed");
  }

  else if (isValidNumber_(input[1]) == false) {
    ui.alert("Error", "You did not enter a valid number. You put in '" + input[1] + "' and something like '2' or '-4.5' was expected.", 
    ui.ButtonSet.OK_CANCEL);
    Logger.log("Num 1 check failed");
  }

  else if (isValidNumber_(input[2]) == false) {
    ui.alert("Error", "You did not enter a valid number. You put in '" + input[2] + "' and something like '2' or '-4.5' was expected.", 
    ui.ButtonSet.OK_CANCEL);
    Logger.log("Num 2 check failed");
  }

  else {
    Logger.log("Checks ended")
    return true
  }
};

// to save the user inputs amongst the script properties
function setProperty_ (key, value) {
  //Example: ('input_1', 'A2,23,4')
  var sp = PropertiesService.getScriptProperties(); 

  sp.setProperty(key,value)
}

function getProperty_(key) {
  var sp = PropertiesService.getScriptProperties(); 
  return sp.getProperty(key) 
}
