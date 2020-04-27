/**
 * @OnlyCurrentDoc Limits the script to only accessing the current sheet.
 */

// onEdit is a simple trigger
// https://developers.google.com/apps-script/guides/triggers
function onEdit(e) {      
    
    // Use Logger to log values, e.g., Logger.log(e)
    // https://developers.google.com/apps-script/guides/logging
    
    if ('' === e.value ) { // If the cell was cleared out, don't do anything.
      return false;
    }
    
    // Working with range: https://developers.google.com/apps-script/reference/spreadsheet/range   
    if (4 !== e.range.getColumn()) { // Looking for changes to column D.
      return false;
    }
    
    if (1 < e.range.getNumColumns() || 1 < e.range.getNumRows()) { // If more than one column is effected, bail out.
      return false;
    }
    
    // Get the Type value that's been set.
    let value = e.value;
    
    // Next, we're looking to get the available Properties for the selected Type.
    // These are found in Column G of the Types Sheet, where Column B is the Type Name.
    // So the plan is: get the Type value, find that Type value in Column B of the Types Sheet, then get the contents of Column G.
    const spreadsheet = SpreadsheetApp.getActive();
    const targetSheet = spreadsheet.getSheetByName("Types");    
    // This gets all the values for Column B, then finds the index of the value.
    const typesRow = targetSheet.getRange(1, 2, 1129).getValues().findIndex(function(r){return r[0] === value;}) + 1;
    
    // This gets all the values for Column G, then removes the Schema.org URL.
    let typesRowProperties = targetSheet.getRange(typesRow, 7).getValues()[0][0];   
    let cleanedProperties = typesRowProperties.replace(/http:\/\/schema.org\//g, '');
    
    // We need to know how many cells to which to apply this validation. In order to do that, we have to start at the row that changed,
    // then go down the column until we find another value. If no value is found, we use getLastRow.
    // Get the row where the value was changed.
    let row = e.range.getRow();    
    
    // Get all the rows below this one.
    // Get all the cells in this column.
    const column = spreadsheet.getRange('D:D');
    // Then get the values.
    const values = column.getValues();    
    // Use findIndex to find the index of the next filled row.
    let nextFilledRow = values.findIndex(function(r, index){      
    
      if ( index <= row - 1 ) {
        return false;
      }
    
      if ( '' === r[0] ) {
        return false;
      }
      
      return true;
    });
    
    // If no rows were found, then we use the last row with data.
    if (-1 === nextFilledRow) {
      nextFilledRow = spreadsheet.getLastRow();
    } else {
      nextFilledRow++; // findIndex starts at 0. If we have a valid response from that, we need to add 1.
    }
    
    // We can define our range for data validation: the row to start on, the column (E), and then the number of rows.
    const cellsForDataValidation = spreadsheet.getSheetByName('Schemas').getRange(row, 5, nextFilledRow - row);
    // Define and build the data validation.
    const propertiesRule = SpreadsheetApp.newDataValidation().requireValueInList(cleanedProperties.split(","), true).build();
    cellsForDataValidation.setDataValidation(propertiesRule);
}

// Runs when the spreadsheet is open.
function onOpen() {
    const spreadsheet = SpreadsheetApp.getActive();
    // Add a menu item, a button and a callback function.
    const menuItems = [
        {name: 'Generate JSON', functionName: 'generateSchemaJSON_'}
    ];
    spreadsheet.addMenu('SCHEMA', menuItems);
}

// Opens a script tag for a JSON-LD block. The HTML entities are used so that the script is readable instead of executed when it's printed out.
function openTag(output) {
    return output += '&lt;script type="application/ld+json"&gt;\n';
}

// When closing, we take the JSON value and add it in. JSON and Output are kept separate for the most part so that we can debug JSON on its own.
function closeTag(jsonString, output) {   
    return output += cleanData(jsonString + '}') + '\n&lt;/script&gt;\n';
}

// JSON does NOT like trailing commas, so we clean them out here.
// I also had a parse JSON check here, but something wasn't work, so I've yanked it for now.
function cleanData(jsonString) {
    let regex = /\,(?!\s*?[\{\[\"\'\w])/g;
    let correct = jsonString.replace(regex, ''); // remove all trailing commas
    return correct; // build a new JSON object based on correct string
}

// Runs when somebody clicks "Generate JSON"
function generateSchemaJSON_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets()[0];

    const rows = sheet.getDataRange().getValues(); // This gets all the data from the spreadsheet in a multidimensional array.
    let output = openTag(''); // Start the open tag.
    let jsonString = '{\n\t"@context": "http://schema.org",\n'; // Start the first JSON block.
    let openID = false;
    let openObject = false;

    // Loop through the spreadsheet one row at a time.
    for (r = 1; r < rows.length; ++r) {

        // Variable setup
        const sId = rows[r][2];
        const sType = rows[r][3];
        const sProperty = rows[r][4];
        const sValue = rows[r][5];

        // Column 2: ID
        if ('' !== sId) {
            if (true === openID) { // If the ID exists, but there's already an open ID, we need to close the tag.
                output = closeTag(jsonString, output);
                output = openTag(output);
                jsonString = '{\n\t"@context": "http://schema.org",\n'; // Reset the jsonString since we're starting over.
            }
            jsonString += '\t"@id": "' + sId + '",\n'; // Add the ID in to the JSON.
            openID = true; // The openID variable really just exists for the first ID, so we don't close the first block before putting anything in.
        }


        // Column 3: Type
        if ('' !== sType) {
            if (true === openObject) {
                jsonString += '\t'; // Just for formatting, add another tab if there's an open object.
            }
            jsonString += '\t"@type": "' + sType + '",\n';
        }


        // Column 4 & 5: Property and Value

        // If value has a "#" in it, then it references an ID on the next row.
        if ('' !== sValue && 0 === sValue.indexOf('#')) {           
            jsonString += '\t"' + sProperty + '": {\n\t\t"@id": "' + sValue + '"\n\t},\n';

            continue; // Skip any other cases for property and value.
        }

        // If property is present but value isn't, then we're setting up a new object in the next row.
        // Need to close this object later on.
        // TODO: unclear what this refers to. Add example.
        if ('' !== sProperty && '' === sValue) {
            jsonString += '\t"' + sProperty + '": {\n';
            openObject = true;
        }

        // If both name and value are present, then add them as they are.
        if ('' !== sProperty && '' !== sValue) {

            if (true === openObject) {
                jsonString += '\t';
            }

            // TODO: update this for any array values.
            if ('sameAs' === sProperty) {
                var items = sValue.split(',');
                jsonString += '\t"' + sProperty + '": [';
                for (i = 0; i < items.length; ++i) {
                    if (0 === i) {
                        jsonString += '\n';
                    }
                    jsonString += '\t\t"' + items[i] + '"';
                    if (i !== items.length - 1) {
                        jsonString += ',';
                    }
                    jsonString += '\n';
                }
                jsonString += '\t],\n';
            } else {
                jsonString += '\t"' + sProperty + '": "' + sValue + '"';

                if ('' !== rows[r + 1][5]) { // Look at the next row down in order to see if we need a comma.
                    jsonString += ',';
                }
                jsonString += '\n';
            }
        }
    }

    if (true === openObject) {
        output += '\t}\n';
        openObject = false;
    }

    output = closeTag(jsonString, output);
    output = '<code style="white-space:pre;">' + output + '</code>';

    // Use the HtmlService class to add the output.
    var htmlOutput = HtmlService.createHtmlOutput(output).setWidth(600).setHeight(800);

    // Output the modal dialog.
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Generated Schema');
}
