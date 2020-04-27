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

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
    var spreadsheet = SpreadsheetApp.getActive();
    var menuItems = [
        {name: 'Generate JSON', functionName: 'generateSchemaJSON_'}
    ];
    spreadsheet.addMenu('SCHEMA', menuItems);
}

function openTag(output) {
    return output += '&lt;script type="application/ld+json"&gt;\n{\n\t"@context": "http://schema.org",\n';
}

function closeTag(output) {
    return output += '}\n&lt;/script&gt;\n';
}

function generateSchemaJSON_() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];

    var rows = sheet.getDataRange().getValues();
    var output = openTag('');
    var openID = false;
    var openObject = false;

    for (r = 1; r < rows.length; ++r) {

        // Variable setup
        var sId = rows[r][2];
        var sType = rows[r][3];
        var sProperty = rows[r][4];
        var sValue = rows[r][5];

        // Column 2: ID
        if ('' !== sId) {
            if (true === openID) {
                output = closeTag(output);
                output = openTag(output);
            }
            output += '\t"@id": "' + sId + '",\n';
            openID = true;
        }


        // Column 3: Type
        if ('' !== sType) {
            if (true === openObject) {
                output += '\t';
            }
            output += '\t"@type": "' + sType + '",\n';
        }


        // Column 4 & 5: Property and Value

        // If value has a "#" in it, then it references an ID on the next row.
        if ('' !== sValue && 0 === sValue.indexOf('#')) {
            // We need to add 3 things here: the property, the value as an @id property, and the type from the next row down, as a @type property.
            // Output should be like this:
            // "address": {
            //     "@type": "PostalAddress",
            //     "@id": "#address"
            // },
            output += '\t"' + sProperty + '": {\n\t\t"@type": "' + rows[r + 1][3] + '",\n\t\t"@id": "' + sValue + '"\n\t},\n';

            continue; // Skip any other cases for property and value.
        }

        // If property is present but value isn't, then we're setting up a new object in the next row.
        // Need to close this object later on.
        if ('' !== sProperty && '' === sValue) {
            output += '\t"' + sProperty + '": {\n';
            openObject = true;
        }

        // If both name and value are present, then add them as they are.
        if ('' !== sProperty && '' !== sValue) {

            if (true === openObject) {
                output += '\t';
            }

            // TODO: update this for any array values.
            if ('sameAs' === sProperty) {
                var items = sValue.split(',');
                output += '\t"' + sProperty + '": [';
                for (i = 0; i < items.length; ++i) {
                    if (0 === i) {
                        output += '\n';
                    }
                    output += '\t\t"' + items[i] + '"';
                    if (i !== items.length - 1) {
                        output += ',';
                    }
                    output += '\n';
                }
                output += '\t],\n';
            } else {
                output += '\t"' + sProperty + '": "' + sValue + '"';

                if ('' !== rows[r + 1][5]) {
                    output += ',';
                }
                output += '\n';
            }
        }
    }

    if (true === openObject) {
        output += '\t}\n';
        openObject = false;
    }

    output = closeTag(output);
    output = '<code style="white-space:pre;">' + output + '</code>';

    var htmlOutput = HtmlService
        .createHtmlOutput(output)
        .setWidth(600)
        .setHeight(800);

    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Generated Schema');
}
