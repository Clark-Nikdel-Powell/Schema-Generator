/**
 * @OnlyCurrentDoc Limits the script to only accessing the current sheet.
 */

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
