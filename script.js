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


function generateSchemaJSON_ () {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
    
  var rows = sheet.getDataRange().getValues(); 
  var output = '&lt;script type="application/ld+json"&gt;\n{\n\t"@context": "http://schema.org",\n';
  var openObject = false;
  
  for(r=1;r<rows.length;++r){
    
    // First, check for an ID.
    var ID = rows[r][2];
    
    if ( '' !== ID ) {
      output += '\t"@id": "' + ID + '",\n';     
    }
    
    // Checks to see if type is present.
    if ( '' !== rows[r][3] ) {      
        if ( true === openObject ) {
          output += '\t';
        }
        output += '\t"@type": "'+ rows[r][3] +'",\n';
    }
    
    // If property is present but value isn't, then we're setting up a new object in the next row.
    // Need to close this object later on.
    if ( '' !== rows[r][4] && '' === rows[r][5] ) {
        output += '\t"'+ rows[r][4] +'": {\n';
        openObject = true;
    }
    
    // If both name and value are present, then add them as they are.
    if ( '' !== rows[r][4] && '' !== rows[r][5] ) {
    
        if ( true === openObject ) {
          output += '\t';
        }
    
        if ( 'sameAs' == rows[r][4] ) {
          var items = rows[r][5].split(',');       
          output += '\t"'+ rows[r][4] +'": [';
          for (i=0;i<items.length;++i) {
            if (0 === i) {
              output += '\n';
            }
            output += '\t\t"'+ items[i] +'",\n';
          }
          output += '\t],\n';
        } else {
          output += '\t"'+ rows[r][4] +'": "'+ rows[r][5] +'",\n';
        }
    }
  }
  
  if (true === openObject) {
    output += '\t}\n';
    openObject = false;
  }
  
  output += '}\n&lt;/script&gt;';
  output = output.replace(',\n]', '\n]').replace(',\n}', '\n}').replace(',\n\t]', '\n\t]').replace(',\n\t}', '\n\t}');
  output = '<code style="white-space:pre;">' + output + '</code>';
  
  var htmlOutput = HtmlService
    .createHtmlOutput(output)
    .setWidth(600)
    .setHeight(800);  
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Generated Schema');
}
