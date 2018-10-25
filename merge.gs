/*  This is the main method that should be invoked. 
 *  Copy and paste the ID of your template Doc in the first line of this method.
 *
 *  Make sure the first row of the data Sheet is column headers.
 *
 *  Reference the column headers in the template by enclosing the header in square brackets.
 *  === Example ===
 *  worksheet values:
 *   name    address1         address2
 *   Alice   123 Paint Dr     Springfield, MO, 60000
 *   Bob     22 Twain House   Murderville, MO, 69999
 *  template text:
 *   "[name], I need you to proceed to [address1], [address2] immediately. Sincerely, The FEDs"
 */
// documentation for apps script for google sheets is at:
// https://developers.google.com/apps-script/reference/spreadsheet/
function doMerge() {
  //Copy and paste the ID of the template document here (you can find this in the document's URL)
  var templateFile = DriveApp.getFileById("1foobarfoobarfoobarfoobarfoobarfoobar");

  var ui = DocumentApp.getUi();
  var response =
    ui.prompt('Confirm Template',
              'This will merge into template document "' + templateFile.getName() +
              '". If this is not what you want to use, please input the documentId of the template file.',
              ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK)
    return;
  var inputTemplateId = response.getResponseText();
  if (inputTemplateId !== ''){
    templateFile = DriveApp.getFileById(inputTemplateId);
  }
  
  //make a copy of the template file to use for the merged File.
  // (Note: It is necessary to make a copy upfront, and do the rest of the content manipulation inside this single copied file.
  //  Otherwise, if the destination file and the template file are separate,
  //   a Google bug will prevent copying of images from the template to the destination.
  //  This bug is tracked at: https://code.google.com/p/google-apps-script-issues/issues/detail?id=1612#c14)
  var mergedFile = templateFile.makeCopy();
  //give a custom name to the new file (otherwise it is called "copy of ...")
  mergedFile.setName("filled_"+templateFile.getName());
  var mergedDoc = DocumentApp.openById(mergedFile.getId());
  //the body of the merged document, which is at this point the same as the template doc.
  var bodyElement = mergedDoc.getBody();
  //make a copy of the body
  var bodyCopy = bodyElement.copy();
  
  //clear the body of the mergedDoc so that we can write the new data in it.
  bodyElement.clear();
  
  var currentSheet = SpreadsheetApp.getActiveSheet();
  var rows = currentSheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  var fieldNames = values[0];

  for (var i = 1; i < numRows; i++) {//data values start from the second row of the sheet 
    var row = values[i];
    var body = bodyCopy.copy();
    
    for (var f = 0; f < fieldNames.length; f++) {
      body.replaceText("\\[" + fieldNames[f] + "\\]", row[f]);//replace [fieldName] with the respective data value
    }
    
		//number of the contents in the template doc
    var numChildren = body.getNumChildren();
    //go through all the content of the template doc, and replicate it for each row of the data
    for (var c = 0; c < numChildren; c++) {
      var child = body.getChild(c);
      child = child.copy();
      if (child.getType() == DocumentApp.ElementType.HORIZONTALRULE) {
        mergedDoc.appendHorizontalRule(child);
      } else if (child.getType() == DocumentApp.ElementType.INLINEIMAGE) {
        mergedDoc.appendImage(child.getBlob());
      } else if (child.getType() == DocumentApp.ElementType.PARAGRAPH) {
        mergedDoc.appendParagraph(child);
      } else if (child.getType() == DocumentApp.ElementType.LISTITEM) {
        mergedDoc.appendListItem(child);
      } else if (child.getType() == DocumentApp.ElementType.TABLE) {
        mergedDoc.appendTable(child);
      } else {
        Logger.log("Unknown element type: " + child);
      }
   }
   //TODO - change to column break?
   mergedDoc.appendPageBreak();//Appending page break. Each row will be merged into a new page.

  }
}

function highlightProblems() {
  // this is experimental code, just for my own specific spreadsheet.
  // TODO - either make more broadly useful, rename project (SheetsToDocsMerge) to fit the scope, or remove this from merge.gs
  var currentSheet = SpreadsheetApp.getActiveSheet();
  var rows = currentSheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  var fieldNames = values[0];

	// TODO: validations:
	// cross-check "Number of children in household receiving gifts" against "Placement Household" + "Email Address"
	// identify duplicates -- "Date of Birth", in red if also match "First Name of child"+"Email Address"
	var iName, iDOB, 
  for (var f = 0; f < fieldNames.length; f++) {
    body.replaceText("\\[" + fieldNames[f] + "\\]", row[f]);//replace [fieldName] with the respective data value
  }
	// proof of concept -- highlight in yellow any cells with values "NA", "N/A", "?", or "-"
  for (var i = 1; i < numRows; i++) {
    var row = values[i];
    for (var f = 0; f < fieldNames.length; f++) {
      switch (row[f].toUpperCase()) {
				case 'NA': case 'N/A':
				case '?': case '-':
          var thisCell = currentSheet.getRange(row, 1);
					thisCell.setBackground("yellow");
	  			break;
			}
    }
	}
	// TODO: auto-apply list of alias names
	// 0) get list of top-1000 names
	// 1) find the "closest" match in top-1000 names, use as alias, and remove from list
	// 1.a) code "closest" match function -- Levenshtein distance? try code at: https://gist.github.com/andrei-m/982927
}

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the doMerge() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Fill template",
    functionName : "doMerge"
  },{
    name : "Highlight duplicates",
    functionName : "highlightProblems"
  },];
  spreadsheet.addMenu("Merge", entries);
};
