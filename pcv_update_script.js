/*
Macro created by Juan Manuel Montealegre Medina - juan.montealegre@bairesdev.com - 07/21/2021
Objective: Automatically creates new sheet and formulas for the new case.
*/

function  PracticalCaseLineUpdate() {

  let ss = SpreadsheetApp.getActive(); 
  let stringPCV = ss.getCurrentCell().offset(0, -1).getValue();
  let indexPCV = stringPCV.indexOf("PCV");
  let indexNext = stringPCV.indexOf("P", indexPCV + 1) - 1;           
  let numPCV = stringPCV.substring(indexPCV,indexNext);                                    
  let linkPCV = 'https://bairesdev.atlassian.net/browse/' + numPCV;
  
  // gets name of the candidate
  indexPCV = stringPCV.indexOf("| ", stringPCV.indexOf("Practical Case |"))+2;
  let getCandidate = stringPCV.substring(indexPCV,150);
  
  indexPCV = stringPCV.indexOf("P", indexNext);
  indexNext = stringPCV.indexOf(" |");
  let caseType = stringPCV.substring(indexPCV,indexNext);
  let currentRow = ss.getActiveRange().getRow();
  let dateCell = "G"+ currentRow
  // verifies that macro is being run on a blank cell
  let cell = ss.getCurrentCell().getValue();

  ss.getCurrentCell().offset(0, 5).setFormula("=TODAY()");


  if (cell == '') {
      
      // pastes link from JIRA ticket
      let richValue = SpreadsheetApp.newRichTextValue()
      .setText(numPCV)
      .setLinkUrl(linkPCV)
      .build();
    ss.getCurrentCell().setRichTextValue(richValue);

    // pastes name of candidate
    ss.getCurrentCell().offset(0, 1).setValue(getCandidate);
    ss.getCurrentCell().offset(0, -1).activate();
    ss.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
    
    if(caseType === 'PMO'){
      ss.getCurrentCell().setValue(caseType);
      let getCandidateSheetFormulaPMO = '='+ "''" + getCandidate + "'" +'!E15';
      
      // duplicates "Original" sheet and changes sheet name to candidate name
      let original = ss.getSheetByName('Original');
      original.copyTo(ss).setName(getCandidate);

      // brings candidate score
      ss.getCurrentCell().offset(0, 3).activate();                                
      ss.getCurrentCell().setFormulaR1C1(getCandidateSheetFormulaPMO);
  } else {
      ss.getCurrentCell().setValue(caseType);
      let getCandidateSheetFormulaPPI = '='+ "''" + getCandidate + "'" +'!E17';
      
      // duplicates "Original - Old Case" sheet and changes sheet name to candidate name
      let original = ss.getSheetByName('Original - Old Case');
      original.copyTo(ss).setName(getCandidate);

      // brings candidate score
      ss.getCurrentCell().offset(0, 3).activate();
      ss.getCurrentCell().setFormulaR1C1(getCandidateSheetFormulaPPI);      
    }

  ss.getRange(dateCell).activate();
  
  // sets today's date
  ss.getRange(dateCell).copyTo(ss.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  ss.getActiveRangeList().setNumberFormat('M/d/yyyy');

  ss.getRange("D" + currentRow).activate();

  } else {
    SpreadsheetApp.getUi().alert("The cell you selected is not empty. Please select an empty cell and run the macro again.");
  }
}
