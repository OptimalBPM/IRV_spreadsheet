function createForm() {
  // Create a new form, with the possibility to vote
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName('Inställningar');
  
  var allCandidates = (settingsSheet.getRange(2, 1, settingsSheet.getLastRow()-1, 1)).getValues();
  var numberOfOptions = (settingsSheet.getRange(1, 4, 1, 1)).getValues();  
  var title = (settingsSheet.getRange(3, 4, 1, 1)).getValues();  
  
  var form = FormApp.create(title);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  for (var currLevel=0; currLevel < numberOfOptions; currLevel++)
  {
    var item = form.addListItem();
    item.setTitle(Number(currLevel +1) +".");
    
    var choices = [];
    for (var currCandidate in allCandidates) {
      choices.push(item.createChoice(allCandidates[currCandidate]))
    }  
  
    item.setChoices(choices);    
  }
  
  settingsSheet.getRange(4,4,2,1).setValues([[form.getEditUrl()],[form.getPublishedUrl()]]);
  Logger.log('Published URL: ' + form.getPublishedUrl());
  Logger.log('Editor URL: ' + form.getEditUrl());
}
