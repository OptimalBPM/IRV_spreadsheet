// This is a Google Spreadsheet function for managing instant run-off (IRV) voting.
// Copyright 2017 Nicklas Börjesson, distributed under an MIT license
// It requires the first sheet to list allowed candidates, the second the votes (timestamp, first choice ... nth choice), and a third to hold the results

function runIRV() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var candidateSheet = ss.getSheets()[0];
  var voteSheet = ss.getSheets()[1];
  var resultSheet = ss.getSheets()[2];
  var lastColumn = voteSheet.getLastColumn()
  var candidates = (candidateSheet.getRange(2, 1, candidateSheet.getLastRow()-1, 1)).getValues();
  var range = voteSheet.getRange(2, 2, voteSheet.getLastRow()-2,lastColumn-1);
  var values = range.getValues();
  var numVoters = values.length;
  var buckets = [];
  var eliminated = [];
  
  resultSheet.clear();
  
  // Function for logging stuff to the result sheet.
  function logToSheet(value) {
    resultSheet.getRange(resultSheet.getLastRow()+1,1,1,value.length).setValues([value]);
  }
  
  Logger.log(values);
  // Start with the first round and work until done
  var currResults = [];
  
  var tally = {}
  
  // Add all valid candidates
  for (var currCandidate in candidates) {
    tally[candidates[currCandidate][0]] = 0;
  }
  logToSheet(Array.concat(["candidates"], [candidates.toString()]));
  
  // TODO: Loop and check so that all are valid
  //Alert("Invalid candidates");
  // Mark those red.

  // Loop all rounds, start with the first choice votes
  for (var round = 0; round < lastColumn-1; round++) { 
    // Collect all with the current round
    for (var row in values) { 
      // Ignore eliminated candidates
      if (eliminated.indexOf(values[row][round])== -1) {
        tally[values[row][round]]++;
      };
    };
    logToSheet(["Round " + round.toString(), tally.toString()]);

    // Log the current tally
    for (var currVote in tally) {
      logToSheet(["", "Vote "+ currVote + " " + tally[currVote]]);
    };
    
    // Sort by votes
    keysSorted = Object.keys(tally).sort(function(a,b){return tally[b]-tally[a]})
    Logger.log(keysSorted.toString());
    // Summarize tally
    var totalTally = 0
    for (currTally in tally) {
      totalTally+= tally[currTally];
    };
    // Anyone have a majority? 
    if (tally[keysSorted[0]] > totalTally/2) {
      logToSheet(["The winner:", keysSorted[0]]);
    }
    else {
      logToSheet(["No winner yet, highest is " + keysSorted[0] + " with " + tally[keysSorted[0]] + "votes of " + totalTally]);
    }
    
    // Eliminate bottom if bottom has less than top
    bottom = keysSorted[keysSorted.length-1];
    if (tally[keysSorted[0]] > tally[bottom]){
      logToSheet([bottom + " was eliminated in round number " + round]);
      delete tally[bottom];
      eliminated.push(bottom);
    }
   
  }
}
