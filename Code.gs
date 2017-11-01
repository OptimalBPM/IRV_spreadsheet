// This is a Google Spreadsheet function for managing instant run-off (IRV) voting.
// Copyright 2017 Nicklas Börjesson, distributed under an MIT license
// It requires the first sheet to list allowed candidates, the second the votes (timestamp, first choice ... nth choice), and a third to hold the results

function runIRV() {// This is a Google Spreadsheet function for managing instant run-off (IRV) voting.
// Copyright 2017 Nicklas Börjesson, distributed under an MIT license
// It requires the first sheet to list allowed candidates, the second the votes (timestamp, first choice ... nth choice), and a third to hold the results

function runIRV() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi()
  var settingsSheet = ss.getSheets()[0];
  var voteSheet = ss.getSheets()[1];
  var resultSheet = ss.getSheets()[2];
  var lastColumn = voteSheet.getLastColumn()
  var allCandidates = (settingsSheet.getRange(2, 1, settingsSheet.getLastRow()-1, 1)).getValues();
  var numberOfOptions = (settingsSheet.getRange(1, 4, 1, 1)).getValues();

  var range = voteSheet.getRange(2, 2, voteSheet.getLastRow()-1,lastColumn-1);
  var values = range.getValues();
  var numVoters = values.length;
  var buckets = [];
  
  resultSheet.clear();

  // Initially, all candidates are remaining but we want to keep the allCandidates range and have this array instead.
  var remaining = [];
  for (var currCandidate in allCandidates) {
    remaining.push(allCandidates[currCandidate][0]);
  }
  
  // Function for logging stuff to the result sheet.
  function logToSheet(value) {
    resultSheet.getRange(resultSheet.getLastRow()+1,1,1,value.length).setValues([value]);
  }
  
  function calculateTallies(candidates, level) {
    // Calculate tallies for selected candidates and sort them, also return a total of votes
    
    // Initialize tally with all candidates
    var tally = {}
    for (var currCandidate in candidates) {
      tally[candidates[currCandidate]] = 0;
    }    
    
    for (var row in values) { 
      // Only count selected candidates
      Logger.log("values[row][level] = " + values[row][level]);
      if (candidates.indexOf(values[row][level]) > -1 ) {
        tally[values[row][level]]++;
      };
    };       
    // Sort by votes
    voteSorted = Object.keys(tally).sort(function(a,b){return tally[b]-tally[a]})
    Logger.log(voteSorted.toString());
    // Summarize tally
    var totalTally = 0;
    for (currTally in tally) {
      totalTally+= tally[currTally];
    };
    
    return [tally, voteSorted, totalTally]
  }
  
  function returnBottom(candidates, tally) {
    // Return the bottom candidate(s) from a list sorted by votes
    
    var bottom = []
    var currCandidate = candidates.length -1;
    var currVotes = tally[candidates[currCandidate]];
    bottom.push(candidates[currCandidate]);
    while (currCandidate > 0 && currVotes == tally[candidates[currCandidate - 1]] ) {
      currCandidate--;
      bottom.push(candidates[currCandidate]);
    }
    return bottom;
  }
  
  
  function eliminateBottom(candidates, level, round) {
    // Calculate tallies
    var tallies = calculateTallies(candidates, level);
    var tally = tallies[0];
    var voteSorted = tallies[1];
    var totalTally = tallies[2];
    
    var bottom = returnBottom(voteSorted, tally);
    if (bottom.length == 1) {
      remaining.splice(remaining.indexOf(bottom[0]),1);
      logToSheet([bottom[0] + " was eliminated in round " + round]);
    }
    else {
      if (level < numberOfOptions) {
        eliminateBottom(bottom, level+1, round)
      }
      else {
        logToSheet(["Asking user to choose who to eliminate among " + bottom.toString()]);
        var buildText = "Enter a number to choose who to eliminate:\n";
        for (var currCandidate in bottom) {
          buildText += "[" + currCandidate.toString() +"] "+ bottom[currCandidate] + "\n";
        }
        
        var result = ui.prompt('Input required', buildText, ui.ButtonSet.OK_CANCEL);
        var button = result.getSelectedButton();
        var text = result.getResponseText();
        if (button == ui.Button.OK) {
          // User clicked "OK".
          remaining.splice(remaining.indexOf(bottom[Number(text)]),1);
      
          logToSheet([bottom[Number(text)] + " was manually chosen to be eliminated in round " + round]);

        } else if (button == ui.Button.CANCEL) {
          // User clicked "Cancel".
          logToSheet(["User exited the procedure"]);
          throw new Error("User exited the procedure");
        } else if (button == ui.Button.CLOSE) {
          // User clicked X in the title bar.
          logToSheet(["User exited the procedure"]);
          throw new Error("User exited the procedure");

        }

      }
    }
  }
      
  
  Logger.log(values);
  // Start with the first round and work until done
  var currResults = [];

  logToSheet(Array.concat(["candidates"], [remaining.toString()]));
  
  // TODO: Loop and check so that allvotes have valid values
  //Alert("Invalid candidates");
  // Mark those red.
  var round = 1;
  var done = false;
  // Loop until we are done
  while (!done) { 

    logToSheet(["Round " + round.toString()]);
    
    // Calculate tallies for first votes
    var tallies = calculateTallies(remaining, 0);
    var tally = tallies[0]; 
    var voteSorted = tallies[1]; 
    var totalTally = tallies[2];

    // Anyone have a majority? 
    if (tally[voteSorted[0]] > totalTally/2) {
      logToSheet(["The winner:", voteSorted[0]]); 
      ui.alert("The winner is: " + voteSorted[0]);
      break;
    }
    else {
      logToSheet(["No winner yet, highest is " + voteSorted[0] + " with " + tally[voteSorted[0]] + "votes of " + totalTally]);
    }  
    
    // Log the current tally
    for (var currVote in tally) {
      logToSheet(["", "Vote "+ currVote + " " + tally[currVote]]);
    };
    
    eliminateBottom(remaining, 0, round);

    round++;
   
  }
}

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi()
  var settingsSheet = ss.getSheets()[0];
  var voteSheet = ss.getSheets()[1];
  var resultSheet = ss.getSheets()[2];
  var lastColumn = voteSheet.getLastColumn()
  var allCandidates = (settingsSheet.getRange(2, 1, settingsSheet.getLastRow()-1, 1)).getValues();
  var numberOfOptions = (settingsSheet.getRange(1, 4, 1, 1)).getValues();

  var range = voteSheet.getRange(2, 2, voteSheet.getLastRow()-1,lastColumn-1);
  var values = range.getValues();
  var numVoters = values.length;
  var buckets = [];
  
  resultSheet.clear();

  // Initially, all candidates are remaining but we want to keep the allCandidates range and have this array instead.
  var remaining = [];
  for (var currCandidate in allCandidates) {
    remaining.push(allCandidates[currCandidate][0]);
  }
  
  // Function for logging stuff to the result sheet.
  function logToSheet(value) {
    resultSheet.getRange(resultSheet.getLastRow()+1,1,1,value.length).setValues([value]);
  }
  
  function calculateTallies(candidates, level) {
    // Calculate tallies for selected candidates and sort them, also return a total of votes
    
    // Initialize tally with all candidates
    var tally = {}
    for (var currCandidate in candidates) {
      tally[candidates[currCandidate]] = 0;
      
    }    
    
    for (var row in values) { 
      // Only count selected candidates
      Logger.log("values[row][level] = " + values[row][level]);
      if (candidates.indexOf(values[row][level]) > -1 ) {
        tally[values[row][level]]++;
      };
    };       
    // Sort by votes
    voteSorted = Object.keys(tally).sort(function(a,b){return tally[b]-tally[a]})
    Logger.log(voteSorted.toString());
    // Summarize tally
    var totalTally = 0;
    for (currTally in tally) {
      totalTally+= tally[currTally];
    };
    
    return [tally, voteSorted, totalTally]
  }
  
  function returnBottom(candidates, tally) {
    // Return the bottom candidate(s) from a list sorted by votes
    
    var bottom = []
    var currCandidate = candidates.length -1;
    var currVotes = tally[candidates[currCandidate]];
    bottom.push(candidates[currCandidate]);
    while (currCandidate > 0 && currVotes == tally[candidates[currCandidate - 1]] ) {
      currCandidate--;
      bottom.push(candidates[currCandidate]);
    }
    return bottom;
  }
  
  
  function eliminateBottom(candidates, level, round) {
    // Calculate tallies
    var tallies = calculateTallies(candidates, level);
    var tally = tallies[0];
    var voteSorted = tallies[1];
    var totalTally = tallies[2];
    
    var bottom = returnBottom(voteSorted, tally);
    if (bottom.length == 1) {
      
      
      remaining.splice(remaining.indexOf(bottom[0]),1);
      logToSheet([bottom[0] + " was eliminated in round " + round]);

    }
    else {
      if (level < numberOfOptions) {
        eliminateBottom(bottom, level+1, round)
      }
      else {
        logToSheet(["Asking user to choose who to eliminate among " + bottom.toString()]);
        var buildText = "Enter a number to choose who to eliminate:\n";
        for (var currCandidate in bottom) {
          buildText += "[" + currCandidate.toString() +"]"+ bottom[currCandidate] + "\n";
        }
        
        var result = ui.prompt('Input required', buildText, ui.ButtonSet.OK_CANCEL);
        var button = result.getSelectedButton();
        var text = result.getResponseText();
        if (button == ui.Button.OK) {
          // User clicked "OK".
          remaining.splice(remaining.indexOf(bottom[Number(text)]),1);
      
          logToSheet([bottom[Number(text)] + " was manually chosen to be eliminated in round " + round]);

        } else if (button == ui.Button.CANCEL) {
          // User clicked "Cancel".
          logToSheet(["User exited the procedure"]);
          throw new Error("User exited the procedure");
        } else if (button == ui.Button.CLOSE) {
          // User clicked X in the title bar.
          logToSheet(["User exited the procedure"]);
          throw new Error("User exited the procedure");

        }

      }
    }
  }
      
  
  Logger.log(values);
  // Start with the first round and work until done
  var currResults = [];
  
 

  logToSheet(Array.concat(["candidates"], [remaining.toString()]));
  
  // TODO: Loop and check so that all are valid
  //Alert("Invalid candidates");
  // Mark those red.
  var round = 1;
  var done = false;
  // Loop until we are done
  while (!done) { 

    logToSheet(["Round " + round.toString()]);
    
    // Calculate tallies for first votes
    var tallies = calculateTallies(remaining, 0);
    var tally = tallies[0]; 
    var voteSorted = tallies[1]; 
    var totalTally = tallies[2];
    

    // Anyone have a majority? 
    if (tally[voteSorted[0]] > totalTally/2) {
      logToSheet(["The winner:", voteSorted[0]]); 
      ui.alert("The winner is: " + voteSorted[0]);
      break;
    }
    else {
      logToSheet(["No winner yet, highest is " + voteSorted[0] + " with " + tally[voteSorted[0]] + "votes of " + totalTally]);
    }  
    
    // Log the current tally
    for (var currVote in tally) {
      logToSheet(["", "Vote "+ currVote + " " + tally[currVote]]);
    };
    
    eliminateBottom(remaining, 0, round);

    
    round++;
   
  }
}
