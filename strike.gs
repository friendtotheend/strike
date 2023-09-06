

function collectReport(color, finish_type, urgency, nesting_level) {



let grandTotalReport = {}
DocumentApp.getActiveDocument().getBody().getListItems().forEach(item => {
  
  //making a string to keep track of a bullet point's attributes
  var itemAttributes= item.getAttributes()
  var backgroundColor = item.editAsText().getBackgroundColor(0) //this cannot be taken from item attributes because value is null for last bullet point bc newline isn't highlighted. So, now just checking first character highlight
  var nestingLevel = itemAttributes.NESTING_LEVEL
  var isUrgent = (item.editAsText().isBold(0)) ? "urgent" : "not_urgent"
  var isFinished = (item.editAsText().isStrikethrough(0)) ? "finished" : "unfinished"


  var grandTotalAssembledCharacteristics = backgroundColor + '|' + isUrgent //we don't care about nesting level or urgency in the grand total

  //either 'all' is specified or the characteristic has to match exactly as a standalone string or a string in an array.
  var colorMatch = (color == 'all' || backgroundColor == color || (Array.isArray(color) && color.includes(backgroundColor))) ? true : false
  var finishMatch = ((finish_type == 'all') || (finish_type == 'finished' && isFinished == "finished") || (finish_type='unfinished' && isFinished == "unfinished")) ? true : false
  var urgencyMatch = (urgency == 'all' || (urgency == "urgent" && isUrgent == "urgent") || (urgency == "not_urgent" && isUrgent == "not_urgent")) ? true : false
  var nestingLevelMatch = (nesting_level == 'all' || nestingLevel == nesting_level  || (Array.isArray(nesting_level) && nesting_level.includes(nestingLevel))) ? true : false

  if (colorMatch && finishMatch && urgencyMatch && nestingLevelMatch && item.getText().length > 0) //if an item matches the query parameters
  {
    
    
    //increment the grand total.
    grandTotalReport[grandTotalAssembledCharacteristics] = (grandTotalReport[grandTotalAssembledCharacteristics] ? grandTotalReport[grandTotalAssembledCharacteristics] : 0) + 1
  }
  
  })

 return(grandTotalReport)
}

//getting rid of all existing tables in document
DocumentApp.getActiveDocument().getBody().getTables().forEach(table => {DocumentApp.getActiveDocument().getBody().removeChild(table)})

//getting rid of the extra space added each time the grand total table is constructed
let body = DocumentApp.getActiveDocument().getBody()

//code I found to remove the extra empty paragraph caused by appending table.

var paragraphs = DocumentApp.getActiveDocument().getBody().getParagraphs();
for (var i=paragraphs.length-1; i>=0; i--){
      var line = paragraphs[i];   
      if (line.getText().trim() ) {
        break;
      }
      else {
        line.clear();
        try {line.merge()}
        catch {}        
      }
    }


////constructing the grand total table/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


let cells = []
let query = collectReport("all", "unfinished", "all", "all")


//populating grandTotalCells
for (const row in query){cells.push(row.split("|").concat([query[row].toFixed(0)]))}

//sorting grandTotalCells
cells = cells.sort(function(a,b){return a[0].localeCompare(b[0])})

//making the table header
if (cells.length > 0) {cells.unshift(["color", "urgency", "unfinished count"])}
appendFormatTable(cells,0)



//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

//sets the text color to be whatever hex code is in the first column of a tally table. Removes backgroundcolor.. Color coding!
function foregroundSetter(table, numRows)
{
  for (i=1; i<numRows; i++)
  {
    let row = table.getRow(i)
    row.setForegroundColor(row.getCell(0).getText()).setBackgroundColor(null)
  }
  
}

//appends a new table, sets header format, sets colors of rows based on listed color code.
function appendFormatTable(cells, index)
{
  DocumentApp.getActiveDocument().getBody().appendTable(cells)
  let table = DocumentApp.getActiveDocument().getBody().getTables()[index]
  let numRows = table.getNumRows()
  //header row should be black text and bolded.
  table.getRow(0).setForegroundColor("#000000").setBold(true).setBackgroundColor(null)
  foregroundSetter(table,numRows)
}

