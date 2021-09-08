let timeZone = Session.getScriptTimeZone()

function getTimeNavigation() {
  
  let ass = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  let sheetName = ass.getName()
  
  if ((sheetName == 'ЖВ Ярославль') || (sheetName == 'ЖВ Самара')) {

    let activeRange = ass.getActiveRange()
    let currentDate = Utilities.formatDate(new Date(), timeZone, 'MM/dd/yyyy')
    let date = isDate(ass.getRange(activeRange.getRowIndex(), 2).getValue())
    let pointA = ass.getRange(activeRange.getRowIndex(), 5)
    let pointB = ass.getRange(activeRange.getRowIndex(), 6)
    
    if ((currentDate === date) && !pointA.isBlank() && !pointB.isBlank()) {
      
      let city = ass.getRange('B1').getValue()
      
      let map = Maps.newDirectionFinder()

      map.setOrigin(pointA.getValue() + ', ' + city)
      map.setDestination(pointB.getValue() + ', ' + city)
      
      let directions = map.getDirections()
      let leg = directions['routes'][0]['legs'][0]
      let duration =  leg['duration']['value']
      let distance =  leg['distance']['value']

      ass.getRange(activeRange.getRowIndex(), 11).setValue(millisToMinutesAndSeconds((duration / 60) * 1000))
      ass.getRange(activeRange.getRowIndex(), 17).setValue(distance)
    }
  }
  
}

function isValidDate(d) {
  if ( Object.prototype.toString.call(d) !== "[object Date]" ) return false
  return !isNaN(d.getTime())
}

function isDate(sDate) {
  if (isValidDate(sDate)) {
    sDate = Utilities.formatDate(new Date(sDate), timeZone, "MM/dd/yyyy")
  }
  return sDate;
}

function millisToMinutesAndSeconds(millis) {
  var minutes = Math.floor(millis / 60000);
  var seconds = ((millis % 60000) / 1000).toFixed(0);
  return minutes + ":" + (seconds < 10 ? '0' : '') + seconds;
}

function getLastRow() {
    let lastRow = SpreadsheetApp.getActiveSheet().getLastRow()
    SpreadsheetApp.getActiveSheet().getRange(lastRow + 1, 1).activate()
}

function onEdit(e) {
  Logger.log(e.source)
}
