/*
 
 MUST DO BEFORE RUNNING SCRIPT
 1. Change event_col to the correct event in tracking sheet.
 2. Change sheet_url to the url of the correct sign-in sheet.
 3. Set the last_row for tracking sheet.
 4. Change formula for cell using correct column in tracking sheet.
 
 */

function addToTracker() {
    const netid_col = 5
    const event_col = 73
    const event_alpha = 'BU'
    const attend_col = 3
    var last_row = 175;
    var event_last_row = 34;
    
    
    var ss = SpreadsheetApp.getActive();
    var trackingsheet = ss.getSheets()[0];
    trackingsheet.sort(netid_col, true);
    
    var sheet_url = PropertiesService.getScriptProperties().getProperty("sheeturl");
    var sss = SpreadsheetApp.openByUrl(sheet_url);
    var eventsheet = sss.getSheets()[0];
    eventsheet.sort(attend_col, true);
    
    var trackingrange = trackingsheet.getRange("R5C1" + ":R" + last_row + "C" + event_col);
    var trackingvalues = trackingrange.getValues();
    //var trackobjects = getRowsData(trackingsheet, trackingrange);
    
    var eventrange = eventsheet.getRange("R4C1" + ":R" + event_last_row + "C3");
    var eventvalues = eventrange.getValues();
    //var eventobjects = getRowsData(eventsheet, eventrange);
    
    
    var count = 0;
    var i = 0;
    
    for (i =0; i < eventvalues.length; i++) {
        
        var left = 0;
        var right = last_row - 1;
        
        var event = eventvalues[i][attend_col - 1];
        
        while (left <= right) {
            var mid = left + Math.floor((right - left) / 2);
            if (mid >= trackingvalues.length)
                break;
            
            var track = trackingvalues[mid][4];
            
            if (track == event) {
                var cell = event_alpha + (mid + 5);
                ss.getRange(cell).setValue('y');
                Logger.log(event);
                count++;
                right = -1;
                
                break;
            }
            
            else if (track < event) {
                left = mid + 1;
            }
            else {
                right = mid - 1;
            }
            
        }
        
        
        
    }
    Logger.log("Number of attendees: ");
    Logger.log(count);
    Logger.log("END OF AUTOMATION");
    
}


