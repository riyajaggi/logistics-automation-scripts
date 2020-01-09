
function addToListserv() {
    const id_col = 2
    const listserv_col = 4
    const member_col = 5
    
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    
    var doc2 = DocumentApp.openById('');
    var body2 = doc2.getBody();
    
    var sheet_url = '';
    var sss = SpreadsheetApp.openByUrl(sheet_url);
    
    var sheet = sss.getSheets()[0];
    sheet.sort(id_col, true);
    
    body.clear();
    body2.clear();
    
    var last_row = sheet.getLastRow();
    var data_range = "R1C1" + ":R" + last_row + "C6";
    //Logger.log(data_range);
    
    var range = sheet.getRange(data_range);
    var sign_ins = range.getValues();
    var count = 0;
    var count2 = 0;
    
    for (i in sign_ins) {
        if(sign_ins[i][listserv_col] == "yes") {
            count++;
            body.appendParagraph(sign_ins[i][id_col] + "@gmail.com");
        }
        
        if(sign_ins[i][member_col] == "yes") {
            count2++;
            body2.appendParagraph(sign_ins[i][id_col] + "@gmail.com");
        }
    }
    
    body.appendHorizontalRule();
    body.appendParagraph("");
    body.appendParagraph("Total sign-ups: " + count)
    body.appendParagraph("END OF AUTOMATION");
    
    body2.appendHorizontalRule();
    body2.appendParagraph("");
    body2.appendParagraph("Total interest: " + count2)
    body2.appendParagraph("END OF AUTOMATION");
}


