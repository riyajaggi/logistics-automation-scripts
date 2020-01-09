/**
 * Iterates row by row in the input range and returns an array of objects.
 * Each object contains all the data for a given row, indexed by its normalized column name.
 * @param {Sheet} sheet The sheet object that contains the data to be processed
 * @param {Range} range The exact range of cells where the data is stored
 * @param {number} columnHeadersRowIndex Specifies the row number where the column names are stored.
 *   This argument is optional and it defaults to the row immediately above range;
 * @return {object[]} An array of objects.
 */
function getRowsData(sheet, range, columnHeadersRowIndex) {
    columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
    var numColumns = range.getEndColumn() - range.getColumn() + 1;
    var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
    var headers = headersRange.getValues()[0];
    return getObjects(range.getValues(), normalizeHeaders(headers));
}

/**
 * For every row of data in data, generates an object that contains the data. Names of
 * object fields are defined in keys.
 * @param {object} data JavaScript 2d array
 * @param {object} keys Array of Strings that define the property names for the objects to create
 * @return {object[]} A list of objects.
 */
function getObjects(data, keys) {
    var objects = [];
    for (var i = 0; i < data.length; ++i) {
        var object = {};
        var hasData = false;
        for (var j = 0; j < data[i].length; ++j) {
            var cellData = data[i][j];
            if (isCellEmpty(cellData)) {
                continue;
            }
            object[keys[j]] = cellData;
            hasData = true;
        }
        if (hasData) {
            objects.push(object);
        }
    }
    return objects;
}

/**
 * Returns an array of normalized Strings.
 * @param {string[]} headers Array of strings to normalize
 * @return {string[]} An array of normalized strings.
 */
function normalizeHeaders(headers) {
    var keys = [];
    for (var i = 0; i < headers.length; ++i) {
        var key = normalizeHeader(headers[i]);
        if (key.length > 0) {
            keys.push(key);
        }
    }
    return keys;
}

/**
 * Normalizes a string, by removing all alphanumeric characters and using mixed case
 * to separate words. The output will always start with a lower case letter.
 * This function is designed to produce JavaScript object property names.
 * @param {string} header The header to normalize.
 * @return {string} The normalized header.
 * @example "First Name" -> "firstName"
 * @example "Market Cap (millions) -> "marketCapMillions
 * @example "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
 */
function normalizeHeader(header) {
    var key = '';
    var upperCase = false;
    for (var i = 0; i < header.length; ++i) {
        var letter = header[i];
        if (letter == ' ' && key.length > 0) {
            upperCase = true;
            continue;
        }
        if (!isAlnum(letter)) {
            continue;
        }
        if (key.length == 0 && isDigit(letter)) {
            continue; // first character must be a letter
        }
        if (upperCase) {
            upperCase = false;
            key += letter.toUpperCase();
        } else {
            key += letter.toLowerCase();
        }
    }
    return key;
}

/**
 * Returns true if the cell where cellData was read from is empty.
 * @param {string} cellData Cell data
 * @return {boolean} True if the cell is empty.
 */
function isCellEmpty(cellData) {
    return typeof(cellData) == 'string' && cellData == '';
}

/**
 * Returns true if the character char is alphabetical, false otherwise.
 * @param {string} char The character.
 * @return {boolean} True if the char is a number.
 */
function isAlnum(char) {
    return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

/**
 * Returns true if the character char is a digit, false otherwise.
 * @param {string} char The character.
 * @return {boolean} True if the char is a digit.
 */
function isDigit(char) {
    return char >= '0' && char <= '9';
}

/**
 * Sends emails from spreadsheet rows.
 */
function sendEmails() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = ss.getSheets()[0];
    var dataRange = dataSheet.getRange(2, 1, dataSheet.getMaxRows() - 1, 4);
    
    var templateSheet = ss.getSheets()[1];
    var emailTemplate = templateSheet.getRange('A1').getValue();
    var signText = templateSheet.getRange('A2').getValue();
    
    var certificateID = '';
    var folder = DriveApp.getFolderById('');
    var files = folder.getFiles();
    
    var myLogoURL = PropertiesService.getScriptProperties().getProperty("my");
    var myLogoBlob = DriveApp.getFileById(myLogoURL)
    .getBlob()
    .setName("myLogoBlob");
    
    // Create one JavaScript object per row of data.
    var objects = getRowsData(dataSheet, dataRange);
    
    for (var i = 0; i < objects.length; ++i) {
        var rowData = objects[i];
        
        cell = 2+i;
        //var filesend = files.next();
        
        if (rowData.confirmation != "email sent")
        {
            var emailText = fillInTemplateFromObject(emailTemplate, rowData);
            var emailSubject = 'Certificate';
            var body =  '<p style = "font-size:14px; font-family:arial,helvetica,sans-serif; line-height:25px;">' + emailText + "<img src='cid:myLogo'>" + '</p>' + '<p style = "font-size:17px; font-family:arial,helvetica,sans-serif; line-height:25px;">' + signText  + '</p>'
            //var file = files.next();
            //Logger.log(file);
            MailApp.sendEmail({to: rowData.email, subject: emailSubject, htmlBody: body, attachments: files.next().getAs(MimeType.PDF).setName(rowData.name + rowData.lastName), inlineImages: {myLogo: myLogoBlob}});
            SpreadsheetApp.getActiveSheet().getRange('D' + cell).setValue("email sent");
        }
        else continue;
        
    }
}

/**
 * Replaces markers in a template string with values define in a JavaScript data object.
 * @param {string} template Contains markers, for instance ${"Column name"}
 * @param {object} data values to that will replace markers.
 *   For instance data.columnName will replace marker ${"Column name"}
 * @return {string} A string without markers. If no data is found to replace a marker,
 *   it is simply removed.
 */
function fillInTemplateFromObject(template, data) {
    var email = template;
    // Search for all the variables to be replaced, for instance ${"Column name"}
    var templateVars = template.match(/\$\{\"[^\"]+\"\}/g);
        
        // Replace variables from the template with the actual values from the data object.
        // If no value is available, replace with the empty string.
        for (var i = 0; templateVars && i < templateVars.length; ++i) {
            // normalizeHeader ignores ${"} so we can call it directly here.
            var variableData = data[normalizeHeader(templateVars[i])];
            email = email.replace(templateVars[i], variableData || '');
        }
        
        return email;
    }
