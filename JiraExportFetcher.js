const Excel = require('exceljs');
const request = require('request');
const prompt = require('prompt');
const fs = require('fs');

var schema = {
    properties: {
        username: {
        required: true
        },
        password: {
        hidden: true
        }
    }
};

var srcFile;
var encodedCredentials;
var minColWidth = 6;
var maxColWidth = 56;
var attachColWidth = 34;

// check arguments for source file
if (process.argv.length != 3) {
    console.log('\n\rWrong number of arguments!\n\r\n\rUsage: node JireExportFetcher.js <JiraExportFile>\n\r');
    process.exit(1);
} else {
    srcFile = process.argv[2];
}

function downloadFile(src, dest) { 
    var success = false;
    var file = fs.createWriteStream(dest);
    var filePath = __dirname + '\\' + dest;
    var stream = request({
        method: "GET", 
        "rejectUnauthorized": false, 
        "url": src,
        "headers" : 
        {
        "Content-Type": "application/json",
        "Authorization": "Basic" + ' ' + encodedCredentials
        }
    },function(err,data,body){ 
        //console.log('status code: ' + data.statusCode);
        if (data.statusCode >= 400) { // html error
            console.log('Error (status code: ' + data.statusCode + ') downloading file: ' + src + ' failed!');
            if (data.statusCode = 401) { // Auth failed
                console.log('Authentification failure! Please check credentials.');
                process.exit(1);
            } else if (data.statusCode = 404) { // File not found
                console.log('File not found. Please check online source manually.');
                //process.exit(1);
            }
        } else {
            //TODO: check status code further (200 = OK)
            success = true;
        }
    }).pipe(fs.createWriteStream(filePath));
    stream.on('finish', function () {
        if (success) {
            console.log(dest + ' has been downloaded!');
            //return true;
        } else {
            // something went wrong delete (empty?) file
            fs.unlinkSync(filePath);
        }
    });
} 

prompt.start();

prompt.get(schema, function (err, result) {
    var credentials = result.username + ':' + result.password;
    encodedCredentials = new Buffer(credentials).toString('base64');

    console.log('parsing file: ' + srcFile);

    // read from a file
    var workbook = new Excel.Workbook();
    workbook.csv.readFile(srcFile)
        .then(function(worksheet) {
            // use workbook or worksheet
            worksheet.columns.forEach(column => {
                // remove first entry (1-based index), use as header
                var colHeader = column.values[1];
                column.header = colHeader;
                //TODO: check if shift() is better than skipping the first row all the time!
                //column.values.shift(); 

                // set min coloumn width
                column.width = minColWidth;              
                // iterate over all current cells in this column
                column.eachCell(function(cell, rowNumber) {
                    // find attachments
                    if (colHeader.indexOf('Attachment') == 0) {
                        if (rowNumber == 1) {
                            // abuse skipping of first entry to set width of attachment column (need special treatment)
                            column.width = attachColWidth;
                        } else {
                            if(cell.value != null) {
                                //console.log('Column ' + attachmentKeys[i] + ', Cell ' + rowNumber + ' = ' + cell.value);
                                var splitted = cell.value.split(';');
                                var url = splitted[splitted.length - 1];
                                var fileName = 'files\\' + worksheet.getCell('B' + rowNumber).value + '_' + splitted[splitted.length - 2];
                                // download file TODO: add return value for FileNotFound
                                downloadFile(url, fileName);
                                // replace cell value with local link to file (TODO: if file was successfully downloaded)
                                cell.value = { text: fileName, hyperlink: fileName };
                                cell.font = { color: {argb: 'FF0000FF'} }
                            }
                        }
                    }

                    // format style
                    if (rowNumber % 2 == 0) {
                        // even rownumbers
                        cell.fill = {
                            type: 'pattern',
                            pattern:'solid',
                            fgColor: {argb: 'FFFFE0B3'}
                        };
                    } else {
                        // not even rownumbers
                        if (rowNumber == 1) { // header
                            cell.fill = {
                                type: 'pattern',
                                pattern: 'solid',
                                fgColor: {argb: '80FF9F0C'}
                            };

                            cell.font = {
                                name: 'Calibri',
                                color: {argb: 'FFFFFFFF'},
                                family: 2,
                                size: 12
                            };
                        }
                    }
                    // adjust width according to content length
                    if(rowNumber != 1 && cell.value && cell.value.length > column.width) {
                        column.width = cell.value.length > maxColWidth ? maxColWidth : cell.value.length;
                    }
                });
            })

             var outFileName = srcFile.replace(/\.[^/.]+$/, "") + '_parsed.xlsx';
             workbook.xlsx.writeFile(outFileName)
                 .then(function() {
                     // done
                     console.log('Parsing Jira Export done! Output file: ' + outFileName);
                     console.log('Please wait for all downloads to finish before using it.\r\n');
                 }).catch(function(rejection) {
                    // write promise got reject
                    console.log('\r\nFailed to write output file ' + outFileName + '!');
                    console.log(rejection.toString());
                    console.log('If the file is currently opened in Excel please close it and try again.');
                });
        }).catch(function(rejection) {
            // read promise got reject
            console.log('\r\nCouldn\'t read from file ' + srcFile + '!');
            console.log(rejection.toString());
        });
});

