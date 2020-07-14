const express = require('express');
const fileUpload = require('express-fileupload');
// const ExcelJS = require('exceljs/dist/es5');
const app = express();

// default options
app.use(fileUpload());

app.post('/upload', function(req, res) {

    if (!req.files || Object.keys(req.files).length === 0) {
        return res.status(400).json({
            status: false,
            err: {
                message: 'No files were uploaded.'
            }
        });
    }

    let sugoFile = req.files.sugoFile;
    let sugoFileCut = sugoFile.name.split('.');
    let extension = sugoFileCut[sugoFileCut.length -1];
    let validExtensions = ['xlsx', 'xlsm', 'xlsb', 'xltx', 'xls'];

    if(validExtensions.indexOf(extension) < 0) {
        return res.status(400).json({
            status: false,
            err: {
                message: 'Valid Extensions: ' + validExtensions.join(', '),
                extension
            }
        });
    }

    sugoFile.mv(`uploads/${sugoFile.name}`, (err) => {
        if (err) {
            return res.status(500).json({
                status: false,
                err
              });
        } else {
            res.json({status: true, message: 'File uploaded!'});

            // let workbook = new Excel.Workbook(); 
            // workbook.xlsx.readFile(sugoFile)
            // .then(() => {
            //     let worksheet = workbook.getWorksheet(sheet);
            //     worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
            //     console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
            //     });
            // });
        }
          
    
        
    });

});

module.exports = app;