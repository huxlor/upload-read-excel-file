require('core-js/modules/es.promise');
require('core-js/modules/es.string.includes');
require('core-js/modules/es.object.assign');
require('core-js/modules/es.object.keys');
require('core-js/modules/es.symbol');
require('core-js/modules/es.symbol.async-iterator');
require('regenerator-runtime/runtime');


const express = require('express');
const fileUpload = require('express-fileupload');
const ExcelJS = require('exceljs/dist/es5');
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

            let workbook = new ExcelJS.Workbook(); 
            workbook.xlsx.readFile(`uploads/${sugoFile.name}`)
            .then(() => {
                let worksheet = workbook.getWorksheet('sugo_Alta_Colaboradores');

                worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                    console.log("Email" + " =" + JSON.stringify(row.values[1]));
                });

                let company = worksheet.getCell('D8').value;
                let RUC = worksheet.getCell('D9').value;

                console.log("company " + " = " + JSON.stringify(company));
                console.log("RUC " + " = " + JSON.stringify(RUC));
            });
        }
          
    
        
    });

});

module.exports = app;