const ExcelJS = require('exceljs/dist/es5');

exports.showUploadPage = (req, res, next) => {  
    res.render("files");
},
exports.uploadExcel = (req, res, next) => {
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
            let message = "¡Archivo cargado Exitosamente!";
            let status = true;
            let workbook = new ExcelJS.Workbook(); 
            workbook.xlsx.readFile(`uploads/${sugoFile.name}`)
            .then(() => {
                let worksheet = workbook.getWorksheet('sugo_Alta_Colaboradores');

                let company = worksheet.getCell('D8').value;
                let RUC = worksheet.getCell('D9').value;

                if (!company)
                {
                    console.log("Error el valor para Compañia no puede estar vacio");
                    message = "Error el valor para Compañia no puede estar vacio";
                    status = false;
                }

                if(status === true)
                {
                    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                        if(row.values[1].includes("@"))
                            console.log("Email" + " =" + JSON.stringify(row.values[1]));
                    });                

                    console.log("company " + " = " + JSON.stringify(company));
                    console.log("RUC " + " = " + JSON.stringify(RUC));
                }
            });

            res.json({status: status, message: message});
        }        
    });
}