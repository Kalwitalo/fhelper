var express = require('express');
var app = express();
var bodyParser = require('body-parser');
var multer = require('multer');
var xlstojson = require("xls-to-json-lc");
var xlsxtojson = require("xlsx-to-json-lc");


// Define font files
var fonts = {
    Roboto: {
        normal: 'fonts/Roboto-Regular.ttf',
        bold: 'fonts/Roboto-Medium.ttf',
        italics: 'fonts/Roboto-Italic.ttf',
        bolditalics: 'fonts/Roboto-MediumItalic.ttf'
    }
};
var PdfPrinter = require('pdfmake');
var printer = new PdfPrinter(fonts);
var fs = require('fs');

app.use(bodyParser.json());

var storage = multer.diskStorage({ //multers disk storage settings
    destination: function (req, file, cb) {
        cb(null, './uploads/')
    },
    filename: function (req, file, cb) {
        var datetimestamp = Date.now();
        cb(null, file.fieldname + '-' + datetimestamp + '.' + file.originalname.split('.')[file.originalname.split('.').length - 1])
    }
});

var upload = multer({ //multer settings
    storage: storage,
    fileFilter: function (req, file, callback) { //file filter
        if (['xls', 'xlsx'].indexOf(file.originalname.split('.')[file.originalname.split('.').length - 1]) === -1) {
            return callback(new Error('Wrong extension type'));
        }
        callback(null, true);
    }
}).single('file');

/** API path that will upload the files */
app.post('/upload', function (req, res) {
    var exceltojson;
    upload(req, res, function (err) {
        if (err) {
            res.json({error_code: 1, err_desc: err});
            return;
        }
        /** Multer gives us file info in req.file object */
        if (!req.file) {
            res.json({error_code: 1, err_desc: "No file passed"});
            return;
        }
        /** Check the extension of the incoming file and
         *  use the appropriate module
         */
        if (req.file.originalname.split('.')[req.file.originalname.split('.').length - 1] === 'xlsx') {
            exceltojson = xlsxtojson;
        } else {
            exceltojson = xlstojson;
        }
        console.log(req.file.path);
        try {
            exceltojson({
                input: req.file.path,
                output: null, //since we don't need output.json
                lowerCaseHeaders: true
            }, function (err, result) {
                if (err) {
                    return res.json({error_code: 1, err_desc: err, data: null});
                }

                var body = [];
                console.log(result);

                result.forEach(function (row) {
                    var dataRow = [];

                    dataRow.push(row.codeid);

                    body.push(dataRow);
                });

                var docDefinition = {
                    pageSize: {
                        width: 145,
                        height: 88
                    },
                    pageMargins: [ 0, 0, 0, 0 ],

                    content: [
                        {
                            layout: 'noBorders', // optional
                            table: {
                                widths: ['100%'],
                                body: body

                            },
                            fontSize: 20,
                            alignment: 'center',
                            bold: true,
                            margin: [0, 0, 0, 0]
                        }
                    ]
                };


                var temp123;
                var pdfDoc = printer.createPdfKitDocument(docDefinition);
                pdfDoc.pipe(temp123 = fs.createWriteStream('document.pdf'));
                pdfDoc.end();


                temp123.on('finish', async function () {
                    // do send PDF file
                    res.download('document.pdf');
                });
            });
        } catch (e) {
            res.json({error_code: 1, err_desc: "Corupted excel file"});
        }
    })

});

app.get('/downloadFile/', (req, res) => {
    res.download('document.pdf');
})

app.get('/', function (req, res) {
    res.sendFile(__dirname + "/index.html");
});
const port = process.env.PORT || 3000;

app.listen(port, function () {
    console.log('running on 3000...');
});
