var express = require('express');
var router = express.Router();
const fs = require('fs');
var numeral = require('numeral');
const cors = require('cors');
var Minio = require("minio");
const docx = require('docx');


var minioClient = new Minio.Client({
    endPoint: process.env.MINIO_ENDPOINT || '127.0.0.1',
    port: process.env.MINIO_PORT ? parseInt(process.env.MINIO_PORT, 10) : 9005,
    useSSL: false,
    accessKey: process.env.ACCESSKEY || 'AKIAIOSFODNN7EXAMPLE',
    secretKey: process.env.SECRETKEY || 'wJalrXUtnFEMIK7MDENGbPxRfiCYEXAMPLEKEY'
});

var data = require('./data.js');

const LETTERS_DIR = data.filePath;
const d_t = new Date();
 
let year = d_t.getFullYear();
let month = ("0" + (d_t.getMonth() + 1)).slice(-2);
let day = ("0" + d_t.getDate()).slice(-2);


const { Document, Footer, ImageRun, Packer, Paragraph, Table, TableCell, TableRow, AlignmentType,Column,
    BorderStyle,WidthType,TextRun  } = docx;


router.use(express.urlencoded({ extended: true }));
router.use(express.json());

router.use(cors());


router.get('/', function (req, res) {
    res.json({ message: 'Reposession send physically letter is ready!' });
});


router.post('/download', function (req, res) {
    const letter_data = req.body;

    const rawaccnumber = letter_data.accnumber;

    const first4 = rawaccnumber.substring(0, 9);
    const accnumber_masked = first4 + 'xxxxx';

    const doc = new Document({
        
        sections: [
            {
                footers: {
                    default: new Footer({
                        children: [new Paragraph({
                            text: data.footerfirst + data.footersecond,
                            alignment: AlignmentType.CENTER,
                        })],
                    }),
                },
                children: [
                    new Table({
                        width: {
                            size: 9035,
                            type: WidthType.DXA,
                        },
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 0
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 0
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 0
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 0
                            },
                        },
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({
                                        rowSpan: 7,
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [
                                            new Paragraph({
                                                children: [
                                                    new ImageRun({
                                                        data: fs.readFileSync("./routes/coop.jpg"),
                                                        transformation: {
                                                            width: 300,
                                                            height: 60,
                                                        },
                                                    }),
                                                ],
                                            }),
                                        ],
                                    }),
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph('The Co-operative Bank of Kenya Limited')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph('Co-operative Bank House')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph('Haile Selassie Avenue')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph('P.O. Box 48231-00100 GPO, Nairobi')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph('Tel: (020) 3276100')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph('Fax: (020) 2227747/2219831')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph('www.co-opbank.co.ke')],
                                    }),
                                ]
                            })
                        ]
                    }),
                    new Paragraph(''),
                    new Paragraph(''),
                    new Table({
                        width: {
                            size: 9035,
                            type: WidthType.DXA,
                        },
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 0
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 0
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 0
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 0
                            },
                        },
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [
                                            new Paragraph('Date: '+ year + "-" + month + "-" + day),
                                        ],
                                    }),
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph('Valid up to: '+ (letter_data.expirydate).toUpperCase())],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph('To')],
                                    }),
                                    new TableCell({
                                        children: [new Paragraph(': ' + letter_data.auctioneername)],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph('Asset Finance Agreement No.')],
                                    }),
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph(': ' + letter_data.assetfaggnum)],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph('Hirer’s Name')],
                                    }),
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph(': ' + letter_data.custname)],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph('Unit Financed')],
                                    }),
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph(': ' + letter_data.vehiclemake + ' & ' + letter_data.vehiclemodel)],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        width: {
                                            size: 4035,
                                            type: WidthType.DXA,
                                        },
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph('Registration No')],
                                    }),
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph(': ' + letter_data.vehicleregno)],
                                    }),
                                ]
                            })
                        ]
                    }),// end table
                    new Paragraph(''),
                    new Paragraph({
                        children: [new TextRun("According to our records, the monthly rental of the above Asset finance Agreement is now in arrears. The total amount due is Kes. " + numeral(Math.abs(letter_data.totalamount)).format('0,0.00'))],
                    }),
                    new Paragraph(''),
                    new Paragraph({
                        children: [new TextRun("Please approach the above named Hirer on our behalf and collect the total sum of Kes. " + numeral(Math.abs(letter_data.totalamount)).format('0,0.00') + " plus your own charges, failing which you may take this letter as your authority to effect immediate re-possession of the above/equipment without further reference to us. HIRER MUST MAKE PAYMENT VIDE CASH OR BY BANKER’S CHEQUE AS PERSONAL CHEQUE(S) WILL NOT BE ACCEPTED. From our records, we are able to give the following additional information regarding this Agreement, which may assist you in your task of locating the hirer and/or the motor vehicle/equipment: -")],
                    }),
                    new Paragraph(''),
                    new Table({
                        width: {
                            size: 9035,
                            type: WidthType.DXA,
                        },
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 0
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 0
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 0
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 0
                            },
                        },
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [
                                            new Paragraph('Postal Address'),
                                        ],
                                    }),
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph(': '+ letter_data.postaladdress || 'N/A')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph('Telephone')],
                                    }),
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph(': '+ letter_data.celnumber || 'N/A')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph('Physical Address/Location')],
                                    }),
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph(': '+ letter_data.place || 'N/A')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph('Type of Business')],
                                    }),
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph(': '+ letter_data.typeofbusiness || 'N/A')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph('Bankers and Branch')],
                                    }),
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph(': '+ letter_data.branchname || 'N/A')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph('Purpose of Vehicle')],
                                    }),
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph(': '+ letter_data.purposeofvehicle || 'N/A')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph('Guarantors')],
                                    }),
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph(': '+ letter_data.guarantors || 'N/A')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph('Guarantors Address')],
                                    }),
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph(': '+ letter_data.guarantorsaddress || 'N/A')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph('Chassis No.')],
                                    }),
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph(': '+ letter_data.chassisnumber || 'N/A')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph('Engine No.')],
                                    }),
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph(': '+ letter_data.engineno || 'N/A')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        width: {
                                            size: 4035,
                                            type: WidthType.DXA,
                                        },
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph('Any other information')],
                                    }),
                                    new TableCell({
                                        borders: {
                                            top: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            right: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                            bottom: {
                                                style: BorderStyle.DASH_DOT_STROKED,
                                                size: 1,
                                                color: "ffffff",
                                            },
                                        },
                                        children: [new Paragraph(letter_data.anyotherinfo)],
                                    }),
                                ]
                            }),
                        ]
                    }),// end table
                    new Paragraph(''),
                    new Paragraph('Vehicle tracked by: '+ letter_data.trackingcompany),
                    new Paragraph(''),
                    new Paragraph('Yours Faithfully,'),
                    new Paragraph(''),
                    new Paragraph(''),
                    new Paragraph(''),
                    new Paragraph(''),
                    new Paragraph('AUTHORISED SIGNATORY,'),
                    new Paragraph(''),
                    new Paragraph('Cc'),
                    new Paragraph(letter_data.custname),
                    new Paragraph(letter_data.postaladdress)
                ],
                
            },
        ],
    });

  
    Packer.toBuffer(doc).then((buffer) => {
        fs.writeFileSync(LETTERS_DIR + accnumber_masked +  "repossession.docx", buffer);
        // save to minio
        const filelocation = LETTERS_DIR + accnumber_masked +  "repossession.docx";
        const bucket = 'demandletters';
        const savedfilename = accnumber_masked + '_' + "repossession.docx"
        var metaData = {
            'Content-Type': 'text/html',
            'Content-Language': 123,
            'X-Amz-Meta-Testing': 1234,
            'example': 5678
        }
        minioClient.fPutObject(bucket, savedfilename, filelocation, metaData, function (error, objInfo) {
            if (error) {
                console.log(error);
                res.status(500).json({
                    success: false,
                    error: error.message
                })
            }
            res.json({
                result: 'success',
                message: LETTERS_DIR + accnumber_masked + "repossession.docx",
                filename: accnumber_masked + "repossession.docx",
                savedfilename: savedfilename,
                objInfo: objInfo
            })
        });
        //save to mino end
    })




}); // end post
function deleteFile(req) {
    fs.unlink(req, (err) => {
        if (err) {
            console.error(err)
            return
        }
        //file removed
    })
}

module.exports = router;
