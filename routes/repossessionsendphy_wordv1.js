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


const { Document, Footer, ImageRun, Packer, Paragraph, Table, TableCell, TableRow, AlignmentType, Column,
    BorderStyle, WidthType, TextRun } = docx;


router.use(express.urlencoded({ extended: true }));
router.use(express.json());

router.use(cors());


router.get('/', async function (req, res) {
    res.json({ message: 'Reposession send physically letter is ready!' });
});


router.post('/download', async function (req, res) {
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
                    new Paragraph('Branch Reference number'),
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
                                            new Paragraph('Date issued: ' + year + "-" + month + "-" + day),
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
                                        children: [new Paragraph('Valid up to: ' + (letter_data.expirydate).toUpperCase())],
                                    }),
                                ]
                            }),
                        ]
                    }),// end table
                    new Paragraph(''),
                    new Paragraph('To:'),
                    new Paragraph(''),
                    new Paragraph(letter_data.auctioneername),
                    new Paragraph(letter_data.auctioneername),
                    new Paragraph(''),
                    new Paragraph('Dear Sir/Madam'),
                    new Paragraph(''),
                    new Paragraph({
                        children: [new TextRun({ text: "RE: REPOSSESSION/COLLECTION ORDER", bold: true, underline: true })]
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
                                            new Paragraph('Account No. '),
                                        ],
                                    }),
                                    new TableCell({

                                        children: [new Paragraph(letter_data.accnumber)],
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
                                        children: [
                                            new Paragraph('Debtor. '),
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
                                        children: [new Paragraph(letter_data.custname)],
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
                                        children: [
                                            new Paragraph('Unit '),
                                        ],
                                    }),
                                    new TableCell({

                                        children: [new Paragraph(letter_data.vehiclemake)],
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
                                        children: [
                                            new Paragraph('Reg No. '),
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
                                        children: [new Paragraph(letter_data.vehicleregno)],
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
                                        children: [
                                            new Paragraph('Chassis No. '),
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
                                        children: [new Paragraph(letter_data.chassisnumber)],
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
                                        children: [
                                            new Paragraph('Engine No. '),
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
                                        children: [new Paragraph(letter_data.engineno)],
                                    }),
                                ]
                            }),
                        ]
                    }),// end table

                    new Paragraph(''),
                    new Paragraph({
                        children: [new TextRun("According to our records, the monthly rental of the above Asset finance Agreement is now in arrears. The total amount due is Kes. " + numeral(Math.abs(letter_data.totalamount)).format('0,0.00'))],
                    }),
                    new Paragraph(''),

                    new Paragraph({
                        children: [new TextRun("Please approach the above named Debtor on our behalf and collect the total sum of Kes. " + numeral(Math.abs(letter_data.totalamount)).format('0,0.00') + " plus your own charges or, failing this, you may take this letter as your authority to effect immediate re-possession of the above vehicle without further reference to us. ")],
                    }),
                    new Paragraph(''),
                    new Paragraph({
                        children: [new TextRun("Upon successful Repossession, the Motor Vehicle / Asset shall be booked at the nearest storage yard as detailed in the attached list of storage yards in the Bank’s panel. ")]
                    }),
                    new Paragraph(''),
                    new Paragraph({
                        children: [new TextRun("All payments MUST be made to the Co-operative Bank Account of the borrower as indicated above.")]
                    }),
                    new Paragraph({
                        children: [new TextRun("From our records, we are able to give the following additional information regarding this Agreement, which may assist you in your task: -")]
                    }),
                    new Paragraph(''),
                    new Table({
                        width: {
                            size: 9035,
                            type: WidthType.DXA,
                        },

                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({

                                        children: [
                                            new Paragraph('Postal address of Debtor'),
                                        ],
                                    }),
                                    new TableCell({

                                        children: [new Paragraph(letter_data.postaladdress || 'N/A')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({

                                        children: [new Paragraph('Telephone')],
                                    }),
                                    new TableCell({

                                        children: [new Paragraph(letter_data.celnumber || 'N/A')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({

                                        children: [new Paragraph('Actual physical address (if known) ')],
                                    }),
                                    new TableCell({

                                        children: [new Paragraph(letter_data.place || 'N/A')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({

                                        children: [new Paragraph('Employer (where applicable)')],
                                    }),
                                    new TableCell({

                                        children: [new Paragraph(letter_data.employer || 'N/A')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({

                                        children: [new Paragraph('Type of business ')],
                                    }),
                                    new TableCell({

                                        children: [new Paragraph(letter_data.typeofbusiness || 'N/A')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({

                                        children: [new Paragraph('Bankers and Branch')],
                                    }),
                                    new TableCell({

                                        children: [new Paragraph(letter_data.branchname || 'N/A')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({

                                        children: [new Paragraph('Purpose for which vehicle is being used ')],
                                    }),
                                    new TableCell({

                                        children: [new Paragraph(letter_data.purposeofvehicle || 'N/A')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({

                                        children: [new Paragraph('Guarantor (if any)')],
                                    }),
                                    new TableCell({

                                        children: [new Paragraph(letter_data.guarantors || 'N/A')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({

                                        children: [new Paragraph('Address of Guarantor')],
                                    }),
                                    new TableCell({

                                        children: [new Paragraph(letter_data.guarantorsaddress || 'N/A')],
                                    }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({

                                        children: [new Paragraph('Tracking Information / Report')],
                                    }),
                                    new TableCell({

                                        children: [new Paragraph(letter_data.trackinginformation || 'N/A')],
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
                                        children: [new Paragraph('Any other information')],
                                    }),
                                    new TableCell({

                                        children: [new Paragraph(letter_data.anyotherinfo)],
                                    }),
                                ]
                            }),
                        ]
                    }),// end table
                    new Paragraph(''),
                    new Paragraph({
                        children: [new TextRun({ text: "Terms and Conditions:", bold: true, underline: true })]
                    }),
                    new Paragraph(''),
                    new Paragraph('1. These instructions DO NOT give the auctioneer the right to sell the securities / motor vehicle / assets seized from the borrower or guarantor.'),

                    new Paragraph('2. Repossession fee and all other costs relating to recovery of the motor vehicle(s) will only be paid to the auctioneer who successfully recovers the asset on behalf of the Co-operative Bank.'),

                    new Paragraph('3. The successful auctioneer must provide the booking form which details the following:'),
                    new Paragraph('     a. Detailed description of the repossessed vehicle'),
                    new Paragraph('     b. Storage Yard Booking Sheet from the designated Yard'),
                    new Paragraph('4. Repossession fee will be paid in line with the contract terms agreed on between the Bank and yourselves and will be done directly by Co-operative Bank.'),
                    new Paragraph('5. These instructions are valid for fourteen (14) calendar days.'),
                    new Paragraph('6. For any exceptions to the above, kindly obtain approval beforehand from the undersigned. '),
                    new Paragraph(''),
                    new Paragraph(''),
                    new Paragraph('Yours Faithfully,'),
                    new Paragraph(''),
                    new Paragraph(''),
                    new Paragraph(''),
                    new Paragraph(''),
                    new Paragraph('Authorised Signatory                                                                                     Authorised Signatory'),
                    new Paragraph('Signature Number                                                                                         Signature Number'),
                    new Paragraph(''),
                    new Paragraph({
                        children: [new TextRun({ text: "Acceptance by Auctioneer:", bold: true, underline: true })]
                    }),
                    new Paragraph(''),
                    new Paragraph('I ………………………………………………… on behalf of ………………………………… confirm that I have read, understood and acceptance with the terms and conditions as stated in the repossession instruction.'),
                    new Paragraph(''),
                    new Paragraph(''),
                    new Paragraph('Name: ……………………………………… Signature: …………………………… Date: ……………………')
                ],

            },
        ],
    });

    try {
        const buffer = await Packer.toBuffer(doc);
        fs.writeFileSync(LETTERS_DIR + accnumber_masked + "repossession.docx", buffer);
        // save to minio
        const filelocation = LETTERS_DIR + accnumber_masked + "repossession.docx";
        const bucket = 'demandletters';
        const savedfilename = accnumber_masked + '_' + "repossession.docx"
        var metaData = {
            'Content-Type': 'text/html',
            'Content-Language': 123,
            'X-Amz-Meta-Testing': 1234,
            'example': 5678
        }
        const objInfo = await minioClient.fPutObject(bucket, savedfilename, filelocation, metaData);
        res.json({
            result: 'success',
            message: LETTERS_DIR + accnumber_masked + "repossession.docx",
            filename: accnumber_masked + "repossession.docx",
            savedfilename: savedfilename,
            objInfo: objInfo
        })
        //save to mino end
    } catch (error) {
        console.log(error);
        res.status(500).json({
            success: false,
            error: error.message
        })
    }








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
