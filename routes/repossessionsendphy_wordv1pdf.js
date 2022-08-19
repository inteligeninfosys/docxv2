var express = require('express');
var router = express.Router();
const fs = require('fs');

const cors = require('cors');
var Minio = require("minio");
const docx = require('docx');
const DATE = new Date();

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

    var dd = {
        pageSize: 'A4',
        pageOrientation: 'portrait',
        // [left, top, right, bottom] or [horizontal, vertical] or just a number for equal margins
        pageMargins: [50, 60, 50, 60],
        footer: {
            columns: [
                //{ text: 'Directors: John Murugu (Chairman), Dr. Gideon Muriuki (Group Managing Director & CEO), M. Malonza (Vice Chairman),J. Sitienei, B. Simiyu, P. Githendu, W. Ongoro, R. Kimanthi, W. Mwambia, R. Simani (Mrs), L. Karissa, G. Mburia.\n\n' }
                { text: data.footeroneline }
            ],
            style: 'superMargin'
        },

        content: [
            {
                alignment: 'justify',
                columns: [
                    {
                        text: './coop.jpg',
                        width: 300
                    },
                    {
                        type: 'none',
                        alignment: 'right',
                        fontSize: 9,
                        ol: [
                            'The Co-operative Bank of Kenya Limited',
                            'Co-operative Bank House',
                            'Haile Selassie Avenue',
                            'P.O. Box 48231-00100 GPO, Nairobi',
                            'Tel: (020) 3276100',
                            'Fax: (020) 2227747/2219831'
                        ]
                    },
                ],
                columnGap: 10
            },
            
            'Our Ref: ',
            '\n\nDate Issued: ' + year+'-'+month+'-'+day + '                                                                            Valid up to: ' + letter_data.expirydate,
            '\n\nTo:',
            '\n' + letter_data.auctioneername,
            

            '\nDear Sir/Madam',



            {
                text: '\nRE: REPOSSESSION/COLLECTION ORDER',
                style: 'subheader'
            },

            { text: '\nAccording to our records, the monthly rental of the above Asset finance Agreement is now in arrears. The total amount due is Kes. 92,791.97.', alignment: 'justify' },
            '\n',
            

            {
                text: 'Please approach the above named Debtor on our behalf and collect the total sum of Kes. 92,791.97 plus your own charges or, failing this, you may take this letter as your authority to effect immediate re-possession of the above vehicle without further reference to us. ',
                fontSize: 10, alignment: 'justify'
            },

            { text: '\nUpon successful Repossession, the Motor Vehicle / Asset shall be booked at the nearest storage yard as detailed in the attached list of storage yards in the Bank’s panel.  ', fontSize: 10, alignment: 'justify' },

            { text: '\nAll payments MUST be made to the Co-operative Bank Account of the borrower as indicated above. From our records, we are able to give the following additional information regarding this Agreement, which may assist you in your task: -'},

            {text: '\n'},

            {
                alignment: 'justify',
                fontSize: 9.5,
                table: {
                    widths: [180, '*'],
                    body: [
                        ['Postal address of Debtor', ''+letter_data.postaladdress||'N/A'],
                        ['Telephone', ''+letter_data.telnumber||'N/A'],
                        ['Actual physical address (if known)', ''+letter_data.physicaladdress||'N/A'],
                        ['Employer (where applicable)', ''+letter_data.employer||'N/A'],
                        ['Type of business', ''+letter_data.typeofbusiness||'N/A'],
                        ['Bankers and Branch', ''+letter_data.branchname||'N/A'],
                        ['Purpose for which vehicle is being used', ''+letter_data.purposeofvehicle||'N/A'],
                        ['Guarantor (if any)', ''+letter_data.guarantors||'N/A'],
                        ['Address of Guarantor', ''+letter_data.guarantorsaddress||'N/A'],
                        ['Tracking Information / Report', ''+letter_data.trackingcompany||'N/A'],
                        ['Any other information', 'Arrears relate to part unpaid instalments for IPF loan and/or Asset Loan inclusive of accrued default interest. Please ensure to collect the amount on our behalf either in cash or Bankers Cheque or else repossess the vehicle. Ensure to collect your charges directly from the client.']
                    ]
                }
            },

          
            {text: '\nTerms and Conditions:', fontSize: 11, decoration: 'underline', bold: true},
            {text: '1. These instructions DO NOT give the auctioneer the right to sell the securities / motor vehicle / assets seized from the borrower or guarantor'},
            {text: '2. Repossession fee and all other costs relating to recovery of the motor vehicle(s) will only be paid to the auctioneer who successfully recovers the asset on behalf of the Co-operative Bank.'},
            {text: '3. The successful auctioneer must provide the booking form which details the following:'},
            {text: '  a. Detailed description of the repossessed vehicle'},
            {text: '  b. Storage Yard Booking Sheet from the designated Yard'},
            {text: '4. Repossession fee will be paid in line with the contract terms agreed on between the Bank and yourselves and will be done directly by Co-operative Bank.'},
            {text: '5. These instructions are valid for fourteen (14) calendar days.'},
            {text: '6. For any exceptions to the above, kindly obtain approval beforehand from the undersigned.'},

            { text: '\nYours Faithfully, ' },
            { text: '\n\nAuthorised Signatory, ' },

            { text: '\nThis letter is electronically generated and is valid without a signature ', fontSize: 9, italics: true, bold: true },

            {text: '\nAcceptance by Auctioneer:', fontSize: 11, decoration: 'underline', bold: true},
            {text: '\nI …………………………………………………………………………… on behalf of ………………………………………………………………… confirm that I have read, understood and acceptance with the terms and conditions as stated in the repossession instruction.'},
            {text: '\n\nName: ………………………………………………………………… Signature: …………………………………………… Date: ……………………………………'}

        ],
        styles: {
            header: {
                fontSize: 18,
                bold: true,
                alignment: 'right',
                margin: [0, 190, 0, 80]
            },
            subheader: {
                fontSize: 12,
                bold: true,
                decoration: 'underline'
            },
            superMargin: {
                margin: [20, 0, 40, 0],
                fontSize: 8, alignment: 'center', opacity: 0.5
            },
            quote: {
                italics: true
            },
            small: {
                fontSize: 8
            }
        },
        defaultStyle: {
            fontSize: 10
        }
    }; // end dd

    var options = {
        // ...
    }

    var pdfDoc = printer.createPdfKitDocument(dd, options);
    // ensures response is sent only after pdf is created
    writeStream = fs.createWriteStream(LETTERS_DIR + accnumber_masked + "repossession.pdf");
    pdfDoc.pipe(writeStream);
    pdfDoc.end();
    writeStream.on('finish', async function () {
        // do stuff with the PDF file
        // save to minio
        const filelocation = LETTERS_DIR + accnumber_masked + "repossession.pdf";
        const bucket = 'demandletters';
        const savedfilename = accnumber_masked + '_' + Date.now() + '_' + "repossession.pdf"
        var metaData = {
            'Content-Type': 'text/html',
            'Content-Language': 123,
            'X-Amz-Meta-Testing': 1234,
            'example': 5678
        }
        const objInfo = await minioClient.fPutObject(bucket, savedfilename, filelocation, metaData);

        res.json({
            result: 'success',
            message: LETTERS_DIR + accnumber_masked + DATE + "repossession.pdf",
            filename: accnumber_masked + DATE + "repossession.pdf",
            savedfilename: savedfilename,
            objInfo: objInfo
        })
        deleteFile(filelocation);

        //save to mino end
    });




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
