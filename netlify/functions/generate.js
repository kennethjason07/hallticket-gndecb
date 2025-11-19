const Busboy = require('busboy');
const xlsx = require('xlsx');
const PDFDocument = require('pdfkit');
const archiver = require('archiver');
const fs = require('fs');
const path = require('path');

// Base64 encoded logo - embedded for serverless reliability
const LOGO_BASE64 = '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';

// Helper function to draw a single ticket
function drawTicket(doc, row, yStart, deptName, examName, logoData, semester) {
    const width = 595.28;

    const boxTop = yStart - 40;
    const boxBottom = yStart + 240;
    const boxHeight = boxBottom - boxTop;
    const boxLeft = 30;
    const boxRight = width - 30;

    doc.lineWidth(1);
    doc.rect(boxLeft, boxTop, boxRight - boxLeft, boxHeight).stroke();

    // College Logo - using base64 data for serverless reliability
    if (logoData) {
        try {
            const logoBuffer = Buffer.from(logoData, 'base64');
            doc.image(logoBuffer, boxLeft + 10, yStart - 10, { width: 50, height: 50 });
        } catch (e) {
            console.log("Error loading logo:", e);
        }
    }

    doc.font('Helvetica-Bold').fontSize(13);
    doc.text("GURU NANAK DEV ENGINEERING COLLEGE, BIDAR", 0, yStart, { align: 'center', width: width });

    doc.font('Helvetica-Bold').fontSize(10);
    doc.text(deptName.toUpperCase(), 0, yStart + 20, { align: 'center', width: width });

    doc.font('Helvetica-Bold').fontSize(11);
    doc.text(`ADMISSION TICKET FOR ${examName.toUpperCase()}`, 0, yStart + 40, { align: 'center', width: width });

    doc.lineWidth(0.5);
    doc.moveTo(40, yStart + 50).lineTo(width - 40, yStart + 50).stroke();
    doc.lineWidth(1);

    doc.font('Helvetica').fontSize(10);

    const semesterDisplay = semester ? `Semester: ${semester}` : "Semester: Not specified";

    doc.text(`1. UNIVERSITY SEAT NO.: ${row['Seat No'] || ''}     ${semesterDisplay}`, 50, yStart + 80);
    doc.text(`2. NAME OF THE CANDIDATE: ${row['Name'] || ''}`, 50, yStart + 100);
    doc.text("3. SUBJECTS APPLIED:", 50, yStart + 120);

    let subjects = [];
    if (row['Subjects Applied']) {
        subjects = String(row['Subjects Applied']).split(",");
    }

    const startX = 70;
    const startY = yStart + 140;

    let currentX = startX;
    let currentY = startY;
    const maxWidth = width - 100;
    const boxWidth = 35;
    const subjectGap = 10;
    const marginBetweenSubjects = 15;

    subjects.forEach(sub => {
        sub = sub.trim();

        const estimatedTextWidth = sub.length * 6;
        const totalSubjectWidth = estimatedTextWidth + subjectGap + boxWidth + marginBetweenSubjects;

        if (currentX + totalSubjectWidth > maxWidth) {
            currentY += 22;
            currentX = startX;
        }

        doc.text(sub, currentX, currentY);

        const boxX = currentX + estimatedTextWidth + subjectGap;
        doc.rect(boxX, currentY - 5, boxWidth, 15).stroke();

        currentX = boxX + boxWidth + marginBetweenSubjects;
    });

    const signatureY = yStart + 200;

    const hodX = boxLeft + 50;
    const hodText = "Signature of HOD";
    const hodTextWidth = hodText.length * 6;
    const hodSignatureX = boxRight - hodTextWidth - 20;

    doc.text("Signature of Student", hodX, signatureY);
    doc.text(hodText, hodSignatureX, signatureY);
}

function parseMultipartForm(event) {
    return new Promise((resolve, reject) => {
        const busboy = Busboy({
            headers: {
                ...event.headers,
                'content-type': event.headers['content-type'] || event.headers['Content-Type']
            }
        });

        const fields = {};
        const files = {};

        busboy.on('file', (fieldname, file, info) => {
            const chunks = [];
            file.on('data', (data) => chunks.push(data));
            file.on('end', () => {
                files[fieldname] = Buffer.concat(chunks);
            });
        });

        busboy.on('field', (fieldname, value) => {
            fields[fieldname] = value;
        });

        busboy.on('finish', () => {
            resolve({ fields, files });
        });

        busboy.on('error', reject);

        const encoding = event.isBase64Encoded ? 'base64' : 'binary';
        busboy.write(event.body, encoding);
        busboy.end();
    });
}

exports.handler = async (event, context) => {
    if (event.httpMethod !== 'POST') {
        return {
            statusCode: 405,
            body: 'Method Not Allowed'
        };
    }

    try {
        const { fields, files } = await parseMultipartForm(event);

        if (!files.excelFile) {
            return {
                statusCode: 400,
                body: 'No Excel file uploaded.'
            };
        }

        const excelBuffer = files.excelFile;

        // Use uploaded logo if present, otherwise use embedded base64 logo
        const logoData = files.logoFile ? null : LOGO_BASE64;

        const deptName = fields.deptName || "INFORMATION SCIENCE ENGINEERING";
        const examName = fields.examName || "B.E EXAMINATION JUNE / JULY 2025";
        const semester = fields.semester || "";
        const customSubjects = fields.customSubjects ? JSON.parse(fields.customSubjects) : null;
        const useManualSubjects = fields.useManualSubjects === 'true';

        const workbook = xlsx.read(excelBuffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(worksheet);

        const archive = archiver('zip', { zlib: { level: 9 } });
        const chunks = [];

        archive.on('data', (chunk) => chunks.push(chunk));

        const archivePromise = new Promise((resolve, reject) => {
            archive.on('end', () => {
                const buffer = Buffer.concat(chunks);
                resolve(buffer);
            });
            archive.on('error', reject);
        });

        let pageNum = 1;
        for (let i = 0; i < data.length; i += 3) {
            const batch = data.slice(i, i + 3);
            const doc = new PDFDocument({ size: 'A4', margin: 0 });
            const filename = `halltickets_page_${pageNum}.pdf`;

            archive.append(doc, { name: filename });

            const yPositions = [50, 310, 570];

            batch.forEach((row, index) => {
                if (useManualSubjects && customSubjects && customSubjects.length > 0) {
                    row['Subjects Applied'] = customSubjects.join(", ");
                }

                drawTicket(doc, row, yPositions[index], deptName, examName, logoData, semester);
            });

            doc.end();
            pageNum++;
        }

        archive.finalize();

        const zipBuffer = await archivePromise;

        return {
            statusCode: 200,
            headers: {
                'Content-Type': 'application/zip',
                'Content-Disposition': 'attachment; filename="halltickets.zip"'
            },
            body: zipBuffer.toString('base64'),
            isBase64Encoded: true
        };

    } catch (error) {
        console.error('Function error:', error);
        return {
            statusCode: 500,
            body: JSON.stringify({ error: 'Error generating tickets: ' + error.message })
        };
    }
};
