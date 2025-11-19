const express = require('express');
const serverless = require('serverless-http');
const multer = require('multer');
const xlsx = require('xlsx');
const PDFDocument = require('pdfkit');
const archiver = require('archiver');
const fs = require('fs');
const path = require('path');

const app = express();
// Use memory storage for serverless environment
const upload = multer({ storage: multer.memoryStorage() });

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Helper function to draw a single ticket
function drawTicket(doc, row, yStart, deptName, examName, logoBuffer, semester) {
    const width = 595.28; // A4 width in points

    // Ticket Boundaries
    const boxTop = yStart - 40;
    const boxBottom = yStart + 240;
    const boxHeight = boxBottom - boxTop;
    const boxLeft = 30;
    const boxRight = width - 30;

    // Draw outer box
    doc.lineWidth(1);
    doc.rect(boxLeft, boxTop, boxRight - boxLeft, boxHeight).stroke();

    // College Logo
    if (logoBuffer) {
        try {
            doc.image(logoBuffer, boxLeft + 10, yStart - 10, { width: 50, height: 50 });
        } catch (e) {
            console.log("Error loading logo:", e);
        }
    }

    // Header
    doc.font('Helvetica-Bold').fontSize(13);
    doc.text("GURU NANAK DEV ENGINEERING COLLEGE, BIDAR", 0, yStart, { align: 'center', width: width });

    // Department name
    doc.font('Helvetica-Bold').fontSize(10);
    doc.text(deptName.toUpperCase(), 0, yStart + 20, { align: 'center', width: width });

    // Exam title
    doc.font('Helvetica-Bold').fontSize(11);
    doc.text(`ADMISSION TICKET FOR ${examName.toUpperCase()}`, 0, yStart + 40, { align: 'center', width: width });

    // Header underline
    doc.lineWidth(0.5);
    doc.moveTo(40, yStart + 50).lineTo(width - 40, yStart + 50).stroke();
    doc.lineWidth(1);

    // Student details
    doc.font('Helvetica').fontSize(10);

    const semesterDisplay = semester ? `Semester: ${semester}` : "Semester: Not specified";

    doc.text(`1. UNIVERSITY SEAT NO.: ${row['Seat No'] || ''}     ${semesterDisplay}`, 50, yStart + 80);
    doc.text(`2. NAME OF THE CANDIDATE: ${row['Name'] || ''}`, 50, yStart + 100);
    doc.text("3. SUBJECTS APPLIED:", 50, yStart + 120);

    let subjects = [];
    if (row['Subjects Applied']) {
        subjects = String(row['Subjects Applied']).split(",");
    }

    // Layout subjects horizontally
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

        // Estimate text width (approx 6pts per char for size 10)
        const estimatedTextWidth = sub.length * 6;
        const totalSubjectWidth = estimatedTextWidth + subjectGap + boxWidth + marginBetweenSubjects;

        // Check if subject fits
        if (currentX + totalSubjectWidth > maxWidth) {
            currentY += 22;
            currentX = startX;
        }

        // Draw subject text
        doc.text(sub, currentX, currentY);

        // Draw signature box
        const boxX = currentX + estimatedTextWidth + subjectGap;
        doc.rect(boxX, currentY - 5, boxWidth, 15).stroke();

        // Move to next position
        currentX = boxX + boxWidth + marginBetweenSubjects;
    });

    // Signatures
    const signatureY = yStart + 200;

    const hodX = boxLeft + 50;
    const hodText = "Signature of HOD";
    const hodTextWidth = hodText.length * 6;
    const hodSignatureX = boxRight - hodTextWidth - 20;

    doc.text("Signature of Student", hodX, signatureY);
    doc.text(hodText, hodSignatureX, signatureY);
}

const router = express.Router();

router.post('/generate', upload.fields([{ name: 'excelFile', maxCount: 1 }, { name: 'logoFile', maxCount: 1 }]), async (req, res) => {
    try {
        if (!req.files['excelFile']) {
            return res.status(400).send('No Excel file uploaded.');
        }

        // Read Excel from buffer
        const excelBuffer = req.files['excelFile'][0].buffer;

        // Handle Logo
        let logoBuffer = null;
        if (req.files['logoFile']) {
            logoBuffer = req.files['logoFile'][0].buffer;
        } else {
            // Read default logo from local file system (bundled with function)
            const defaultLogoPath = path.join(__dirname, 'logo.jpg');
            if (fs.existsSync(defaultLogoPath)) {
                logoBuffer = fs.readFileSync(defaultLogoPath);
            }
        }

        const deptName = req.body.deptName || "INFORMATION SCIENCE ENGINEERING";
        const examName = req.body.examName || "B.E EXAMINATION JUNE / JULY 2025";
        const semester = req.body.semester || "";
        const customSubjects = req.body.customSubjects ? JSON.parse(req.body.customSubjects) : null;
        const useManualSubjects = req.body.useManualSubjects === 'true';

        const workbook = xlsx.read(excelBuffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(worksheet);

        const archive = archiver('zip', {
            zlib: { level: 9 }
        });

        res.attachment('halltickets.zip');
        archive.pipe(res);

        // Process students in batches of 3
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

                drawTicket(doc, row, yPositions[index], deptName, examName, logoBuffer, semester);
            });

            doc.end();
            pageNum++;
        }

        archive.finalize();

    } catch (error) {
        console.error(error);
        res.status(500).send('Error generating tickets');
    }
});

app.use('/.netlify/functions/api', router);

module.exports.handler = serverless(app);
