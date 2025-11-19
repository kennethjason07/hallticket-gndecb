const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const PDFDocument = require('pdfkit');
const archiver = require('archiver');
const fs = require('fs');
const path = require('path');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(express.static('public'));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Helper function to draw a single ticket
function drawTicket(doc, row, yStart, deptName, examName, logoPath, semester) {
    const width = 595.28; // A4 width in points

    // Ticket Boundaries
    // Python: box_top = y_start + 40, box_bottom = y_start - 240
    // Node (Top-Down): boxTop = yStart - 40, boxBottom = yStart + 240
    const boxTop = yStart - 40;
    const boxBottom = yStart + 240;
    const boxHeight = boxBottom - boxTop;
    const boxLeft = 30;
    const boxRight = width - 30;

    // Draw outer box
    doc.lineWidth(1);
    doc.rect(boxLeft, boxTop, boxRight - boxLeft, boxHeight).stroke();

    // College Logo
    // Python: y_start + 10 (relative to y_start, but below box top)
    // Node: yStart - 10
    if (logoPath && fs.existsSync(logoPath)) {
        try {
            doc.image(logoPath, boxLeft + 10, yStart - 10, { width: 50, height: 50 });
        } catch (e) {
            console.log("Error loading logo:", e);
        }
    }

    // Header
    // Python: y_start
    doc.font('Helvetica-Bold').fontSize(13);
    doc.text("GURU NANAK DEV ENGINEERING COLLEGE, BIDAR", 0, yStart, { align: 'center', width: width });

    // Department name
    // Python: y_start - 20
    doc.font('Helvetica-Bold').fontSize(10);
    doc.text(deptName.toUpperCase(), 0, yStart + 20, { align: 'center', width: width });

    // Exam title
    // Python: y_start - 40
    doc.font('Helvetica-Bold').fontSize(11);
    doc.text(`ADMISSION TICKET FOR ${examName.toUpperCase()}`, 0, yStart + 40, { align: 'center', width: width });

    // Header underline
    // Python: y_start - 50
    doc.lineWidth(0.5);
    doc.moveTo(40, yStart + 50).lineTo(width - 40, yStart + 50).stroke();
    doc.lineWidth(1);

    // Student details
    doc.font('Helvetica').fontSize(10);

    const semesterDisplay = semester ? `Semester: ${semester}` : "Semester: Not specified";

    // Python: y_start - 80
    doc.text(`1. UNIVERSITY SEAT NO.: ${row['Seat No'] || ''}     ${semesterDisplay}`, 50, yStart + 80);

    // Python: y_start - 100
    doc.text(`2. NAME OF THE CANDIDATE: ${row['Name'] || ''}`, 50, yStart + 100);

    // Subjects section
    // Python: y_start - 120
    doc.text("3. SUBJECTS APPLIED:", 50, yStart + 120);

    let subjects = [];
    if (row['Subjects Applied']) {
        subjects = String(row['Subjects Applied']).split(",");
    }

    // Layout subjects horizontally
    const startX = 70;
    // Python: y_start - 140
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
            // Move to next line
            // Python: current_y -= 22
            // Node: currentY += 22
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
    // Python: box_bottom + 40 -> (y_start - 240) + 40 = y_start - 200
    // Node: yStart + 200
    const signatureY = yStart + 200;

    const hodX = boxLeft + 50;
    const hodText = "Signature of HOD";
    const hodTextWidth = hodText.length * 6;
    const hodSignatureX = boxRight - hodTextWidth - 20;

    doc.text("Signature of Student", hodX, signatureY);
    doc.text(hodText, hodSignatureX, signatureY);
}

app.post('/generate', upload.fields([{ name: 'excelFile', maxCount: 1 }, { name: 'logoFile', maxCount: 1 }]), async (req, res) => {
    try {
        if (!req.files['excelFile']) {
            return res.status(400).send('No Excel file uploaded.');
        }

        const excelPath = req.files['excelFile'][0].path;
        const logoPath = req.files['logoFile'] ? req.files['logoFile'][0].path : null;
        const deptName = req.body.deptName || "INFORMATION SCIENCE ENGINEERING";
        const examName = req.body.examName || "B.E EXAMINATION JUNE / JULY 2025";
        const semester = req.body.semester || "";
        const customSubjects = req.body.customSubjects ? JSON.parse(req.body.customSubjects) : null;
        const useManualSubjects = req.body.useManualSubjects === 'true';

        const workbook = xlsx.readFile(excelPath);
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

            const width = 595.28;
            const height = 841.89;

            // Y Positions for 3 tickets
            // Python: [height-50, height-310, height-570]
            // Node (Top-Down): 
            // Ticket 1: yStart = 50
            // Ticket 2: yStart = 310
            // Ticket 3: yStart = 570
            const yPositions = [50, 310, 570];

            batch.forEach((row, index) => {
                if (useManualSubjects && customSubjects && customSubjects.length > 0) {
                    row['Subjects Applied'] = customSubjects.join(", ");
                }

                drawTicket(doc, row, yPositions[index], deptName, examName, logoPath, semester);
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

const PORT = 3000;
app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
});
