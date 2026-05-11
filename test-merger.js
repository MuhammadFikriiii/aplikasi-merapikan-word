const DocxMerger = require('docx-merger');
const fs = require('fs');
const path = require('path');

const coverPath = path.join(__dirname, 'public', 'COVER 2026.docx');
const sampleDocPath = path.join(__dirname, 'test.docx');
const outputPath = path.join(__dirname, 'merged.docx');

// let's create a quick test docx first using the docx package to merge
const { Document, Packer, Paragraph, TextRun } = require('docx');

async function test() {
    const doc = new Document({
        sections: [{
            children: [
                new Paragraph({ children: [new TextRun("Hello World")] })
            ]
        }]
    });
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(sampleDocPath, buffer);

    const coverBuffer = fs.readFileSync(coverPath);
    const docBuffer = fs.readFileSync(sampleDocPath);

    var docx = new DocxMerger({},[coverBuffer, docBuffer]);
    docx.save('nodebuffer', function(data) {
        fs.writeFileSync(outputPath, data);
        console.log("Merged successfully");
    });
}

test().catch(console.error);
