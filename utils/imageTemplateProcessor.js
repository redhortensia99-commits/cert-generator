/**
 * Generate a merged Word document from an image template + student data.
 * Each page = certificate background image + positioned floating text boxes.
 *
 * Uses raw Word XML built with PizZip (no external Word library needed).
 * A4 page in EMU: 7560000 x 10692000
 */
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');

// A4 in EMU (1 inch = 914400 EMU, 1 cm = 360000 EMU)
const PAGE_W_EMU = 7560000;  // 21 cm
const PAGE_H_EMU = 10692000; // 29.7 cm
// A4 in twips (for w:pgSz)
const PAGE_W_TWIPS = 11906;
const PAGE_H_TWIPS = 16838;

function escapeXml(str) {
    return String(str || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

function buildBackgroundImageXml(relId, drawingId) {
    return `<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="${PAGE_W_EMU}" cy="${PAGE_H_EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${drawingId}" name="bg${drawingId}"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${drawingId}" name="bg"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${PAGE_W_EMU}" cy="${PAGE_H_EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`;
}

function buildTextBoxXml(text, field, drawingId) {
    const xEmu = Math.round((field.x / 100) * PAGE_W_EMU);
    const yEmu = Math.round((field.y / 100) * PAGE_H_EMU);
    const wEmu = Math.round(((field.width || 40) / 100) * PAGE_W_EMU);
    const hEmu = Math.round(((field.height || 5) / 100) * PAGE_H_EMU);
    const fontSize = field.fontSize || 22; // half-points, 22 = 11pt
    const bold = field.bold ? '<w:b/><w:bCs/>' : '';
    const align = field.align || 'left';
    const dispText = escapeXml(text);

    return `<w:p><w:pPr><w:jc w:val="${align}"/></w:pPr><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${xEmu}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>${yEmu}</wp:posOffset></wp:positionV><wp:extent cx="${wEmu}" cy="${hEmu}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${drawingId}" name="tb${drawingId}"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr txBx="1"><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr><wps:spPr><a:xfrm><a:off x="${xEmu}" y="${yEmu}"/><a:ext cx="${wEmu}" cy="${hEmu}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="${align}"/></w:pPr><w:r><w:rPr>${bold}<w:sz w:val="${fontSize}"/><w:szCs w:val="${fontSize}"/></w:rPr><w:t xml:space="preserve">${dispText}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr rot="0" anchor="t" lIns="36000" rIns="36000" tIns="36000" bIns="36000"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`;
}

function buildStudentDataMap(student) {
    return {
        ho_ten: student.ho_ten || '',
        ngay_sinh: student.ngay_sinh || '',
        noi_sinh: student.noi_sinh || '',
        khoa_hoc: student.khoa_hoc || '',
        tu_ngay: student.tu_ngay || '',
        den_ngay: student.den_ngay || '',
        ket_qua: student.xep_loai || student.ket_qua || '',
        xep_loai: student.xep_loai || '',
        so_quyet_dinh: student.so_quyet_dinh || '',
        so_vao_so: student.so_vao_so || '',
        ngay_ky: student.ngay_ky || '',
        thang_ky: student.thang_ky || '',
        nam_ky: student.nam_ky || '',
    };
}

async function generateFromImageTemplate(templatePath, students, fields) {
    const imageBuffer = fs.readFileSync(templatePath);
    const ext = path.extname(templatePath).replace('.', '').toLowerCase();
    const mimeType = ext === 'png' ? 'image/png' : 'image/jpeg';
    const relType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

    // --- Build document XML body ---
    let bodyContent = '';
    let drawingCounter = 1;
    const bgRelId = 'rId100'; // Fixed rId for the shared background image

    for (let i = 0; i < students.length; i++) {
        const student = students[i];
        const dataMap = buildStudentDataMap(student);

        // Background image
        bodyContent += buildBackgroundImageXml(bgRelId, drawingCounter++);

        // Text boxes for each field
        for (const field of fields) {
            const value = dataMap[field.name] || '';
            bodyContent += buildTextBoxXml(value, field, drawingCounter++);
        }

        // Section break (page break) between students, last student uses the outer sectPr
        if (i < students.length - 1) {
            bodyContent += `<w:p><w:pPr><w:sectPr><w:pgSz w:w="${PAGE_W_TWIPS}" w:h="${PAGE_H_TWIPS}"/><w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr></w:pPr></w:p>`;
        }
    }

    // Final section
    bodyContent += `<w:sectPr><w:pgSz w:w="${PAGE_W_TWIPS}" w:h="${PAGE_H_TWIPS}"/><w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>`;

    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" mc:Ignorable="w14 wp14">
<w:body>${bodyContent}</w:body></w:document>`;

    const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="${bgRelId}" Type="${relType}" Target="media/cert_bg.${ext}"/>
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`;

    const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Default Extension="${ext}" ContentType="${mimeType}"/>
</Types>`;

    const dotRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/><w:rPr><w:sz w:val="22"/></w:rPr></w:style>
</w:styles>`;

    const settingsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
</w:settings>`;

    // Assemble DOCX using PizZip
    const zip = new PizZip();
    zip.file('[Content_Types].xml', contentTypesXml);
    zip.file('_rels/.rels', dotRelsXml);
    zip.file('word/document.xml', documentXml);
    zip.file('word/_rels/document.xml.rels', relsXml);
    zip.file('word/styles.xml', stylesXml);
    zip.file('word/settings.xml', settingsXml);
    zip.file(`word/media/cert_bg.${ext}`, imageBuffer, { binary: true });

    return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
}

module.exports = { generateFromImageTemplate };
