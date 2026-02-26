const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const ImageModule = require('docxtemplater-image-module-free');

/**
 * Generate merged certificate Word document for all students.
 * @param {string} templatePath - Path to .docx template file
 * @param {Array} students - Array of student objects
 * @param {Object} imageMap - Map of { rowIdx: { anh_the: {data, extension}, chu_ky: {data, extension} } }
 * @returns {Buffer} - Buffer of the generated merged .docx file
 */
async function generateCertificates(templatePath, students, imageMap) {
    const templateBuffer = fs.readFileSync(templatePath);
    const outputDocuments = [];

    for (let i = 0; i < students.length; i++) {
        const student = students[i];
        const images = imageMap[i] || {};

        try {
            const doc = await fillTemplate(templateBuffer, student, images);
            outputDocuments.push(doc);
            console.log(`  [${i + 1}/${students.length}] ✅ ${student.ho_ten || 'Học viên ' + (i + 1)}`);
        } catch (err) {
            console.error(`  [${i + 1}/${students.length}] ❌ Lỗi học viên ${student.ho_ten}:`, err.message);
            throw new Error(`Lỗi tạo chứng chỉ cho học viên "${student.ho_ten || i + 1}": ${err.message}`);
        }
    }

    if (outputDocuments.length === 1) {
        return outputDocuments[0];
    }

    return mergeDocuments(outputDocuments);
}

/**
 * Fill a single certificate template for one student.
 */
async function fillTemplate(templateBuffer, student, images) {
    // Build image getter
    const imageGetter = (tagName) => {
        const img = images[tagName];
        if (img && img.data) {
            return img.data;
        }
        return null;
    };

    // Check if template has image placeholders
    const hasImages = Object.keys(images).length > 0;

    let imageModule = null;
    if (hasImages) {
        imageModule = new ImageModule({
            centered: false,
            fileType: 'docx',
            getImage: (tagValue, tagName) => {
                return imageGetter(tagName);
            },
            getSize: (img, tagValue, tagName) => {
                if (tagName === 'anh_the') return [85, 113]; // 3x4 cm in points roughly
                if (tagName === 'chu_ky') return [150, 60];
                return [100, 80];
            }
        });
    }

    const zip = new PizZip(templateBuffer);

    const docOptions = {
        paragraphLoop: true,
        linebreaks: true,
    };

    if (imageModule) {
        docOptions.modules = [imageModule];
    }

    const doc = new Docxtemplater(zip, docOptions);

    // Build data object for template
    const data = buildTemplateData(student, images);

    doc.render(data);

    return doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });
}

/**
 * Build the data object to pass to docxtemplater
 */
function buildTemplateData(student, images) {
    const data = {
        // Text fields
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

    // Add image fields if available
    if (images.anh_the) {
        data.anh_the = images.anh_the.data;
    }
    if (images.chu_ky) {
        data.chu_ky = images.chu_ky.data;
    }

    return data;
}

/**
 * Merge multiple .docx Buffers into a single document.
 * Strategy: take the first document as base, append XML body content from others.
 */
function mergeDocuments(buffers) {
    if (buffers.length === 0) throw new Error('Không có tài liệu để gộp');
    if (buffers.length === 1) return buffers[0];

    const JSZip = require('jszip');

    // We'll do synchronous merging using PizZip
    const baseZip = new PizZip(buffers[0]);
    let baseXml = baseZip.file('word/document.xml').asText();

    // Parse base body content
    // We need to insert page breaks + content from other docs into the base body
    const bodyCloseTag = '</w:body>';
    const bodyCloseIdx = baseXml.lastIndexOf(bodyCloseTag);

    if (bodyCloseIdx === -1) {
        throw new Error('Không thể phân tích file Word template (thiếu body tag).');
    }

    let insertionContent = '';

    for (let i = 1; i < buffers.length; i++) {
        const otherZip = new PizZip(buffers[i]);
        const otherXml = otherZip.file('word/document.xml').asText();

        // Extract body content from the other doc
        const bodyContent = extractBodyContent(otherXml);
        if (bodyContent) {
            // Add a page break before each new certificate
            const pageBreak = `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
            insertionContent += pageBreak + bodyContent;
        }

        // Merge any images from other docs into base zip
        mergeMediaFiles(baseZip, otherZip, i);
    }

    // Insert before closing body tag
    const newXml = baseXml.slice(0, bodyCloseIdx) + insertionContent + baseXml.slice(bodyCloseIdx);
    baseZip.file('word/document.xml', newXml);

    return baseZip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
}

/**
 * Extract the inner content of <w:body>...</w:body> (excluding the sectPr at the end)
 */
function extractBodyContent(xml) {
    const bodyStart = xml.indexOf('<w:body>');
    const bodyEnd = xml.lastIndexOf('</w:body>');

    if (bodyStart === -1 || bodyEnd === -1) return null;

    let content = xml.slice(bodyStart + '<w:body>'.length, bodyEnd);

    // Remove the last <w:sectPr> block (section properties) to avoid duplicate section breaks
    const lastSectPr = content.lastIndexOf('<w:sectPr');
    if (lastSectPr !== -1) {
        const sectPrEnd = content.lastIndexOf('</w:sectPr>');
        if (sectPrEnd !== -1) {
            content = content.slice(0, lastSectPr) + content.slice(sectPrEnd + '</w:sectPr>'.length);
        }
    }

    return content.trim();
}

/**
 * Copy media (images) from other zip into base zip, renaming to avoid conflicts.
 */
function mergeMediaFiles(baseZip, otherZip, docIndex) {
    try {
        const otherFiles = Object.keys(otherZip.files);
        const mediaFiles = otherFiles.filter(f => f.startsWith('word/media/'));

        mediaFiles.forEach(mediaPath => {
            const filename = path.basename(mediaPath);
            const ext = path.extname(filename);
            const newFilename = `word/media/doc${docIndex}_${filename}`;

            // Check if this media already exists in base
            if (!baseZip.files[newFilename]) {
                const data = otherZip.file(mediaPath).asBinary();
                baseZip.file(newFilename, data, { binary: true });
            }
        });

        // Also update the document.xml to reference new media filenames
        // This is a simplistic approach - for a more robust solution, we'd need to
        // also update word/_rels/document.xml.rels in the base
        const otherRelsPath = 'word/_rels/document.xml.rels';
        if (otherZip.files[otherRelsPath]) {
            const otherRels = otherZip.file(otherRelsPath).asText();
            const baseRelsPath = 'word/_rels/document.xml.rels';
            let baseRels = baseZip.file(baseRelsPath) ? baseZip.file(baseRelsPath).asText() : '';

            // Extract media relationships from other doc
            const relMatches = [...otherRels.matchAll(/<Relationship[^>]*Target="media\/([^"]+)"[^>]*Id="([^"]+)"[^>]*\/>/g)];

            relMatches.forEach(match => {
                const originalFilename = match[1];
                const relId = match[2];
                const newFilename = `doc${docIndex}_${originalFilename}`;
                const newRelId = `doc${docIndex}_${relId}`;

                // Add to base rels if not already there
                if (!baseRels.includes(`Id="${newRelId}"`)) {
                    const newRel = `<Relationship Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${newFilename}"/>`;
                    baseRels = baseRels.replace('</Relationships>', newRel + '</Relationships>');
                }
            });

            baseZip.file(baseRelsPath, baseRels);
        }
    } catch (e) {
        console.warn('⚠️ Không thể gộp file media:', e.message);
    }
}

module.exports = { generateCertificates };
