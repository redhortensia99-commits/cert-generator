const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

/**
 * Parse Excel file and extract student data + embedded images
 * @param {string} filePath - Path to .xlsx file
 * @returns {{ students: Object[], imageMap: Object }}
 */
async function parseExcel(filePath) {
    const workbook = xlsx.readFile(filePath, {
        cellStyles: true,
        cellHTML: false,
        sheetStubs: false,
        type: 'file'
    });

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Convert sheet to JSON (first row = headers)
    const rows = xlsx.utils.sheet_to_json(sheet, {
        header: 1,
        raw: false,
        defval: ''
    });

    if (rows.length < 2) {
        throw new Error('File Excel phải có ít nhất 1 dòng tiêu đề và 1 dòng dữ liệu.');
    }

    const headers = rows[0].map(h => String(h).trim());
    const dataRows = rows.slice(1).filter(row => row.some(cell => cell !== '' && cell !== null && cell !== undefined));

    const students = dataRows.map((row, idx) => {
        const student = {};
        headers.forEach((header, colIdx) => {
            student[header] = row[colIdx] !== undefined ? String(row[colIdx]).trim() : '';
        });
        student._rowIndex = idx + 1; // 1-based data row index
        return student;
    });

    // Extract embedded images from Excel
    const imageMap = {};

    if (workbook.Custprops) {
        // Sometimes images are in custom properties
    }

    // Try to extract images using xlsx's image extraction
    if (workbook.Sheets[sheetName]['!images']) {
        const images = workbook.Sheets[sheetName]['!images'];
        const headerRow = 0; // Headers are row 0

        images.forEach(img => {
            if (!img.position) return;
            const row = img.position.r; // 0-indexed row
            const col = img.position.c; // 0-indexed col

            if (row === 0) return; // Skip header row

            const dataRowIdx = row - 1; // Convert to 0-based data row index
            if (!imageMap[dataRowIdx]) imageMap[dataRowIdx] = {};

            // Map column index to header name
            const headerName = headers[col];
            if (headerName && img.data) {
                imageMap[dataRowIdx][headerName] = {
                    data: img.data,
                    type: img.type || 'png',
                    extension: img.extension || 'png'
                };
            }
        });
    }

    // Alternative: Use xlsx-image-parser approach via raw ZIP
    try {
        const JSZip = require('jszip');
        const rawBuffer = fs.readFileSync(filePath);
        const zip = await JSZip.loadAsync(rawBuffer);

        // Read xl/drawings/_rels/ to map images to cells
        const drawingRelFiles = Object.keys(zip.files).filter(f =>
            f.startsWith('xl/drawings/_rels/') && f.endsWith('.xml.rels')
        );

        const drawingFiles = Object.keys(zip.files).filter(f =>
            f.startsWith('xl/drawings/drawing') && f.endsWith('.xml') && !f.includes('_rels')
        );

        // Build a map: imagePath -> buffer
        const rawImageBuffers = {};
        const mediaFiles = Object.keys(zip.files).filter(f => f.startsWith('xl/media/'));
        for (const mf of mediaFiles) {
            rawImageBuffers[mf] = await zip.file(mf).async('nodebuffer');
        }

        // Parse each drawing XML to get anchor info (row, col -> image)
        for (const drawingFile of drawingFiles) {
            const drawingContent = await zip.file(drawingFile).async('string');
            const relFile = drawingFile.replace('xl/drawings/', 'xl/drawings/_rels/') + '.rels';

            if (!zip.files[relFile]) continue;
            const relContent = await zip.file(relFile).async('string');

            // Parse relationships: Id -> Target (image path)
            const relMap = {};
            const relMatches = relContent.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g);
            for (const match of relMatches) {
                const id = match[1];
                const target = match[2].replace('..', 'xl');
                relMap[id] = target;
            }

            // Parse two-cell anchors
            const anchorRegex = /<xdr:twoCellAnchor[^>]*>([\s\S]*?)<\/xdr:twoCellAnchor>/g;
            const picRegex = /<xdr:pic>([\s\S]*?)<\/xdr:pic>/;
            const fromRegex = /<xdr:from>([\s\S]*?)<\/xdr:from>/;
            const rowRegex = /<xdr:row>(\d+)<\/xdr:row>/;
            const colRegex = /<xdr:col>(\d+)<\/xdr:col>/;
            const relIdRegex = /r:embed="([^"]+)"/;

            let anchorMatch;
            while ((anchorMatch = anchorRegex.exec(drawingContent)) !== null) {
                const anchorContent = anchorMatch[1];
                const fromMatch = fromRegex.exec(anchorContent);
                const picMatch = picRegex.exec(anchorContent);

                if (!fromMatch || !picMatch) continue;

                const fromContent = fromMatch[1];
                const rows = [...fromContent.matchAll(rowRegex)];
                const cols = [...fromContent.matchAll(colRegex)];

                if (rows.length === 0 || cols.length === 0) continue;

                const row = parseInt(rows[0][1]); // 0-indexed
                const col = parseInt(cols[0][1]); // 0-indexed

                if (row === 0) continue; // header row

                const relIdMatch = relIdRegex.exec(picMatch[1]);
                if (!relIdMatch) continue;

                const relId = relIdMatch[1];
                const imagePath = relMap[relId];

                if (!imagePath || !rawImageBuffers[imagePath]) continue;

                const dataRowIdx = row - 1; // 0-indexed data row
                if (!imageMap[dataRowIdx]) imageMap[dataRowIdx] = {};

                const headerName = headers[col];
                if (headerName) {
                    const ext = path.extname(imagePath).replace('.', '') || 'png';
                    imageMap[dataRowIdx][headerName] = {
                        data: rawImageBuffers[imagePath],
                        extension: ext
                    };
                }
            }
        }
    } catch (zipErr) {
        console.warn('⚠️ Không thể đọc ảnh từ Excel ZIP:', zipErr.message);
    }

    return { students, imageMap };
}

module.exports = { parseExcel };
