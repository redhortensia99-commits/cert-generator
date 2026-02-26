const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { parseExcel } = require('../utils/excelParser');
const { generateCertificates } = require('../utils/wordGenerator');
const { generateFromImageTemplate } = require('../utils/imageTemplateProcessor');
const { getTemplate, getTemplatePath } = require('../utils/templateStore');

const router = express.Router();

const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        const d = path.join(__dirname, '../uploads');
        if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
        cb(null, d);
    },
    filename: (req, file, cb) => {
        cb(null, Date.now() + '_' + Math.random().toString(36).slice(2) + path.extname(file.originalname));
    }
});

const upload = multer({
    storage,
    limits: { fileSize: 50 * 1024 * 1024 },
    fileFilter: (req, file, cb) => {
        const allowed = ['.xlsx', '.xls'];
        if (allowed.includes(path.extname(file.originalname).toLowerCase())) return cb(null, true);
        cb(new Error('Chỉ chấp nhận file Excel (.xlsx, .xls)'));
    }
});

router.post('/generate', upload.single('excel'), async (req, res) => {
    const excelFile = req.file;
    const templateId = req.body.templateId;

    if (!excelFile) return res.status(400).json({ error: 'Vui lòng tải lên file Excel.' });
    if (!templateId) return res.status(400).json({ error: 'Vui lòng chọn mẫu chứng chỉ.' });

    const template = getTemplate(templateId);
    if (!template) return res.status(404).json({ error: 'Mẫu chứng chỉ không tồn tại.' });

    try {
        console.log(`📊 Đọc Excel: ${excelFile.originalname}`);
        const { students, imageMap } = await parseExcel(excelFile.path);

        if (!students || students.length === 0) {
            return res.status(400).json({ error: 'File Excel không có dữ liệu học viên.' });
        }

        console.log(`✅ ${students.length} học viên | Template: ${template.name} (${template.type})`);

        const templateFilePath = getTemplatePath(template.filename);
        let outputBuffer;

        if (template.type === 'docx') {
            outputBuffer = await generateCertificates(templateFilePath, students, imageMap);
        } else if (template.type === 'image') {
            if (!template.fields || template.fields.length === 0) {
                return res.status(400).json({ error: 'Mẫu ảnh chưa được cài đặt vị trí các trường. Vui lòng vào "Quản Lý Mẫu" để cài đặt.' });
            }
            outputBuffer = await generateFromImageTemplate(templateFilePath, students, template.fields);
        } else {
            return res.status(400).json({ error: `Loại mẫu "${template.type}" chưa được hỗ trợ.` });
        }

        try { fs.unlinkSync(excelFile.path); } catch (e) { }

        const filename = `chung-chi-${Date.now()}.docx`;
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);
        res.setHeader('X-Student-Count', students.length);
        res.send(outputBuffer);

        console.log(`✅ Hoàn thành ${students.length} chứng chỉ.`);
    } catch (err) {
        console.error('❌ Lỗi:', err);
        try { if (excelFile) fs.unlinkSync(excelFile.path); } catch (e) { }
        res.status(500).json({ error: err.message || 'Có lỗi xảy ra khi tạo chứng chỉ.' });
    }
});

module.exports = router;
