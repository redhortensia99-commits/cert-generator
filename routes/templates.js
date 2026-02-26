const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { v4: uuidv4 } = require('uuid');
const { listTemplates, getTemplate, addTemplate, updateTemplateLayout, deleteTemplate, TEMPLATES_DIR } = require('../utils/templateStore');

const router = express.Router();

const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        if (!fs.existsSync(TEMPLATES_DIR)) fs.mkdirSync(TEMPLATES_DIR, { recursive: true });
        cb(null, TEMPLATES_DIR);
    },
    filename: (req, file, cb) => {
        const ext = path.extname(file.originalname).toLowerCase();
        cb(null, `tpl_${uuidv4()}${ext}`);
    }
});

const upload = multer({
    storage,
    limits: { fileSize: 30 * 1024 * 1024 },
    fileFilter: (req, file, cb) => {
        const allowed = ['.docx', '.jpg', '.jpeg', '.png', '.gif', '.webp'];
        const ext = path.extname(file.originalname).toLowerCase();
        if (allowed.includes(ext)) return cb(null, true);
        cb(new Error('Chỉ chấp nhận file .docx, .jpg, .png'));
    }
});

// GET /api/templates – list all
router.get('/templates', (req, res) => {
    try {
        res.json(listTemplates());
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// POST /api/templates – upload new template
router.post('/templates', upload.single('file'), (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'Vui lòng chọn file mẫu.' });
    const name = req.body.name || req.file.originalname;
    const ext = path.extname(req.file.originalname).toLowerCase();
    const type = ext === '.docx' ? 'docx' : 'image';
    try {
        const template = addTemplate({
            name,
            type,
            originalName: req.file.originalname,
            filename: req.file.filename
        });
        res.json(template);
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// PUT /api/templates/:id/layout – save field positions for image template
router.put('/templates/:id/layout', express.json(), (req, res) => {
    try {
        const { fields } = req.body;
        if (!Array.isArray(fields)) return res.status(400).json({ error: 'fields phải là array.' });
        const updated = updateTemplateLayout(req.params.id, fields);
        res.json(updated);
    } catch (err) {
        res.status(400).json({ error: err.message });
    }
});

// DELETE /api/templates/:id
router.delete('/templates/:id', (req, res) => {
    try {
        deleteTemplate(req.params.id);
        res.json({ success: true });
    } catch (err) {
        res.status(400).json({ error: err.message });
    }
});

// GET /api/templates/:id/preview – serve template image
router.get('/templates/:id/preview', (req, res) => {
    try {
        const template = getTemplate(req.params.id);
        if (!template) return res.status(404).json({ error: 'Không tìm thấy template.' });
        if (template.type !== 'image') return res.status(400).json({ error: 'Không phải image template.' });
        const filePath = path.join(TEMPLATES_DIR, template.filename);
        if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'File không tồn tại.' });
        res.sendFile(filePath);
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

module.exports = router;
