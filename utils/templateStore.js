const fs = require('fs');
const path = require('path');
const { v4: uuidv4 } = require('uuid');

const TEMPLATES_DIR = path.join(__dirname, '../templates');
const INDEX_FILE = path.join(TEMPLATES_DIR, 'index.json');

function ensureDir() {
    if (!fs.existsSync(TEMPLATES_DIR)) fs.mkdirSync(TEMPLATES_DIR, { recursive: true });
}

function readIndex() {
    ensureDir();
    if (!fs.existsSync(INDEX_FILE)) return { templates: [] };
    try { return JSON.parse(fs.readFileSync(INDEX_FILE, 'utf8')); }
    catch { return { templates: [] }; }
}

function writeIndex(data) {
    ensureDir();
    fs.writeFileSync(INDEX_FILE, JSON.stringify(data, null, 2), 'utf8');
}

function listTemplates() {
    return readIndex().templates;
}

function getTemplate(id) {
    return readIndex().templates.find(t => t.id === id);
}

function addTemplate({ name, type, originalName, filename }) {
    const data = readIndex();
    const template = {
        id: uuidv4(),
        name,
        type,      // 'docx' | 'image'
        originalName,
        filename,
        fields: [], // for image templates: [{name, label, x, y, width, height, fontSize, bold, align}]
        createdAt: new Date().toISOString()
    };
    data.templates.push(template);
    writeIndex(data);
    return template;
}

function updateTemplateLayout(id, fields) {
    const data = readIndex();
    const idx = data.templates.findIndex(t => t.id === id);
    if (idx === -1) throw new Error('Template không tồn tại');
    data.templates[idx].fields = fields;
    writeIndex(data);
    return data.templates[idx];
}

function deleteTemplate(id) {
    const data = readIndex();
    const template = data.templates.find(t => t.id === id);
    if (!template) throw new Error('Template không tồn tại');
    // Remove file
    const filePath = path.join(TEMPLATES_DIR, template.filename);
    if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
    data.templates = data.templates.filter(t => t.id !== id);
    writeIndex(data);
}

function getTemplatePath(filename) {
    return path.join(TEMPLATES_DIR, filename);
}

module.exports = { listTemplates, getTemplate, addTemplate, updateTemplateLayout, deleteTemplate, getTemplatePath, TEMPLATES_DIR };
