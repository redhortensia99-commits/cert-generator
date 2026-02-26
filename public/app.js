// ==================== STATE ====================
let templates = [];
let selectedExcelFile = null;
let downloadBlob = null;
let currentModalTemplateId = null;
let selectedFieldName = null;
let placedFields = []; // {name, label, x, y, width, height, fontSize, bold, align}

const ALL_FIELDS = [
    { name: 'ho_ten', label: 'Họ và tên' },
    { name: 'ngay_sinh', label: 'Ngày sinh' },
    { name: 'noi_sinh', label: 'Nơi sinh' },
    { name: 'khoa_hoc', label: 'Khoá học' },
    { name: 'tu_ngay', label: 'Từ ngày' },
    { name: 'den_ngay', label: 'Đến ngày' },
    { name: 'ket_qua', label: 'Kết quả' },
    { name: 'so_quyet_dinh', label: 'Số QĐ' },
    { name: 'so_vao_so', label: 'Số vào sổ' },
    { name: 'ngay_ky', label: 'Ngày ký' },
    { name: 'thang_ky', label: 'Tháng ký' },
    { name: 'nam_ky', label: 'Năm ký' },
];

// ==================== INIT ====================
document.addEventListener('DOMContentLoaded', () => {
    loadTemplates();
    setupExcelDrop();
    document.getElementById('tpl-file').addEventListener('change', e => {
        const f = e.target.files[0];
        if (f) {
            const lbl = document.getElementById('tpl-file-label');
            lbl.textContent = f.name;
            lbl.classList.add('has-file');
        }
    });
    document.getElementById('gen-template-select').addEventListener('change', checkGenerateReady);
});

// ==================== TABS ====================
function switchTab(tab) {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.panel').forEach(p => p.classList.add('hidden'));
    document.getElementById(`tab-${tab}`).classList.add('active');
    document.getElementById(`panel-${tab}`).classList.remove('hidden');
    if (tab === 'generate') refreshTemplateSelect();
}

// ==================== TEMPLATE MANAGEMENT ====================
async function loadTemplates() {
    try {
        const res = await fetch('/api/templates');
        templates = await res.json();
        renderTemplateList();
    } catch (e) {
        document.getElementById('tpl-list').innerHTML = '<div class="empty-state">⚠️ Không thể tải danh sách mẫu.</div>';
    }
}

function renderTemplateList() {
    const el = document.getElementById('tpl-list');
    if (!templates.length) {
        el.innerHTML = '<div class="empty-state">📭 Chưa có mẫu nào. Hãy tải lên mẫu đầu tiên!</div>';
        return;
    }
    el.innerHTML = templates.map(t => {
        const icon = t.type === 'docx' ? '📄' : '🖼️';
        const fieldCount = (t.fields || []).length;
        const needsSetup = t.type === 'image' && fieldCount === 0;
        return `<div class="tpl-item">
      <span class="tpl-icon">${icon}</span>
      <div class="tpl-info">
        <div class="tpl-name">${esc(t.name)}</div>
        <div class="tpl-meta">
          <span class="tpl-badge ${t.type}">${t.type === 'docx' ? 'Word Template' : 'Ảnh Template'}</span>
          ${t.type === 'image' ? `<span style="margin-left:8px;color:${needsSetup ? 'var(--warning)' : 'var(--success)'}">
            ${needsSetup ? '⚠️ Chưa cài đặt vị trí' : `✅ ${fieldCount} trường đã đặt`}</span>` : ''}
        </div>
      </div>
      <div class="tpl-actions">
        ${t.type === 'image' ? `<button class="btn-sm config" onclick="openFieldModal('${t.id}')">🎯 Cài đặt vị trí</button>` : ''}
        <button class="btn-sm del" onclick="deleteTemplate('${t.id}')">🗑️ Xoá</button>
      </div>
    </div>`;
    }).join('');
}

async function uploadTemplate() {
    const name = document.getElementById('tpl-name').value.trim();
    const fileInput = document.getElementById('tpl-file');
    const file = fileInput.files[0];
    if (!name) return alert('Vui lòng nhập tên mẫu chứng chỉ.');
    if (!file) return alert('Vui lòng chọn file mẫu.');

    const fd = new FormData();
    fd.append('name', name);
    fd.append('file', file);

    try {
        const res = await fetch('/api/templates', { method: 'POST', body: fd });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error || 'Upload thất bại');
        document.getElementById('tpl-name').value = '';
        fileInput.value = '';
        const lbl = document.getElementById('tpl-file-label');
        lbl.textContent = 'Chưa chọn';
        lbl.classList.remove('has-file');
        await loadTemplates();
        alert(`✅ Đã tải lên mẫu "${name}" thành công!${data.type === 'image' ? '\n\nHãy nhấn "Cài đặt vị trí" để thiết lập vị trí các trường dữ liệu trên ảnh.' : ''}`);
    } catch (e) {
        alert('❌ ' + e.message);
    }
}

async function deleteTemplate(id) {
    const t = templates.find(t => t.id === id);
    if (!t) return;
    if (!confirm(`Xoá mẫu "${t.name}"?`)) return;
    try {
        const res = await fetch(`/api/templates/${id}`, { method: 'DELETE' });
        if (!res.ok) throw new Error((await res.json()).error);
        await loadTemplates();
    } catch (e) {
        alert('❌ ' + e.message);
    }
}

function refreshTemplateSelect() {
    const sel = document.getElementById('gen-template-select');
    const curr = sel.value;
    sel.innerHTML = '<option value="">-- Chọn mẫu chứng chỉ --</option>';
    templates.forEach(t => {
        const opt = document.createElement('option');
        opt.value = t.id;
        opt.textContent = t.name + (t.type === 'image' ? ' 🖼️' : ' 📄');
        sel.appendChild(opt);
    });
    if (curr) sel.value = curr;
    checkGenerateReady();
}

// ==================== FIELD PLACEMENT MODAL ====================
function openFieldModal(templateId) {
    const t = templates.find(t => t.id === templateId);
    if (!t) return;

    currentModalTemplateId = templateId;
    placedFields = JSON.parse(JSON.stringify(t.fields || []));
    selectedFieldName = null;

    // Set image
    const img = document.getElementById('tpl-preview-img');
    img.src = `/api/templates/${templateId}/preview?t=${Date.now()}`;

    renderFieldBtns();
    renderFieldMarkers();
    renderPlacedList();

    document.getElementById('field-modal').classList.remove('hidden');
    document.body.style.overflow = 'hidden';

    // Click to place
    const container = document.getElementById('canvas-container');
    container.onclick = handleCanvasClick;
}

function closeFieldModal() {
    document.getElementById('field-modal').classList.add('hidden');
    document.body.style.overflow = '';
    currentModalTemplateId = null;
    selectedFieldName = null;
    const container = document.getElementById('canvas-container');
    container.classList.remove('placing');
    container.onclick = null;
}

function renderFieldBtns() {
    const el = document.getElementById('field-btns');
    el.innerHTML = ALL_FIELDS.map(f => {
        const placed = placedFields.find(p => p.name === f.name);
        return `<button class="field-btn ${selectedFieldName === f.name ? 'selected' : ''} ${placed ? 'placed' : ''}"
      onclick="selectField('${f.name}')" style="${placed ? 'opacity:.5;' : ''}">
      ${placed ? '✓ ' : ''}${f.label} <small style="color:var(--text-m)">(${f.name})</small>
    </button>`;
    }).join('');
}

function selectField(name) {
    selectedFieldName = name;
    renderFieldBtns();
    document.getElementById('canvas-container').classList.add('placing');
}

function handleCanvasClick(e) {
    if (!selectedFieldName) return;
    const container = document.getElementById('canvas-container');
    const img = document.getElementById('tpl-preview-img');
    const rect = img.getBoundingClientRect();
    const containerRect = container.getBoundingClientRect();

    const xPct = ((e.clientX - rect.left) / rect.width) * 100;
    const yPct = ((e.clientY - rect.top) / rect.height) * 100;

    const fieldDef = ALL_FIELDS.find(f => f.name === selectedFieldName);
    const fontSize = parseInt(document.getElementById('field-fontsize').value) * 2 || 22; // half-points
    const bold = document.getElementById('field-bold').checked;

    // Remove existing placement of this field
    placedFields = placedFields.filter(p => p.name !== selectedFieldName);
    placedFields.push({
        name: selectedFieldName,
        label: fieldDef ? fieldDef.label : selectedFieldName,
        x: parseFloat(xPct.toFixed(2)),
        y: parseFloat(yPct.toFixed(2)),
        width: 40,
        height: 6,
        fontSize,
        bold,
        align: 'left'
    });

    selectedFieldName = null;
    container.classList.remove('placing');
    renderFieldBtns();
    renderFieldMarkers();
    renderPlacedList();
}

function renderFieldMarkers() {
    const container = document.getElementById('field-markers');
    const img = document.getElementById('tpl-preview-img');
    const imgW = img.offsetWidth || img.naturalWidth;
    const imgH = img.offsetHeight || img.naturalHeight;

    container.innerHTML = placedFields.map((f, i) => {
        const left = (f.x / 100) * imgW;
        const top = (f.y / 100) * imgH;
        return `<div class="field-marker" style="left:${left}px;top:${top}px" 
      title="${f.label}" data-idx="${i}">${f.label}</div>`;
    }).join('');
}

document.getElementById('tpl-preview-img').addEventListener('load', function () {
    renderFieldMarkers();
});

function renderPlacedList() {
    const el = document.getElementById('placed-fields-list');
    if (!placedFields.length) { el.innerHTML = '<div style="color:var(--text-m);font-size:.78rem">Chưa đặt trường nào.</div>'; return; }
    el.innerHTML = placedFields.map((f, i) =>
        `<div class="placed-field-item">
      <span>${f.label}</span>
      <button onclick="removeField(${i})" title="Xoá">✕</button>
    </div>`
    ).join('');
}

function removeField(idx) {
    placedFields.splice(idx, 1);
    renderFieldBtns();
    renderFieldMarkers();
    renderPlacedList();
}

async function saveFieldLayout() {
    if (!currentModalTemplateId) return;
    try {
        const res = await fetch(`/api/templates/${currentModalTemplateId}/layout`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ fields: placedFields })
        });
        if (!res.ok) throw new Error((await res.json()).error);
        await loadTemplates();
        closeFieldModal();
        alert(`✅ Đã lưu vị trí ${placedFields.length} trường!`);
    } catch (e) {
        alert('❌ ' + e.message);
    }
}

// ==================== GENERATE ====================
function setupExcelDrop() {
    const zone = document.getElementById('excel-zone');
    zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
    zone.addEventListener('drop', e => {
        e.preventDefault(); zone.classList.remove('drag-over');
        const f = e.dataTransfer.files[0];
        if (f) setExcelFile(f);
    });
    document.getElementById('excel-input').addEventListener('change', e => {
        if (e.target.files[0]) setExcelFile(e.target.files[0]);
    });
}

function setExcelFile(file) {
    selectedExcelFile = file;
    const lbl = document.getElementById('excel-filename');
    lbl.textContent = '✅ ' + file.name;
    lbl.classList.add('has-file');
    document.getElementById('excel-zone').classList.add('has-file');
    checkGenerateReady();
}

function checkGenerateReady() {
    const tplId = document.getElementById('gen-template-select').value;
    document.getElementById('btn-generate').disabled = !(selectedExcelFile && tplId);
}

function showGenCard(id) {
    ['gen-upload-card', 'gen-progress-card', 'gen-success-card', 'gen-error-card']
        .forEach(cid => document.getElementById(cid).classList.add('hidden'));
    document.getElementById(id).classList.remove('hidden');
}

async function generateCertificates() {
    const tplId = document.getElementById('gen-template-select').value;
    if (!selectedExcelFile || !tplId) return;

    showGenCard('gen-progress-card');
    setProgress(5);
    document.getElementById('progress-msg').textContent = 'Đang tải dữ liệu...';
    animateProgress();

    const fd = new FormData();
    fd.append('excel', selectedExcelFile);
    fd.append('templateId', tplId);

    try {
        setProgress(20);
        document.getElementById('progress-msg').textContent = 'Đang xử lý chứng chỉ...';

        const res = await fetch('/api/generate', { method: 'POST', body: fd });
        setProgress(85);

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.error || `HTTP ${res.status}`);
        }

        downloadBlob = await res.blob();
        const count = res.headers.get('X-Student-Count') || '?';
        setProgress(100);

        setTimeout(() => {
            showGenCard('gen-success-card');
            document.getElementById('success-msg').textContent = `Đã tạo xong ${count} chứng chỉ trong 1 file Word!`;
        }, 300);

    } catch (err) {
        showGenCard('gen-error-card');
        document.getElementById('error-msg').textContent = err.message;
    }
}

let progressTimer = null;
function animateProgress() {
    if (progressTimer) clearInterval(progressTimer);
    let pct = 20;
    progressTimer = setInterval(() => {
        if (pct < 80) { pct += Math.random() * 4; setProgress(Math.min(80, Math.round(pct))); }
        else clearInterval(progressTimer);
    }, 400);
}
function setProgress(pct) {
    document.getElementById('progress-fill').style.width = pct + '%';
    document.getElementById('progress-pct').textContent = pct + '%';
}

function downloadFile() {
    if (!downloadBlob) return;
    const url = URL.createObjectURL(downloadBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Chung-Chi-${new Date().toISOString().slice(0, 10)}.docx`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function resetGenerate() {
    selectedExcelFile = null;
    downloadBlob = null;
    document.getElementById('excel-input').value = '';
    const lbl = document.getElementById('excel-filename');
    lbl.textContent = 'Chưa chọn file'; lbl.classList.remove('has-file');
    document.getElementById('excel-zone').classList.remove('has-file');
    showGenCard('gen-upload-card');
    checkGenerateReady();
}

// ==================== UTILS ====================
function esc(str) {
    return String(str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
