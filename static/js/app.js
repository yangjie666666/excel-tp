/* Excel 图片提取 · 前端逻辑 */

const state = { fileId: null, sheets: [], headers: [], imageCount: 0 };
const $ = (s) => document.querySelector(s);
const $$ = (s) => document.querySelectorAll(s);

/* ── 安全DOM ── */
function fillSelect(el, items, placeholder) {
    el.textContent = '';
    if (placeholder) {
        const o = document.createElement('option');
        o.value = '';
        o.textContent = placeholder;
        el.appendChild(o);
    }
    items.forEach(item => {
        const o = document.createElement('option');
        o.value = item.value;
        o.textContent = item.label;
        el.appendChild(o);
    });
}

/* ── 上传 ── */
const uploadArea = $('#upload-area');
const fileInput  = $('#file-input');

uploadArea.addEventListener('click', () => fileInput.click());
uploadArea.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); }
});
uploadArea.addEventListener('dragover', e => { e.preventDefault(); uploadArea.classList.add('dragover'); });
uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('dragover'));
uploadArea.addEventListener('drop', e => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    if (e.dataTransfer.files[0]) handleUpload(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', e => { if (e.target.files[0]) handleUpload(e.target.files[0]); });

async function handleUpload(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['xlsx', 'xls'].includes(ext)) { alert('仅支持 .xlsx 和 .xls 格式'); return; }

    const pEl = $('#upload-progress');
    const fEl = $('#upload-fill');
    const tEl = $('#upload-text');

    pEl.hidden = false;
    fEl.style.width = '30%';
    tEl.textContent = '上传中…';

    try {
        fEl.style.width = '60%';
        tEl.textContent = '解析中…';
        const form = new FormData();
        form.append('file', file);
        const resp = await fetch('/api/upload', { method: 'POST', body: form });
        const data = await resp.json();
        if (!resp.ok) throw new Error(data.error || '上传失败');

        fEl.style.width = '100%';
        tEl.textContent = '解析完成';

        state.fileId     = data.file_id;
        state.sheets     = data.sheets;
        state.headers    = data.headers;
        state.imageCount = data.image_count;

        renderConfig(data);
        goStep('config');
    } catch (err) {
        alert(err.message);
        pEl.hidden = true;
    }
}

/* ── 配置 ── */
function renderConfig(data) {
    const chip = $('#file-chip');
    const chipText = $('#file-chip-text');
    chipText.textContent = data.original_name + '  ·  ' + data.image_count + ' 张图片';
    chip.hidden = false;

    fillSelect($('#sheet-select'), data.sheets.map(s => ({ value: s, label: s })));
    renderColumns(data.headers);
    $('#image-info').textContent = '当前 Sheet 检测到 ' + data.image_count + ' 张浮动图片';
}

function renderColumns(headers) {
    const items = headers.map(h => ({ value: h.index, label: h.letter + ' · ' + h.name }));
    fillSelect($('#image-col-select'),  items, '选择图片列');
    fillSelect($('#name-col-select'),   items, '选择命名列');
    fillSelect($('#prefix-col-select'), items, '不使用');

    const imgM = headers.find(h => /图片|image|photo|照片/i.test(h.name));
    if (imgM) $('#image-col-select').value = imgM.index;
    const nameM = headers.find(h => /产品名称|商品名称|名称|name|品名/i.test(h.name));
    if (nameM) $('#name-col-select').value = nameM.index;

    refreshBtn();
    refreshPreview();
}

$('#sheet-select').addEventListener('change', async e => {
    const name = e.target.value;
    if (!name) return;
    try {
        const resp = await fetch('/api/sheet-info', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ file_id: state.fileId, sheet_name: name }),
        });
        const d = await resp.json();
        if (!resp.ok) throw new Error(d.error);
        state.headers = d.headers;
        state.imageCount = d.image_count;
        $('#image-info').textContent = '当前 Sheet 检测到 ' + d.image_count + ' 张浮动图片';
        renderColumns(d.headers);
    } catch (err) { alert(err.message); }
});

$('#image-col-select').addEventListener('change', () => { refreshBtn(); refreshPreview(); });
$('#name-col-select').addEventListener('change',  () => { refreshBtn(); refreshPreview(); });
$('#prefix-col-select').addEventListener('change', refreshPreview);
$('#manual-prefix').addEventListener('input', refreshPreview);

function refreshBtn() {
    $('#btn-extract').disabled = !($('#image-col-select').value && $('#name-col-select').value);
}

function refreshPreview() {
    const nIdx = parseInt($('#name-col-select').value);
    const pIdx = parseInt($('#prefix-col-select').value);
    const mpfx = $('#manual-prefix').value;
    const box  = $('#naming-preview');
    const code = $('#preview-example');

    if (!nIdx) { box.hidden = true; return; }

    const nH = state.headers.find(h => h.index === nIdx);
    const pH = pIdx ? state.headers.find(h => h.index === pIdx) : null;

    const parts = [];
    if (mpfx) parts.push(mpfx);
    if (pH)   parts.push('[' + pH.name + '的值]');
    parts.push('[' + nH.name + '的值]');

    code.textContent = parts.join('-') + '.png';
    box.hidden = false;
}

$('#btn-reupload').addEventListener('click', resetAll);

/* ── 提取 ── */
$('#btn-extract').addEventListener('click', async () => {
    const imageCol     = parseInt($('#image-col-select').value);
    const nameCol      = parseInt($('#name-col-select').value);
    const prefixCol    = $('#prefix-col-select').value ? parseInt($('#prefix-col-select').value) : null;
    const manualPrefix = $('#manual-prefix').value;

    goStep('result');

    const pEl = $('#extract-progress');
    const fEl = $('#extract-fill');
    const tEl = $('#extract-text');

    pEl.hidden = false;
    fEl.style.width = '50%';
    tEl.textContent = '提取中…';

    try {
        const resp = await fetch('/api/extract', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                file_id: state.fileId,
                sheet_name: $('#sheet-select').value,
                image_col: imageCol,
                name_col: nameCol,
                prefix_col: prefixCol,
                manual_prefix: manualPrefix,
            }),
        });
        const d = await resp.json();
        if (!resp.ok) throw new Error(d.error);

        fEl.style.width = '100%';
        tEl.textContent = '完成';
        showResult(d);
    } catch (err) {
        tEl.textContent = '失败';
        alert(err.message);
    }
});

function showResult(d) {
    $('#stat-extracted').textContent = d.extracted;
    $('#stat-skipped').textContent   = d.skipped;
    $('#stat-total').textContent     = d.total;
    $('#result-stats').hidden = false;

    if (d.errors && d.errors.length) {
        const el = $('#error-list');
        el.textContent = '';
        d.errors.forEach(msg => {
            const p = document.createElement('p');
            p.textContent = msg;
            el.appendChild(p);
        });
        el.hidden = false;
    }

    if (d.download_url) {
        $('#btn-download').href = d.download_url;
        $('#download-actions').hidden = false;
    }
}

$('#btn-restart').addEventListener('click', resetAll);

/* ── 步骤导航 ── */
function goStep(step) {
    const order = ['upload', 'config', 'result'];
    const idx   = order.indexOf(step);

    // 解锁配置面板
    if (idx >= 1) {
        $('#sec-config').classList.remove('disabled');
    }

    // 显示结果面板
    if (idx >= 2) {
        $('#sec-result').classList.remove('hidden');
    }

    // 平滑滚动到对应面板
    const target = $('#sec-' + step);
    if (target) {
        setTimeout(() => {
            target.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }, 50);
    }
}

function resetAll() {
    state.fileId = null;
    state.sheets = [];
    state.headers = [];
    state.imageCount = 0;

    fileInput.value = '';
    $('#upload-progress').hidden = true;
    $('#upload-fill').style.width = '0%';
    $('#result-stats').hidden = true;
    $('#error-list').hidden = true;
    $('#download-actions').hidden = true;
    $('#extract-progress').hidden = true;
    $('#file-chip').hidden = true;

    // 重置面板状态
    $('#sec-config').classList.add('disabled');
    $('#sec-result').classList.add('hidden');

    // 滚动回顶部
    window.scrollTo({ top: 0, behavior: 'smooth' });
}
