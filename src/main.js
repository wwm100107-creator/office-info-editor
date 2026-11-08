import './style.css';
import JSZip from 'jszip';
import { loadOffice, buildOffice, familyOf, FIELDS, GROUPS, PERSONAL_KEYS, nowLocal, zipDate, isDirty, changedFields, customChanged } from './ooxml.js';

const $ = (s) => document.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
const fmtSize = (b) => (b < 1024 ? `${b} B` : b < 1048576 ? `${(b / 1024).toFixed(1)} KB` : `${(b / 1048576).toFixed(1)} MB`);
const FAMILY_NAME = { word: 'Word', ppt: 'PowerPoint', excel: 'Excel' };
const CUSTOM_TYPES = [['text', 'Văn bản'], ['number', 'Số'], ['date', 'Ngày'], ['bool', 'Có/Không']];

const state = { files: [], sel: null };
let uid = 0;
const cur = () => state.files.find((f) => f.id === state.sel);
const clone = (o) => JSON.parse(JSON.stringify(o));

function toast(msg, type = 'info') {
  const el = document.createElement('div');
  el.className = `toast ${type}`;
  el.textContent = msg;
  $('#toasts').appendChild(el);
  setTimeout(() => el.remove(), 4500);
}

async function addFiles(list) {
  let first = null;
  for (const file of list) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (['doc', 'ppt', 'xls'].includes(ext)) {
      toast(`"${file.name}": định dạng cũ .${ext} không được hỗ trợ — hãy mở bằng Office và lưu lại thành .${ext}x`, 'error');
      continue;
    }
    if (!familyOf(ext)) {
      toast(`"${file.name}": không phải file Word/PowerPoint/Excel`, 'error');
      continue;
    }
    try {
      const doc = await loadOffice(file);
      const f = { id: ++uid, ...doc, values: clone(doc.original.values), custom: clone(doc.original.custom) };
      state.files.push(f);
      first ??= f.id;
    } catch (e) {
      toast(`Không đọc được "${file.name}": ${e.message}`, 'error');
    }
  }
  if (first) state.sel = first;
  render();
}

function render() {
  const has = state.files.length > 0;
  $('#workspace').hidden = !has;
  $('#drop').classList.toggle('compact', has);
  $('#btn-download-all').hidden = state.files.length < 2;
  $('#btn-apply-all').hidden = state.files.length < 2;
  renderList();
  renderForm();
}

function renderList() {
  $('#file-list').innerHTML = state.files.map((f) => `
    <li class="${f.id === state.sel ? 'active' : ''}" data-id="${f.id}">
      <span class="tag ${f.family}">${esc(f.ext)}</span>
      <span class="fname" title="${esc(f.name)}">${esc(f.name)}</span>
      ${isDirty(f) ? '<span class="dot" title="Đã chỉnh sửa"></span>' : ''}
      <button class="x" data-remove="${f.id}" title="Bỏ file này">×</button>
    </li>`).join('');
}

function inputHTML(fd, v) {
  const attrs = `id="f-${fd.key}" data-key="${fd.key}"`;
  if (fd.input === 'textarea') return `<textarea ${attrs} rows="3">${esc(v)}</textarea>`;
  if (fd.input === 'date') return `<input ${attrs} type="datetime-local" step="1" value="${esc(v)}" />`;
  if (fd.input === 'number') return `<input ${attrs} type="number" min="0" step="1" value="${esc(v)}" />`;
  return `<input ${attrs} type="text" value="${esc(v)}" />`;
}

function customValueHTML(p, i) {
  const a = `data-ci="${i}" data-cf="value"`;
  if (p.type === 'bool') {
    return `<select ${a}><option value="true" ${p.value === 'true' ? 'selected' : ''}>Có</option><option value="false" ${p.value !== 'true' ? 'selected' : ''}>Không</option></select>`;
  }
  if (p.type === 'date') return `<input ${a} type="datetime-local" step="1" value="${esc(p.value)}" />`;
  if (p.type === 'number') return `<input ${a} type="number" step="any" value="${esc(p.value)}" />`;
  return `<input ${a} type="text" value="${esc(p.value)}" />`;
}

function renderForm() {
  const f = cur();
  if (!f) return;
  $('#cur-name').textContent = f.name;
  $('#cur-meta').textContent = `${FAMILY_NAME[f.family]} · ${fmtSize(f.size)}`;

  const groups = GROUPS.map(([g, title, note]) => {
    const fields = FIELDS.filter((fd) => fd.group === g && (!fd.only || fd.only.includes(f.family)));
    if (!fields.length) return '';
    return `<fieldset class="card"><legend>${title}</legend>
      ${note ? `<p class="note">${note}</p>` : ''}
      <div class="grid">${fields.map((fd) => {
        const v = f.values[fd.key] ?? '';
        const changed = v !== (f.original.values[fd.key] ?? '');
        return `<div class="field ${fd.input === 'textarea' ? 'wide' : ''} ${changed ? 'changed' : ''}">
          <div class="lbl-row"><label for="f-${fd.key}">${esc(fd.label)} <small>${esc(fd.hint)}</small></label>
          ${fd.input === 'date' ? `<button type="button" class="link" data-now="${fd.key}">Bây giờ</button>` : ''}</div>${inputHTML(fd, v)}</div>`;
      }).join('')}</div></fieldset>`;
  }).join('');

  const custom = `<fieldset class="card ${customChanged(f) ? 'changed' : ''}"><legend>Thuộc tính tùy chỉnh <small>Custom properties</small></legend>
    ${f.custom.length ? `<div class="custom-head"><span>Tên</span><span>Kiểu</span><span>Giá trị</span><span></span></div>` : '<p class="note">Chưa có thuộc tính tùy chỉnh nào.</p>'}
    ${f.custom.map((p, i) => `<div class="custom-row">
      <input data-ci="${i}" data-cf="name" type="text" placeholder="Tên thuộc tính" value="${esc(p.name)}" />
      <select data-ci="${i}" data-cf="type">${CUSTOM_TYPES.map(([t, n]) => `<option value="${t}" ${p.type === t ? 'selected' : ''}>${n}</option>`).join('')}</select>
      ${customValueHTML(p, i)}
      <button type="button" class="x" data-cdel="${i}" title="Xóa thuộc tính">×</button>
    </div>`).join('')}
    <button type="button" class="btn small" id="btn-add-custom">+ Thêm thuộc tính</button>
  </fieldset>`;

  const isFinal = f.custom.some((p) => p.name === '_MarkAsFinal' && p.value === 'true');
  $('#btn-final').classList.toggle('primary', isFinal);
  $('#btn-final').textContent = isFinal ? '🔒 Đã khóa Mark as Final' : '🔒 Khóa chỉnh sửa (Mark as Final)';

  $('#form').innerHTML = groups + custom;
}

function refreshDirty(el) {
  const f = cur();
  const key = el?.dataset.key;
  if (key) el.closest('.field').classList.toggle('changed', (f.values[key] ?? '') !== (f.original.values[key] ?? ''));
  renderList();
}

function saveBlob(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 10000);
}

// ---- Events ----
$('#file-input').addEventListener('change', (e) => {
  addFiles([...e.target.files]);
  e.target.value = '';
});
['dragover', 'drop'].forEach((t) => window.addEventListener(t, (e) => {
  e.preventDefault();
  if (t === 'drop') addFiles([...e.dataTransfer.files]);
}));
window.addEventListener('dragenter', () => $('#drop').classList.add('over'));
window.addEventListener('dragleave', (e) => !e.relatedTarget && $('#drop').classList.remove('over'));
window.addEventListener('drop', () => $('#drop').classList.remove('over'));

$('#file-list').addEventListener('click', (e) => {
  const rm = e.target.dataset.remove;
  if (rm) {
    state.files = state.files.filter((f) => f.id !== +rm);
    if (state.sel === +rm) state.sel = state.files[0]?.id ?? null;
  } else {
    const li = e.target.closest('li');
    if (li) state.sel = +li.dataset.id;
  }
  render();
});

$('#btn-clear').addEventListener('click', () => {
  state.files = [];
  state.sel = null;
  render();
});

const onFormInput = (e) => {
  const f = cur();
  const t = e.target;
  if (t.dataset.key) {
    f.values[t.dataset.key] = t.value;
    refreshDirty(t);
  } else if (t.dataset.cf) {
    f.custom[+t.dataset.ci][t.dataset.cf] = t.value;
    if (t.dataset.cf === 'type') {
      const p = f.custom[+t.dataset.ci];
      if (p.type !== 'text') p.value = { number: '0', bool: 'true', date: nowLocal() }[p.type];
      renderForm();
    } else {
      t.closest('fieldset').classList.toggle('changed', customChanged(f));
    }
    renderList();
  }
};
$('#form').addEventListener('input', onFormInput);
$('#form').addEventListener('change', onFormInput);

$('#form').addEventListener('click', (e) => {
  const f = cur();
  const t = e.target;
  if (t.dataset.now) {
    const input = $(`#f-${t.dataset.now}`);
    input.value = f.values[t.dataset.now] = nowLocal();
    refreshDirty(input);
  } else if (t.dataset.cdel) {
    f.custom.splice(+t.dataset.cdel, 1);
    renderForm();
    renderList();
  } else if (t.id === 'btn-add-custom') {
    f.custom.push({ name: '', type: 'text', value: '' });
    renderForm();
    document.querySelector(`[data-ci="${f.custom.length - 1}"][data-cf="name"]`).focus();
  }
});

$('#btn-final').addEventListener('click', () => {
  const f = cur();
  const idx = f.custom.findIndex((p) => p.name === '_MarkAsFinal');
  if (idx >= 0 && f.custom[idx].value === 'true') {
    f.custom.splice(idx, 1);
    toast('Đã tắt chế độ khóa Mark as Final');
  } else {
    if (idx >= 0) f.custom[idx] = { name: '_MarkAsFinal', type: 'bool', value: 'true' };
    else f.custom.push({ name: '_MarkAsFinal', type: 'bool', value: 'true' });
    toast('Đã bật Mark as Final — khi mở bằng Office sẽ ở chế độ Chỉ đọc (Read-Only), tránh bị lưu đè mất thông tin', 'success');
  }
  render();
});

$('#btn-personal').addEventListener('click', () => {
  const f = cur();
  PERSONAL_KEYS.forEach((k) => (f.values[k] = ''));
  render();
  toast('Đã xóa Tác giả, Người sửa cuối, Công ty, Quản lý, Ngày in — bấm "Tải file này" để lưu');
});

$('#btn-reset').addEventListener('click', () => {
  const f = cur();
  f.values = clone(f.original.values);
  f.custom = clone(f.original.custom);
  render();
});

$('#btn-apply-all').addEventListener('click', () => {
  const src = cur();
  const keys = changedFields(src);
  const withCustom = customChanged(src);
  if (!keys.length && !withCustom) return toast('File này chưa có thay đổi nào để áp dụng', 'error');
  const removed = src.original.custom.filter((o) => !src.custom.some((p) => p.name === o.name)).map((o) => o.name);
  const others = state.files.filter((f) => f !== src);
  for (const f of others) {
    for (const fd of keys) if (!fd.only || fd.only.includes(f.family)) f.values[fd.key] = src.values[fd.key];
    if (withCustom) {
      f.custom = f.custom.filter((p) => !removed.includes(p.name));
      for (const p of src.custom.filter((p) => p.name.trim())) {
        const i = f.custom.findIndex((q) => q.name === p.name);
        if (i >= 0) f.custom[i] = { ...p };
        else f.custom.push({ ...p });
      }
    }
  }
  renderList();
  toast(`Đã áp dụng ${keys.length + (withCustom ? 1 : 0)} thay đổi cho ${others.length} file khác`, 'success');
});

$('#btn-download').addEventListener('click', async () => {
  const f = cur();
  try {
    saveBlob(await buildOffice(f), f.name);
  } catch (e) {
    toast(`Lỗi khi tạo file: ${e.message}`, 'error');
  }
});

$('#btn-download-zip').addEventListener('click', async () => {
  const f = cur();
  try {
    const zip = new JSZip();
    zip.file(f.name, await buildOffice(f), { date: zipDate(f) });
    saveBlob(await zip.generateAsync({ type: 'blob' }), f.name.replace(/\.[^.]+$/, '') + '.zip');
  } catch (e) {
    toast(`Lỗi khi tạo file: ${e.message}`, 'error');
  }
});

$('#btn-download-all').addEventListener('click', async () => {
  try {
    const zip = new JSZip();
    const used = new Set();
    for (const f of state.files) {
      let name = f.name;
      for (let i = 1; used.has(name.toLowerCase()); i++) name = f.name.replace(/(\.[^.]+)$/, ` (${i})$1`);
      used.add(name.toLowerCase());
      zip.file(name, await buildOffice(f), { date: zipDate(f) });
    }
    saveBlob(await zip.generateAsync({ type: 'blob' }), 'office-info-edited.zip');
  } catch (e) {
    toast(`Lỗi khi tạo file: ${e.message}`, 'error');
  }
});
