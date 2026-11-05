import JSZip from 'jszip';

const XMLNS = 'http://www.w3.org/2000/xmlns/';
const NS = {
  cp: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
  dc: 'http://purl.org/dc/elements/1.1/',
  dcterms: 'http://purl.org/dc/terms/',
  dcmitype: 'http://purl.org/dc/dcmitype/',
  xsi: 'http://www.w3.org/2001/XMLSchema-instance',
  ep: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
  vt: 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
  custom: 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
};

const PARTS = {
  core: {
    def: 'docProps/core.xml',
    rel: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
    ct: 'application/vnd.openxmlformats-package.core-properties+xml',
    skeleton: `<cp:coreProperties xmlns:cp="${NS.cp}" xmlns:dc="${NS.dc}" xmlns:dcterms="${NS.dcterms}" xmlns:dcmitype="${NS.dcmitype}" xmlns:xsi="${NS.xsi}"/>`,
  },
  app: {
    def: 'docProps/app.xml',
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
    ct: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
    skeleton: `<Properties xmlns="${NS.ep}" xmlns:vt="${NS.vt}"/>`,
  },
  custom: {
    def: 'docProps/custom.xml',
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties',
    ct: 'application/vnd.openxmlformats-officedocument.custom-properties+xml',
  },
};

const EXT_FAMILY = {
  docx: 'word', docm: 'word', dotx: 'word', dotm: 'word',
  pptx: 'ppt', pptm: 'ppt', potx: 'ppt', potm: 'ppt', ppsx: 'ppt', ppsm: 'ppt',
  xlsx: 'excel', xlsm: 'excel', xltx: 'excel', xltm: 'excel',
};
export const familyOf = (ext) => EXT_FAMILY[ext];

const MIME = {
  word: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  ppt: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  excel: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

export const GROUPS = [
  ['props', 'Thuộc tính'],
  ['people', 'Người liên quan'],
  ['dates', 'Ngày tháng & thời gian'],
  ['stats', 'Thống kê', 'Word/PowerPoint có thể tự tính lại các số này khi file được lưu lại.'],
  ['advanced', 'Nâng cao'],
];

// part: core = docProps/core.xml, app = docProps/app.xml. only: giới hạn theo loại file.
export const FIELDS = [
  { key: 'title', part: 'core', tag: 'dc:title', label: 'Tiêu đề', hint: 'Title', group: 'props' },
  { key: 'subject', part: 'core', tag: 'dc:subject', label: 'Chủ đề', hint: 'Subject', group: 'props' },
  { key: 'keywords', part: 'core', tag: 'cp:keywords', label: 'Thẻ', hint: 'Tags', group: 'props' },
  { key: 'category', part: 'core', tag: 'cp:category', label: 'Thể loại', hint: 'Categories', group: 'props' },
  { key: 'contentStatus', part: 'core', tag: 'cp:contentStatus', label: 'Trạng thái', hint: 'Status', group: 'props' },
  { key: 'Company', part: 'app', tag: 'Company', label: 'Công ty', hint: 'Company', group: 'props' },
  { key: 'HyperlinkBase', part: 'app', tag: 'HyperlinkBase', label: 'Hyperlink base', hint: 'Hyperlink Base', group: 'props' },
  { key: 'Template', part: 'app', tag: 'Template', label: 'Mẫu', hint: 'Template', group: 'props' },
  { key: 'description', part: 'core', tag: 'dc:description', label: 'Chú thích', hint: 'Comments', group: 'props', input: 'textarea' },

  { key: 'creator', part: 'core', tag: 'dc:creator', label: 'Tác giả', hint: 'Author — nhiều người: cách nhau bằng ;', group: 'people' },
  { key: 'lastModifiedBy', part: 'core', tag: 'cp:lastModifiedBy', label: 'Người sửa cuối', hint: 'Last Modified By', group: 'people' },
  { key: 'Manager', part: 'app', tag: 'Manager', label: 'Quản lý', hint: 'Manager', group: 'people' },

  { key: 'created', part: 'core', tag: 'dcterms:created', label: 'Ngày tạo', hint: 'Created', group: 'dates', input: 'date', w3cdtf: true },
  { key: 'modified', part: 'core', tag: 'dcterms:modified', label: 'Sửa lần cuối', hint: 'Last Modified', group: 'dates', input: 'date', w3cdtf: true },
  { key: 'lastPrinted', part: 'core', tag: 'cp:lastPrinted', label: 'In lần cuối', hint: 'Last Printed', group: 'dates', input: 'date' },
  { key: 'TotalTime', part: 'app', tag: 'TotalTime', label: 'Tổng thời gian chỉnh sửa (phút)', hint: 'Total Editing Time', group: 'dates', input: 'number' },

  { key: 'Pages', part: 'app', tag: 'Pages', label: 'Số trang', hint: 'Pages', group: 'stats', input: 'number', only: ['word'] },
  { key: 'Slides', part: 'app', tag: 'Slides', label: 'Số slide', hint: 'Slides', group: 'stats', input: 'number', only: ['ppt'] },
  { key: 'HiddenSlides', part: 'app', tag: 'HiddenSlides', label: 'Slide ẩn', hint: 'Hidden Slides', group: 'stats', input: 'number', only: ['ppt'] },
  { key: 'Notes', part: 'app', tag: 'Notes', label: 'Số ghi chú', hint: 'Notes', group: 'stats', input: 'number', only: ['ppt'] },
  { key: 'MMClips', part: 'app', tag: 'MMClips', label: 'Media clip', hint: 'Multimedia Clips', group: 'stats', input: 'number', only: ['ppt'] },
  { key: 'Words', part: 'app', tag: 'Words', label: 'Số từ', hint: 'Words', group: 'stats', input: 'number', only: ['word', 'ppt'] },
  { key: 'Characters', part: 'app', tag: 'Characters', label: 'Số ký tự', hint: 'Characters', group: 'stats', input: 'number', only: ['word'] },
  { key: 'CharactersWithSpaces', part: 'app', tag: 'CharactersWithSpaces', label: 'Ký tự (kể cả khoảng trắng)', hint: 'Characters with spaces', group: 'stats', input: 'number', only: ['word'] },
  { key: 'Lines', part: 'app', tag: 'Lines', label: 'Số dòng', hint: 'Lines', group: 'stats', input: 'number', only: ['word'] },
  { key: 'Paragraphs', part: 'app', tag: 'Paragraphs', label: 'Số đoạn', hint: 'Paragraphs', group: 'stats', input: 'number', only: ['word', 'ppt'] },
  { key: 'PresentationFormat', part: 'app', tag: 'PresentationFormat', label: 'Khổ trình chiếu', hint: 'Presentation Format', group: 'stats', only: ['ppt'] },

  { key: 'revision', part: 'core', tag: 'cp:revision', label: 'Số lần sửa', hint: 'Revision Number', group: 'advanced', input: 'number' },
  { key: 'Application', part: 'app', tag: 'Application', label: 'Ứng dụng', hint: 'Application', group: 'advanced' },
  { key: 'AppVersion', part: 'app', tag: 'AppVersion', label: 'Phiên bản ứng dụng', hint: 'AppVersion, vd 16.0000', group: 'advanced' },
  { key: 'language', part: 'core', tag: 'dc:language', label: 'Ngôn ngữ', hint: 'Language, vd vi-VN', group: 'advanced' },
  { key: 'identifier', part: 'core', tag: 'dc:identifier', label: 'Mã định danh', hint: 'Identifier', group: 'advanced' },
  { key: 'version', part: 'core', tag: 'cp:version', label: 'Phiên bản', hint: 'Version', group: 'advanced' },
];

export const PERSONAL_KEYS = ['creator', 'lastModifiedBy', 'Manager', 'Company', 'lastPrinted'];

const VT_TYPE = {
  lpwstr: 'text', lpstr: 'text', bstr: 'text', bool: 'bool', filetime: 'date', date: 'date',
  i1: 'number', i2: 'number', i4: 'number', i8: 'number', int: 'number',
  ui1: 'number', ui2: 'number', ui4: 'number', ui8: 'number', uint: 'number',
  r4: 'number', r8: 'number', decimal: 'number', cy: 'number',
};

const pad = (n) => String(n).padStart(2, '0');
export function toLocal(iso) {
  const d = iso ? new Date(iso) : null;
  if (!d || isNaN(d)) return '';
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}
const toIso = (local) => {
  const d = local ? new Date(local) : null;
  return d && !isNaN(d) ? d.toISOString().replace(/\.\d{3}Z$/, 'Z') : '';
};
export const nowLocal = () => toLocal(new Date().toISOString());

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const escXml = (s) => String(s).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
  .replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[c]);
function parse(text) {
  const d = new DOMParser().parseFromString(text, 'application/xml');
  if (d.getElementsByTagName('parsererror').length) throw new Error('XML bên trong file bị lỗi');
  return d;
}
const serialize = (d) => XML_DECL + new XMLSerializer().serializeToString(d.documentElement);
const childEl = (root, ns, local) => [...root.children].find((c) => c.namespaceURI === ns && c.localName === local);
const splitTag = (tag) => (tag.includes(':') ? tag.split(':') : [null, tag]);
const nsOf = (fd) => (fd.part === 'app' ? NS.ep : NS[splitTag(fd.tag)[0]]);

async function partPaths(zip) {
  const paths = { core: PARTS.core.def, app: PARTS.app.def, custom: PARTS.custom.def };
  const rels = zip.file('_rels/.rels');
  if (rels) {
    for (const r of parse(await rels.async('string')).getElementsByTagNameNS(NS.rel, 'Relationship')) {
      for (const [k, p] of Object.entries(PARTS)) {
        if (r.getAttribute('Type') === p.rel) paths[k] = r.getAttribute('Target').replace(/^\//, '');
      }
    }
  }
  return paths;
}

export async function loadOffice(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  const buffer = await file.arrayBuffer();
  let zip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch {
    throw new Error('file bị hỏng hoặc đang được đặt mật khẩu');
  }
  if (!zip.file('[Content_Types].xml')) throw new Error('không phải file Office Open XML');

  const paths = await partPaths(zip);
  const xml = {};
  for (const k of Object.keys(PARTS)) {
    const z = zip.file(paths[k]);
    if (z) xml[k] = parse(await z.async('string'));
  }

  const values = {};
  for (const fd of FIELDS) {
    const root = xml[fd.part]?.documentElement;
    const el = root && childEl(root, nsOf(fd), splitTag(fd.tag)[1]);
    const v = el ? el.textContent.trim() : '';
    values[fd.key] = fd.input === 'date' ? toLocal(v) : v;
  }

  const custom = [];
  for (const p of xml.custom?.documentElement.children ?? []) {
    if (p.localName !== 'property') continue;
    const vEl = p.firstElementChild;
    const type = VT_TYPE[vEl?.localName] || 'text';
    let value = vEl?.textContent ?? '';
    if (type === 'date') value = toLocal(value);
    if (type === 'bool') value = /^(true|1)$/i.test(value.trim()) ? 'true' : 'false';
    custom.push({ name: p.getAttribute('name') || '', type, value });
  }

  return { name: file.name, ext, family: familyOf(ext), size: file.size, buffer, original: { values, custom } };
}

export const customChanged = (f) => JSON.stringify(f.custom) !== JSON.stringify(f.original.custom);
export const changedFields = (f) => FIELDS.filter((fd) => (f.values[fd.key] ?? '') !== (f.original.values[fd.key] ?? ''));
export const isDirty = (f) => customChanged(f) || changedFields(f).length > 0;

function setField(d, fd, raw) {
  const root = d.documentElement;
  const [prefix, local] = splitTag(fd.tag);
  const ns = nsOf(fd);
  let value = String(raw ?? '').trim();
  if (fd.input === 'date') value = toIso(value);
  if (fd.input === 'number' && value) value = String(Math.max(0, Math.round(Number(value)) || 0));

  let el = childEl(root, ns, local);
  if (!value) return el?.remove();
  if (prefix && root.lookupNamespaceURI(prefix) !== ns) root.setAttributeNS(XMLNS, `xmlns:${prefix}`, ns);
  if (!el) el = root.appendChild(d.createElementNS(ns, fd.tag));
  el.textContent = value;
  if (fd.w3cdtf) {
    if (root.lookupNamespaceURI('xsi') !== NS.xsi) root.setAttributeNS(XMLNS, 'xmlns:xsi', NS.xsi);
    el.setAttributeNS(NS.xsi, 'xsi:type', 'dcterms:W3CDTF');
  }
}

function encodeCustom(p) {
  if (p.type === 'number') {
    const n = Number(p.value) || 0;
    return [Number.isInteger(n) && Math.abs(n) < 2 ** 31 ? 'i4' : 'r8', String(n)];
  }
  if (p.type === 'bool') return ['bool', p.value === 'true' ? 'true' : 'false'];
  if (p.type === 'date') return ['filetime', toIso(p.value) || toIso(nowLocal())];
  return ['lpwstr', p.value ?? ''];
}

function customXml(list) {
  const seen = new Set();
  const props = list.filter((p) => {
    const n = p.name.trim();
    return n && !seen.has(n) && seen.add(n);
  }).map((p, i) => {
    const [vt, text] = encodeCustom(p);
    return `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${i + 2}" name="${escXml(p.name.trim())}"><vt:${vt}>${escXml(text)}</vt:${vt}></property>`;
  }).join('');
  return `${XML_DECL}<Properties xmlns="${NS.custom}" xmlns:vt="${NS.vt}">${props}</Properties>`;
}

async function registerPart(zip, part, path) {
  const ct = parse(await zip.file('[Content_Types].xml').async('string'));
  const pn = `/${path}`.toLowerCase();
  if (![...ct.documentElement.children].some((c) => c.getAttribute('PartName')?.toLowerCase() === pn)) {
    const o = ct.createElementNS(NS.ct, 'Override');
    o.setAttribute('PartName', `/${path}`);
    o.setAttribute('ContentType', PARTS[part].ct);
    ct.documentElement.appendChild(o);
    zip.file('[Content_Types].xml', serialize(ct));
  }

  const relsFile = zip.file('_rels/.rels');
  const rels = parse(relsFile ? await relsFile.async('string') : `<Relationships xmlns="${NS.rel}"/>`);
  const list = [...rels.documentElement.children];
  if (!list.some((c) => c.getAttribute('Type') === PARTS[part].rel)) {
    const ids = new Set(list.map((c) => c.getAttribute('Id')));
    let i = 1;
    while (ids.has(`rId${i}`)) i++;
    const r = rels.createElementNS(NS.rel, 'Relationship');
    r.setAttribute('Id', `rId${i}`);
    r.setAttribute('Type', PARTS[part].rel);
    r.setAttribute('Target', path);
    rels.documentElement.appendChild(r);
    zip.file('_rels/.rels', serialize(rels));
  }
}

// Chỉ ghi lại các trường đã thay đổi; mọi phần khác của file được giữ nguyên.
export async function buildOffice(f) {
  const zip = await JSZip.loadAsync(f.buffer);
  const paths = await partPaths(zip);
  const changed = changedFields(f);
  const parts = new Set(changed.map((fd) => fd.part));
  if (customChanged(f)) parts.add('custom');

  for (const part of parts) {
    const path = paths[part];
    const existing = zip.file(path);
    if (part === 'custom') {
      zip.file(path, customXml(f.custom));
    } else {
      const d = parse(existing ? await existing.async('string') : PARTS[part].skeleton);
      for (const fd of changed) if (fd.part === part) setField(d, fd, f.values[fd.key]);
      zip.file(path, serialize(d));
    }
    if (!existing) await registerPart(zip, part, path);
  }
  return zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 }, mimeType: MIME[f.family] });
}
