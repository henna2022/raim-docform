// 올린 파일을 공통 문서 구조로 읽는다.
// doc = { title, subtitle, blocks, warnings }
// block: {t:'h', level, text} | {t:'p', runs} | {t:'li', runs, level, kind:'bullet'|'num'|'check', list}
//        | {t:'note', runs} | {t:'table', rows} | {t:'img', src, w, h}
// run: {text, b?, href?}

const MAX_IMG = 1600; // 긴 변 픽셀. 원본 그대로 넣으면 수백 MB가 된다

async function readDocument(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  const buf = await file.arrayBuffer();
  const doc = { title: '', subtitle: '', blocks: [], warnings: [], source: ext };
  if (ext === 'docx') await readDocx(buf, doc);
  else if (ext === 'pptx') await readPptx(buf, doc);
  else if (ext === 'hwpx') await readHwpx(buf, doc);
  else if (ext === 'hwp') await readHwp(buf, doc);
  else if (ext === 'doc' || ext === 'ppt') throw new Error('예전 형식(.doc/.ppt)은 Word나 PowerPoint에서 docx/pptx로 다시 저장한 뒤 올려주세요.');
  else throw new Error('docx, pptx, hwp, hwpx 파일만 올릴 수 있어요.');
  absorbCover(doc);
  dropSourceToc(doc);
  inferHeadings(doc);
  detectNotes(doc);
  if (!doc.title) doc.title = file.name.replace(/\.[^.]+$/, '');
  if (!doc.blocks.length) doc.warnings.push('문서에서 내용을 찾지 못했어요.');
  return doc;
}

// ---------- 공통 ----------

const xml = s => new DOMParser().parseFromString(s, 'application/xml');
const kids = (el, name) => [...el.children].filter(c => c.localName === name);
const kid = (el, name) => kids(el, name)[0];
const desc = (el, name) => [...el.getElementsByTagNameNS('*', name)];
const runsText = runs => runs.map(r => r.text).join('');

function pushText(doc, text, extra = {}) {
  text = text.trim();
  if (text.trim()) doc.blocks.push({ t: 'p', runs: [{ text }], ...extra });
}

// 그림을 한 번 캔버스에 그려 회전·자르기·축소·EXIF 방향을 굽는다.
// crop: {l,t,r,b} 0~1 비율, rot: 시계 방향 도
async function normImage(blob, { rot = 0, crop } = {}) {
  let bmp;
  try { bmp = await createImageBitmap(blob, { imageOrientation: 'from-image' }); }
  catch { return null; } // EMF·WMF·TIFF 등 브라우저가 못 여는 그림
  let sx = 0, sy = 0, sw = bmp.width, sh = bmp.height;
  if (crop) {
    sx = bmp.width * crop.l; sy = bmp.height * crop.t;
    sw = bmp.width * (1 - crop.l - crop.r); sh = bmp.height * (1 - crop.t - crop.b);
  }
  const quarter = ((Math.round(rot / 90) % 4) + 4) % 4;
  const k = Math.min(1, MAX_IMG / Math.max(sw, sh));
  const dw = Math.round(sw * k), dh = Math.round(sh * k);
  const [cw, ch] = quarter % 2 ? [dh, dw] : [dw, dh];
  const c = document.createElement('canvas');
  c.width = cw; c.height = ch;
  const g = c.getContext('2d');
  g.translate(cw / 2, ch / 2);
  g.rotate(quarter * Math.PI / 2);
  g.drawImage(bmp, sx, sy, sw, sh, -dw / 2, -dh / 2, dw, dh);
  const png = blob.type === 'image/png' || blob.type === 'image/gif';
  const src = c.toDataURL(png ? 'image/png' : 'image/jpeg', 0.88);
  return { t: 'img', src, w: cw, h: ch };
}

const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml' };
const mimeOf = path => MIME[path.split('.').pop().toLowerCase()] || 'application/octet-stream';

async function addImage(doc, bytes, path, opts) {
  const img = await normImage(new Blob([bytes], { type: mimeOf(path) }), opts);
  if (img) doc.blocks.push(img);
  else doc._skippedImages = (doc._skippedImages || 0) + 1;
}

function flushSkipped(doc) {
  if (doc._skippedImages) doc.warnings.push(`열 수 없는 그림 ${doc._skippedImages}개(EMF·WMF 등)는 옮기지 못했어요.`);
  delete doc._skippedImages;
}

// 제목 스타일이 하나도 없으면 "1. 개요", "제2장", 굵은 짧은 줄을 제목으로 추정한다
function inferHeadings(doc) {
  if (doc.blocks.some(b => b.t === 'h')) return;
  let found = 0;
  doc.blocks = doc.blocks.map(b => {
    if (b.t !== 'p') return b;
    const text = runsText(b.runs).trim();
    if (text.length > 40 || /[.。]$|다$/.test(text)) return b;
    let level = 0;
    if (/^(chapter\s*\d+|제\s*\d+\s*[장편부]|[IVX]+\.|\d+\.)\s*\S/i.test(text)) level = 1;
    else if (/^\d+\.\d+\.?\s*\S/.test(text)) level = 2;
    else if (b.runs.every(r => r.b || !r.text.trim())) level = 2;
    if (!level) return b;
    found++;
    return { t: 'h', level, text };
  });
  if (found) doc.warnings.push('제목 스타일이 없어서 번호가 붙거나 굵은 짧은 줄을 제목으로 잡았어요. 결과를 한 번 확인해 주세요.');
}

// 이 양식으로 만든 문서를 다시 올리면 표지·개정 이력이 본문에 섞인다.
// 첫 제목 앞에 문서번호·버전 표가 있으면 표지로 보고, 값은 표지 정보로 옮긴다
const META_KEYS = { 문서번호: 'docNo', 버전: 'version', 작성부서: 'dept', 작성자: 'author', 승인자: 'approver' };
// "2026년 9월 8일" → date 입력칸 값
function setDate(meta, text) {
  const m = text.match(/(\d{4})\s*[년.-]\s*(\d{1,2})\s*[월.-]\s*(\d{1,2})/);
  if (m) meta.date = `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
}
function absorbCover(doc) {
  const first = doc.blocks.findIndex(b => b.t === 'h');
  const head = doc.blocks.slice(0, first < 0 ? 0 : first);
  const table = head.find(b => b.t === 'table' && b.rows.filter(r => /^(문서번호|버전|작성일)$/.test(r[0].trim())).length >= 2);
  if (!table) return;
  doc.meta = { extra: [] };
  for (const [k, v] of table.rows) {
    if (!v?.trim() || k.trim() === '항목') continue;
    if (k.trim() === '작성일') setDate(doc.meta, v);
    else if (META_KEYS[k.trim()]) doc.meta[META_KEYS[k.trim()]] = v.trim();
    else doc.meta.extra.push([k.trim(), v.trim()]);
  }
  const texts = head.filter(b => b.t === 'p').map(b => runsText(b.runs).trim()).filter(t => !/^(서울로봇인공지능과학관|개정\s*이력|목\s*차|차\s*례)$/.test(t) && !/필드 업데이트/.test(t));
  if (!doc.title && texts[0]) doc.title = texts[0];
  if (!doc.subtitle && texts[1] && texts[1].length <= 80) doc.subtitle = texts[1];
  doc.blocks.splice(0, first);
}

// "목차" 제목과 그 아래 짧은 줄들은 원본 목차라 뺀다
function dropSourceToc(doc) {
  const i = doc.blocks.findIndex(b => b.t === 'h' && /^(목\s*차|차\s*례|contents|table of contents)$/i.test(b.text.trim()));
  if (i < 0) return;
  let j = i + 1;
  while (j < doc.blocks.length && (doc.blocks[j].t === 'p' || doc.blocks[j].t === 'li') && runsText(doc.blocks[j].runs).length < 80) j++;
  doc.blocks.splice(i, j - i);
}

// "주의", "※" 로 시작하는 문단은 주의 표시로
function detectNotes(doc) {
  for (const b of doc.blocks) {
    if (b.t === 'p' && /^\s*(※|주의|유의|경고)/.test(runsText(b.runs))) b.t = 'note';
  }
}

// ---------- docx ----------

async function readDocx(buf, doc) {
  const styleMap = [
    "p[style-name='Title'] => p.doc-title:fresh",
    "p[style-name='제목'] => p.doc-title:fresh",
    "p[style-name='Subtitle'] => p.doc-subtitle:fresh",
    "p[style-name='부제'] => p.doc-subtitle:fresh",
    "p[style-name='제목 1'] => h1:fresh",
    "p[style-name='제목 2'] => h2:fresh",
    "p[style-name='제목 3'] => h3:fresh",
    // 원본 목차는 버린다. 양식이 목차를 새로 만든다
    "p[style-name='toc 1'] => !", "p[style-name='toc 2'] => !", "p[style-name='toc 3'] => !",
    "p[style-name='TOC Heading'] => !", "p[style-name='목차 1'] => !", "p[style-name='목차 2'] => !",
  ];
  const images = new Map();
  const convertImage = mammoth.images.imgElement(async img => {
    const key = 'img' + images.size;
    images.set(key, await normImage(new Blob([await img.readAsArrayBuffer()], { type: img.contentType })));
    return { src: key };
  });
  const { value } = await mammoth.convertToHtml({ arrayBuffer: buf }, { styleMap, convertImage, ignoreEmptyParagraphs: true });
  const body = new DOMParser().parseFromString(value, 'text/html').body;

  const inline = (node, style = {}, out = []) => {
    for (const n of node.childNodes) {
      if (n.nodeType === 3) { if (n.textContent) out.push({ text: n.textContent, ...style }); continue; }
      const tag = n.localName;
      if (tag === 'br') out.push({ text: '\n', ...style });
      else if (tag === 'img') out.push({ img: n.getAttribute('src') });
      else inline(n, { ...style, ...(tag === 'strong' || tag === 'b' ? { b: true } : {}), ...(tag === 'a' && /^https?:|^mailto:/.test(n.getAttribute('href') || '') ? { href: n.getAttribute('href') } : {}) }, out);
    }
    return out;
  };
  // 문단 안에 그림이 섞여 있으면 글과 그림을 순서대로 나눈다
  const pushRuns = (all, make) => {
    let cur = [];
    const flush = () => { if (runsText(cur).trim()) doc.blocks.push(make(cur)); cur = []; };
    for (const r of all) {
      if (r.img) { flush(); const im = images.get(r.img); if (im) doc.blocks.push(im); else doc._skippedImages = (doc._skippedImages || 0) + 1; }
      else cur.push(r);
    }
    flush();
  };
  let listNo = 0;
  const walkList = (el, level) => {
    const list = ++listNo;
    for (const li of kids(el, 'li')) {
      const nested = [...li.children].filter(c => c.localName === 'ul' || c.localName === 'ol');
      nested.forEach(n => n.remove());
      pushRuns(inline(li), runs => ({ t: 'li', runs, level: Math.min(level, 2), kind: el.localName === 'ol' ? 'num' : 'bullet', list }));
      nested.forEach(n => walkList(n, level + 1));
    }
  };
  for (const el of body.children) {
    const tag = el.localName;
    const text = el.textContent.trim();
    if (/^h[1-6]$/.test(tag)) { if (text) doc.blocks.push({ t: 'h', level: Math.min(+tag[1], 3), text }); }
    else if (tag === 'p' && el.className === 'doc-title') { if (!doc.title) doc.title = text; else doc.blocks.push({ t: 'h', level: 1, text }); }
    else if (tag === 'p' && el.className === 'doc-subtitle') { if (!doc.subtitle) doc.subtitle = text; }
    else if (tag === 'p') {
      const m = text.match(/^(☐|□|■|☑|✓|\[\s?\])\s*/);
      pushRuns(inline(el), runs => m ? { t: 'li', runs: stripLead(runs, m[0].length), level: 0, kind: 'check', list: 0 } : { t: 'p', runs });
    }
    else if (tag === 'ul' || tag === 'ol') walkList(el, 0);
    else if (tag === 'table') {
      const rows = desc(el, 'tr').filter(tr => tr.closest('table') === el)
        .map(tr => kids(tr, 'td').concat(kids(tr, 'th')).map(td => [...td.children].map(c => c.textContent.trim()).filter(Boolean).join('\n') || td.textContent.trim()));
      desc(el, 'img').forEach(img => { const im = images.get(img.getAttribute('src')); if (im) doc.blocks.push(im); });
      pushTable(doc, rows);
    }
  }
  flushSkipped(doc);
}

function stripLead(runs, n) {
  const out = runs.map(r => ({ ...r }));
  for (const r of out) {
    const lead = r.text.length - r.text.trimStart().length;
    const cut = Math.min(r.text.length, n + lead);
    r.text = r.text.slice(cut); n -= cut - lead;
    if (n <= 0) break;
  }
  return out.filter(r => r.text);
}

function pushTable(doc, rows) {
  rows = rows.filter(r => r.some(c => c.trim()));
  if (!rows.length) return;
  const cols = Math.max(...rows.map(r => r.length));
  if (cols === 1) { rows.forEach(r => pushText(doc, r[0])); return; } // 한 칸짜리 표는 글상자로 쓴 것
  doc.blocks.push({ t: 'table', rows: rows.map(r => [...r, ...Array(cols - r.length).fill('')]) });
}

// ---------- pptx ----------

async function readPptx(buf, doc) {
  const zip = await JSZip.loadAsync(buf);
  const read = p => zip.file(p)?.async('string');
  const rels = async path => {
    const relPath = path.replace(/([^/]+)$/, '_rels/$1.rels');
    const s = await read(relPath);
    const map = {};
    if (s) for (const r of desc(xml(s).documentElement, 'Relationship')) map[r.getAttribute('Id')] = new URL(r.getAttribute('Target'), 'http://x/' + path).pathname.slice(1);
    return map;
  };
  const pres = xml(await read('ppt/presentation.xml')).documentElement;
  const presRels = await rels('ppt/presentation.xml');
  const sz = desc(pres, 'sldSz')[0];
  const slideW = +sz.getAttribute('cx'), slideH = +sz.getAttribute('cy');
  // p14:sldId(구역 정보)도 이름이 같아서 r:id가 있는 것만 쓴다
  const slidePaths = desc(pres, 'sldId').map(s => presRels[s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')]).filter(p => p && zip.file(p));

  for (let si = 0; si < slidePaths.length; si++) {
    const path = slidePaths[si];
    const root = xml(await read(path)).documentElement;
    const rel = await rels(path);
    let title = '', subtitle = '';
    const items = []; // {y, x, kind, ...}
    const edges = []; // 위쪽 가장자리의 짧은 글(머리말). 제목이 없을 때 제목 후보

    const walk = (tree, dx = 0, dy = 0) => {
      for (const el of tree.children) {
        const off = desc(el, 'off')[0];
        const x = (off ? +off.getAttribute('x') : 0) + dx, y = (off ? +off.getAttribute('y') : 0) + dy;
        const ext = desc(el, 'ext').find(e => e.hasAttribute('cy'));
        const h = ext ? +ext.getAttribute('cy') : 0;
        if (el.localName === 'grpSp') { walk(el); continue; }
        const ph = desc(el, 'ph')[0]?.getAttribute('type') || '';
        if (el.localName === 'sp') {
          if (['sldNum', 'ftr', 'dt', 'hdr'].includes(ph)) continue;
          const body = kid(el, 'txBody');
          if (!body) continue;
          const paras = kids(body, 'p').map(p => {
            const runs = [];
            for (const r of p.children) {
              if (r.localName === 'r' || r.localName === 'fld') {
                const rPr = kid(r, 'rPr');
                const link = rPr && desc(rPr, 'hlinkClick')[0];
                const href = link && rel[link.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')];
                runs.push({ text: kid(r, 't')?.textContent || '', ...(rPr?.getAttribute('b') === '1' ? { b: true } : {}), ...(href && /^https?:/.test(href) ? { href } : {}) });
              } else if (r.localName === 'br') runs.push({ text: '\n' });
            }
            const pPr = kid(p, 'pPr');
            const bullet = pPr && (kid(pPr, 'buChar') || kid(pPr, 'buAutoNum'));
            return { runs, level: +(pPr?.getAttribute('lvl') || 0), bullet: bullet ? (bullet.localName === 'buAutoNum' ? 'num' : 'bullet') : '' };
          }).filter(p => runsText(p.runs).trim());
          if (!paras.length) continue;
          const text = paras.map(p => runsText(p.runs)).join('\n').trim();
          if (ph === 'title' || ph === 'ctrTitle') { title = title || text; continue; }
          if (ph === 'subTitle') { subtitle = subtitle || text; continue; }
          // 머리말·꼬리말처럼 위아래 가장자리에 붙은 짧은 글상자는 버린다
          if (si > 0 && text.length < 40 && (y + h < slideH * 0.13 || y > slideH * 0.9)) { // 첫 장(표지)은 맨 아랫줄까지 내용
            if (y < slideH * 0.13 && x < slideW * 0.5 && !/^(\d+|\d+\s*\/\s*\d+|step\b.*|p\.?\s*\d+)$/i.test(text)) edges.push(text);
            continue;
          }
          const size = Math.max(0, ...desc(el, 'rPr').map(r => +(r.getAttribute('sz') || 0))) || 1800;
          const bold = desc(el, 'rPr').some(r => r.getAttribute('b') === '1');
          items.push({ x, y, kind: 'text', paras, text, size, bold });
        } else if (el.localName === 'pic') {
          const blip = desc(el, 'blip')[0];
          const target = blip && rel[blip.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed')];
          if (!target || !zip.file(target)) continue;
          const xfrm = desc(el, 'xfrm')[0];
          const src = desc(el, 'srcRect')[0];
          const pct = a => (+(src?.getAttribute(a) || 0)) / 100000;
          items.push({ x, y, kind: 'pic', target, rot: +(xfrm?.getAttribute('rot') || 0) / 60000, crop: src ? { l: pct('l'), t: pct('t'), r: pct('r'), b: pct('b') } : null });
        } else if (el.localName === 'graphicFrame') {
          const tbl = desc(el, 'tbl')[0];
          if (tbl) items.push({ x, y, kind: 'table', rows: desc(tbl, 'tr').map(tr => kids(tr, 'tc').map(tc => kids(kid(tc, 'txBody') || tc, 'p').map(p => p.textContent).join('\n').trim())) });
        }
      }
    };
    walk(desc(root, 'spTree')[0]);
    items.sort((a, b) => Math.abs(a.y - b.y) < slideH * 0.03 ? a.x - b.x : a.y - b.y);

    // 원본 차례 슬라이드는 버리되, 큰 묶음 이름("PART 1  ON  ·  전시물 켜기")은 새 차례에 쓰려고 챙긴다
    if (/^(차\s*례|목\s*차|contents)$/i.test(title || edges[0] || '')) {
      doc.parts = items.filter(i => i.kind === 'text' && i.paras.length === 1 && i.size >= 1400 && i.bold).map(i => i.text);
      continue;
    }
    const short = i => i.kind === 'text' && i.paras.length === 1 && i.text.length <= 60 && !/^\d{1,2}$/.test(i.text);
    // 첫 장에 그림·표가 없으면 표지. 가장 큰 글씨가 제목, 그 바로 아래 줄이 부제
    if (si === 0 && !items.some(i => i.kind !== 'text') && (title || items.length)) {
      const big = title ? null : items.filter(short).sort((a, b) => b.size - a.size)[0];
      doc.title = title || big?.text || items[0].text;
      const below = big ? items.filter(i => i.y > big.y && i.text.length <= 80)[0] : null;
      doc.subtitle = subtitle || below?.text || '';
      // 표지의 "문서번호  RAIM-…" 같은 같은 줄 짝은 표지 정보로. 모르는 항목(대상 전시물 등)은 extra로 그대로 옮긴다
      doc.meta = { extra: [] };
      items.forEach((it, k) => {
        const v = items[k + 1];
        if (!v || it === big || v === big || it.text.length > 10 || v.x <= it.x || Math.abs(v.y - it.y) >= slideH * 0.02) return;
        if (META_KEYS[it.text]) doc.meta[META_KEYS[it.text]] = v.text;
        else if (it.text === '작성일') setDate(doc.meta, v.text);
        else doc.meta.extra.push([it.text, v.text]);
      });
      continue;
    }
    // 제목 틀이 없으면 큰 글씨(20pt 이상) → 위쪽 굵은 한 줄 → 머리말 순으로 제목을 찾는다
    if (!title && items.length > 1) {
      const cand = items.filter(i => short(i) && i.size >= 2000).sort((a, b) => b.size - a.size)[0]
        || items.filter(i => short(i) && i.y < slideH * 0.35 && (i.size >= 1400 || i.bold)).sort((a, b) => b.size - a.size || a.y - b.y)[0];
      if (cand) { title = cand.text; items.splice(items.indexOf(cand), 1); }
    }
    if (!title && edges.length) title = edges[0];
    // 구분 슬라이드의 장식용 큰 번호("01")는 버리고, 그런 장은 큰 구분(1수준)으로 본다
    let decorated = false;
    for (let k = items.length - 1; k >= 0; k--) if (items[k].kind === 'text' && /^\d{1,2}$/.test(items[k].text) && items[k].size >= 2000) { items.splice(k, 1); decorated = true; }
    const label = edges.find(e => e !== title); // 머리말("ON", "미디어 파사드")
    if (!title && !items.length) continue;
    // 제목만 있는 장은 큰 구분(장), 나머지는 한 장이 소제목 하나.
    // 앞 장과 제목이 같으면(단계별 슬라이드) repeat 표시: 슬라이드는 나누되 문서에는 제목을 한 번만 쓴다
    const prev = doc.blocks.findLast(b => b.t === 'h');
    doc.blocks.push({ t: 'h', level: items.length && !decorated ? 2 : 1, text: title, ...(label ? { label } : {}), ...(title && prev?.text === title ? { repeat: true } : {}) });
    let list = 0;
    for (const it of items) {
      if (it.kind === 'text') {
        list++;
        for (const p of it.paras) {
          if (p.bullet) doc.blocks.push({ t: 'li', runs: p.runs, level: Math.min(p.level, 2), kind: p.bullet, list: 1000 + list });
          else doc.blocks.push({ t: 'p', runs: p.runs });
        }
      } else if (it.kind === 'pic') await addImage(doc, await zip.file(it.target).async('uint8array'), it.target, { rot: it.rot, crop: it.crop });
      else pushTable(doc, it.rows);
    }
  }
  flushSkipped(doc);
}

// ---------- hwpx ----------

async function readHwpx(buf, doc) {
  const zip = await JSZip.loadAsync(buf);
  const header = xml(await zip.file('Contents/header.xml').async('string')).documentElement;
  const styleLevel = {}, boldChar = {}, outlinePara = {};
  for (const s of desc(header, 'style')) {
    const m = (s.getAttribute('name') || '').match(/개요\s*(\d)|제목\s*(\d)/) || (s.getAttribute('engName') || '').match(/Outline\s*(\d)|Heading\s*(\d)/i);
    if (m) styleLevel[s.getAttribute('id')] = +(m[1] || m[2]);
  }
  for (const c of desc(header, 'charPr')) if (kid(c, 'bold')) boldChar[c.getAttribute('id')] = true;
  for (const p of desc(header, 'paraPr')) {
    const hd = kid(p, 'heading');
    if (hd?.getAttribute('type') === 'OUTLINE') outlinePara[p.getAttribute('id')] = +hd.getAttribute('level') + 1;
  }
  // 그림 id → 파일 경로
  const bin = {};
  const hpf = zip.file('Contents/content.hpf');
  if (hpf) for (const it of desc(xml(await hpf.async('string')).documentElement, 'item')) bin[it.getAttribute('id')] = it.getAttribute('href');

  const sections = Object.keys(zip.files).filter(n => /^Contents\/section\d+\.xml$/.test(n)).sort((a, b) => a.match(/\d+/)[0] - b.match(/\d+/)[0]);
  const cellText = tc => desc(tc, 't').map(t => t.textContent).join('').trim() && desc(tc, 'p').map(p => desc(p, 't').map(t => t.textContent).join('')).filter(s => s.trim()).join('\n');

  const para = async p => {
    const runs = [];
    const later = [];
    for (const r of kids(p, 'run')) {
      for (const c of r.children) {
        if (c.localName === 't') runs.push({ text: c.textContent, ...(boldChar[r.getAttribute('charPrIDRef')] ? { b: true } : {}) });
        else if (c.localName === 'tbl') later.push(c);
        else if (c.localName === 'pic') later.push(c);
      }
    }
    const text = runsText(runs).trim();
    if (text) {
      const level = styleLevel[p.getAttribute('styleIDRef')] || outlinePara[p.getAttribute('paraPrIDRef')];
      const m = text.match(/^(☐|□|☑|\[\s?\])\s*/);
      if (level && text.length <= 60) doc.blocks.push({ t: 'h', level: Math.min(level, 3), text });
      else if (m) doc.blocks.push({ t: 'li', runs: stripLead(runs, m[0].length), level: 0, kind: 'check', list: 0 });
      else doc.blocks.push({ t: 'p', runs });
    }
    for (const c of later) {
      if (c.localName === 'tbl') pushTable(doc, kids(c, 'tr').map(tr => kids(tr, 'tc').map(tc => cellText(tc) || '')));
      else {
        const ref = desc(c, 'img')[0]?.getAttribute('binaryItemIDRef');
        const href = bin[ref] || Object.keys(zip.files).find(n => n.startsWith('BinData/' + ref + '.'));
        if (href && zip.file(href)) await addImage(doc, await zip.file(href).async('uint8array'), href);
      }
    }
  };
  for (const s of sections) {
    const root = xml(await zip.file(s).async('string')).documentElement;
    for (const p of kids(root, 'p')) await para(p);
  }
  flushSkipped(doc);
}

// ---------- hwp (한글 97 이후 바이너리) ----------

async function inflateRaw(u8) {
  const reader = new Blob([u8]).stream().pipeThrough(new DecompressionStream('deflate-raw')).getReader();
  const parts = [];
  let n = 0;
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      parts.push(value); n += value.length;
    }
  } catch (e) { if (!n) throw e; } // 압축 끝 뒤에 붙은 찌꺼기 바이트는 한글도 무시한다
  const out = new Uint8Array(n);
  let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
}

async function readHwp(buf, doc) {
  const cfb = CFB.read(new Uint8Array(buf), { type: 'array' });
  const find = name => cfb.FileIndex[cfb.FullPaths.findIndex(p => p.endsWith('/' + name))];
  const header = find('FileHeader');
  const head = header && new Uint8Array(header.content);
  if (!head || !new TextDecoder().decode(head.subarray(0, 17)).startsWith('HWP Document File')) throw new Error('한글 97 이전 형식이거나 손상된 hwp라 읽을 수 없어요. 한글에서 hwpx로 다시 저장해 주세요.');
  const flags = new DataView(head.buffer, head.byteOffset).getUint32(36, true);
  if (flags & 2) throw new Error('암호가 걸린 hwp는 읽을 수 없어요.');
  if (flags & 4) throw new Error('배포용 hwp는 읽을 수 없어요. 원본에서 hwpx로 저장해 주세요.');
  const get = async name => {
    const e = find(name);
    if (!e) return null;
    const u8 = new Uint8Array(e.content);
    return flags & 1 ? inflateRaw(u8) : u8;
  };
  const sections = [];
  for (let i = 0; ; i++) {
    const s = await get('BodyText/Section' + i);
    if (!s) break;
    sections.push(s);
  }
  parseHwpRecords(await get('DocInfo'), sections, doc);
  doc.warnings.push('hwp 문서라 표 모양과 글자 서식은 단순하게 옮겨졌어요. 그림은 옮기지 못해요.');
}

// 레코드 단위 파서. 브라우저 API를 안 써서 node에서도 돈다
function parseHwpRecords(docInfo, sections, doc) {
  const records = function* (u8) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let o = 0;
    while (o + 4 <= u8.length) {
      const h = dv.getUint32(o, true); o += 4;
      let size = h >>> 20;
      if (size === 0xfff) { size = dv.getUint32(o, true); o += 4; }
      yield { tag: h & 0x3ff, level: (h >>> 10) & 0x3ff, data: u8.subarray(o, o + size) };
      o += size;
    }
  };
  const u16 = (d, o) => d[o] | (d[o + 1] << 8);
  const utf16 = (d, o, n) => String.fromCharCode(...Array.from({ length: n }, (_, i) => u16(d, o + i * 2)));

  // 스타일 이름 → 개요 수준
  const styleLevel = [];
  if (docInfo) for (const r of records(docInfo)) {
    if (r.tag !== 26) continue; // HWPTAG_STYLE
    const name = utf16(r.data, 2, u16(r.data, 0));
    const m = name.match(/개요\s*(\d)|제목\s*(\d)/);
    styleLevel.push(m ? +(m[1] || m[2]) : 0);
  }

  const decode = d => {
    let s = '';
    for (let i = 0; i + 1 < d.length; i += 2) {
      const c = u16(d, i);
      if (c >= 32) s += String.fromCharCode(c);
      else if (c === 10) s += '\n';
      else if (c === 30 || c === 31) s += ' ';
      else if (c === 9) { s += '\t'; i += 14; }
      else if (c === 0 || c === 13 || (c >= 24 && c <= 29)) { /* 한 글자짜리 제어 문자 */ }
      else i += 14; // 인라인·확장 제어 문자는 8글자(16바이트) 차지
    }
    return s.replace(/\s+$/, '');
  };

  for (const sec of sections) {
    let style = 0, table = null;
    // 셀마다 적힌 행·열 주소로 격자에 넣는다. 병합은 풀리고 빈 줄·빈 열은 지운다
    // ponytail: 병합 모양은 안 살림. 필요하면 IR 표에 span을 추가
    const finishTable = () => {
      const { rows, cols, cells } = table;
      table = null;
      if (!rows || !cols || cells.some(c => c.row >= rows || c.col >= cols)) { cells.forEach(c => pushText(doc, c.text.join(' '))); return; }
      const grid = Array.from({ length: rows }, () => Array(cols).fill(''));
      for (const c of cells) grid[c.row][c.col] = c.text.join('\n');
      const keep = [...Array(cols).keys()].filter(ci => grid.some(r => r[ci].trim()));
      pushTable(doc, grid.map(r => keep.map(ci => r[ci])));
    };
    for (const r of records(sec)) {
      if (table && r.level <= table.level) finishTable();
      const d = r.data;
      if (r.tag === 71 && !table) { // CTRL_HEADER
        const id = String.fromCharCode(d[3], d[2], d[1], d[0]);
        if (id === 'tbl ') table = { level: r.level, rows: 0, cols: 0, cells: [] };
      } else if (table && r.tag === 77 && r.level === table.level + 1 && !table.rows) { // TABLE
        table.rows = u16(d, 4); table.cols = u16(d, 6);
      } else if (table && r.tag === 72 && r.level === table.level + 1) { // LIST_HEADER = 셀 시작
        table.cells.push({ col: u16(d, 8), row: u16(d, 10), text: [] });
      } else if (r.tag === 66) { // PARA_HEADER
        style = d[10];
      } else if (r.tag === 67) { // PARA_TEXT
        const text = decode(d);
        if (!text.trim()) continue;
        if (table) { if (table.cells.length) table.cells[table.cells.length - 1].text.push(text.trim()); continue; }
        const level = styleLevel[style];
        if (level && text.length <= 60) doc.blocks.push({ t: 'h', level: Math.min(level, 3), text: text.trim() });
        else pushText(doc, text);
      }
    }
    if (table) finishTable();
  }
  doc.warnings = [...new Set(doc.warnings)];
}

if (typeof module !== 'undefined') module.exports = { parseHwpRecords };
