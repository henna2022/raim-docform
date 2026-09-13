// 공통 문서 구조를 A~D 양식으로 그린다. HTML(미리보기·PDF), docx, pptx
// 값은 Desktop/과학관업무/운영매뉴얼/표준양식_시안 의 docx에서 뽑았다

const THEMES = {
  A: { name: 'A 기본형', desc: '선 최소, 번호 없음', ink: '1C1C1E', sub: '6B7280', label: '9CA3AF', link: '1D4ED8', hair: 'ECEEF1',
    theadLine: 'D1D5DB', theadSz: 4, rule: '1C1C1E', ruleSz: 12, number: 'none', titlePt: 28, h1Pt: 10.5 },
  B: { name: 'B 공문서형', desc: '격자 표, 1. / 1.1 번호, 가운데 표지', ink: '1C1C1E', sub: '5B6470', label: '5B6470', link: '1D4ED8', hair: 'C9CED5',
    theadLine: '9AA1AA', theadSz: 4, theadFill: 'F2F3F5', grid: true, center: true, rule: '1C1C1E', ruleSz: 8, number: 'decimal', titlePt: 20, h1Pt: 11 },
  C: { name: 'C 컬러 액센트', desc: 'A 구조에 남색 포인트', ink: '1C1C1E', sub: '6B7280', label: '9CA3AF', link: '14396B', hair: 'ECEEF1',
    theadLine: '14396B', theadSz: 8, accent: '14396B', band: true, rule: '14396B', ruleSz: 16, number: 'decimal', titlePt: 26, h1Pt: 11 },
  D: { name: 'D 에디토리얼', desc: '명조 제목, 01 / 02 큰 번호, 넓은 여백', ink: '20242A', sub: '6B7280', label: 'B0B6BE', link: '1D4ED8', hair: 'EEF0F2',
    theadLine: '20242A', theadSz: 8, serif: true, rule: null, ruleSz: 0, number: 'big', titlePt: 32, h1Pt: 13, wide: true },
};
for (const k in THEMES) THEMES[k].key = k;

// PPT(사진 위주 매뉴얼)를 올렸을 때 고르는 슬라이드 양식. P1은 기기매뉴얼 원본 배치, P2~P4는 문서 양식 B~D와 짝
const SLIDE_THEMES = {
  P1: { ...THEMES.A, key: 'P1', deck: 'basic', name: '기본형', desc: '지금 기기매뉴얼 그대로' },
  P2: { ...THEMES.B, key: 'P2', deck: 'frame', name: '액자형', desc: '사진마다 가는 테두리, 가운데 구분 장', label: '8A929C', number: 'none' },
  P3: { ...THEMES.C, key: 'P3', deck: 'accent', name: '컬러 포인트', desc: '남색 선과 남색 구분 장', number: 'none' },
  P4: { ...THEMES.D, key: 'P4', deck: 'editorial', name: '에디토리얼', desc: '명조 제목, 큰 단계 번호', number: 'none' },
};

const ORG = '서울로봇인공지능과학관';
const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const NUM_RE = /^\s*(chapter\s*\d+[.:]?|제\s*\d+\s*[장절편부][.:]?|\d+(\.\d+)*[.)]?|[IVX]+\.|[①-⑳]|[가-하][.)])\s+/i;

function fonts(theme, opt) {
  const body = opt.font === 'malgun' ? '맑은 고딕' : 'Pretendard';
  const head = theme.serif ? (opt.font === 'malgun' ? '바탕' : 'Nanum Myeongjo') : body;
  return { body, head };
}

// 제목 수준을 1부터 맞추고 양식 번호를 붙인다
function outline(doc, theme) {
  const hs = doc.blocks.filter(b => b.t === 'h' && b.text.trim());
  const top = hs.length ? Math.min(...hs.map(h => h.level)) : 1;
  let n1 = 0, n2 = 0, lastNum = '';
  return doc.blocks.map(b => {
    if (b.t !== 'h') return b;
    const level = Math.max(1, Math.min(b.level - top + 1, 3));
    const text = theme.number === 'none' ? b.text.trim() : b.text.replace(NUM_RE, '').trim();
    if (!text) return { ...b, level, text, num: '' };
    if (b.repeat) return { ...b, level, text, num: lastNum };
    let num = '';
    if (level === 1) { n1++; n2 = 0; num = theme.number === 'decimal' ? `${n1}.` : theme.number === 'big' ? String(n1).padStart(2, '0') : ''; }
    else if (level === 2) { n2++; num = theme.number === 'decimal' ? `${Math.max(n1, 1)}.${n2}` : ''; }
    lastNum = num;
    return { ...b, level, text, num };
  });
}

// 목록 기호. 같은 목록 안에서 수준별로 번호를 센다
function listMarker() {
  const counts = {};
  return b => {
    if (b.kind === 'check') return '☐';
    if (b.kind === 'bullet') return ['–', '·', '–'][b.level] || '–';
    for (const k of Object.keys(counts)) if (k.startsWith(b.list + ':') && +k.split(':')[1] > b.level) delete counts[k];
    const key = b.list + ':' + b.level;
    counts[key] = (counts[key] || 0) + 1;
    return counts[key] + '.';
  };
}

function noteParts(runs) {
  const m = runsText(runs).match(/^\s*(※|주의|유의|경고)\s*[:：]?\s*/);
  return { label: m ? m[1] : '주의', runs: m ? stripLead(runs, m[0].length) : runs };
}

// extra: 원본 표지에만 있던 줄(대상 전시물, 운영 기간 등)
const metaRows = meta => [['문서번호', meta.docNo], ['버전', meta.version], ['작성일', meta.dateText], ...(meta.extra || []), ['작성부서', meta.dept], ['작성자', meta.author], ['승인자', meta.approver]];

// ---------- HTML (미리보기와 PDF) ----------

function themeVars(theme, opt) {
  const f = fonts(theme, opt);
  const c = k => '#' + theme[k];
  return [`--ink:${c('ink')}`, `--sub:${c('sub')}`, `--label:${c('label')}`, `--link:${c('link')}`, `--hair:${c('hair')}`,
    `--thead-line:${c('theadLine')}`, `--thead-w:${theme.theadSz / 8}pt`, `--thead-fill:${theme.theadFill ? c('theadFill') : 'transparent'}`,
    `--accent:${c(theme.accent ? 'accent' : 'ink')}`, `--rule:${theme.rule ? '#' + theme.rule : 'transparent'}`, `--rule-w:${theme.ruleSz / 8}pt`,
    `--title-pt:${theme.titlePt}pt`, `--h1-pt:${theme.h1Pt}pt`, `--font:'${f.body}'`, `--head:'${f.head}'`].join(';');
}

function runsHtml(runs) {
  return runs.map(r => {
    let h = esc(r.text).replace(/\n/g, '<br>');
    if (r.color || r.size || r.font) h = `<span style="${r.color ? `color:#${r.color};` : ''}${r.size ? `font-size:${r.size}pt;` : ''}${r.font ? `font-family:'${r.font}',Pretendard,serif;` : ''}">${h}</span>`;
    if (r.b) h = `<b>${h}</b>`;
    if (r.href) h = `<a href="${esc(r.href)}">${h}</a>`;
    return h;
  }).join('');
}

function tableHtml(rows) {
  const cell = (tag, c) => `<${tag}>${esc(c).replace(/\n/g, '<br>')}</${tag}>`;
  return `<table class="tbl"><thead><tr>${rows[0].map(c => cell('th', c)).join('')}</tr></thead><tbody>${rows.slice(1).map(r => `<tr>${r.map(c => cell('td', c)).join('')}</tr>`).join('')}</tbody></table>`;
}

function toHtml(doc, theme, meta, opt) {
  const blocks = outline(doc, theme);
  const mark = listMarker();
  let h = `<article class="doc t-${theme.key}" style="${themeVars(theme, opt)}">`;
  if (opt.cover) {
    h += `<section class="page cover">${theme.band ? `<div class="band">${ORG}</div>` : `<p class="org">${ORG}</p>`}
      <h1 class="title">${esc(meta.title)}</h1>${meta.subtitle ? `<p class="subtitle">${esc(meta.subtitle)}</p>` : ''}<hr class="cover-rule">
      <div class="meta">${tableHtml([['항목', '내용'], ...metaRows(meta).map(([k, v]) => [k, v || ''])])}</div></section>`;
  }
  if (opt.front) {
    const toc = blocks.filter(b => b.t === 'h' && b.text && !b.repeat && b.level <= 2);
    h += `<section class="page front"><h2 class="h1 plain">개정 이력</h2>${tableHtml([['버전', '개정일', '개정 내용', '작성자'], [meta.version, meta.date, '최초 작성', meta.author], ['', '', '', '']])}
      <h2 class="h1 plain toc-head">목차</h2><ol class="toc">${toc.map(b => `<li class="l${b.level}">${b.num ? `<span class="num">${esc(b.num)}</span>` : ''}${esc(b.text)}</li>`).join('')}</ol></section>`;
  }
  h += '<section class="body">';
  for (const b of blocks) {
    if (b.t === 'h') { if (b.text && !b.repeat) h += `<h${b.level + 1} class="h${b.level}">${b.num ? `<span class="num">${esc(b.num)}</span>` : ''}${esc(b.text)}</h${b.level + 1}>`; }
    else if (b.t === 'p') h += `<p>${runsHtml(b.runs)}</p>`;
    else if (b.t === 'li') h += `<p class="li k-${b.kind}" style="--lv:${b.level}" data-mark="${mark(b)}">${runsHtml(b.runs)}</p>`;
    else if (b.t === 'note') { const n = noteParts(b.runs); h += `<p class="note"><b class="note-label">${esc(n.label)}</b>${runsHtml(n.runs)}</p>`; }
    else if (b.t === 'table') h += tableHtml(b.rows);
    else if (b.t === 'img') h += `<figure><img src="${b.src}" alt=""></figure>`;
  }
  return h + '</section></article>';
}

// ---------- docx ----------

const dataBytes = src => Uint8Array.from(atob(src.slice(src.indexOf(',') + 1)), c => c.charCodeAt(0));

async function toDocx(doc, theme, meta, opt) {
  const D = docx;
  const f = fonts(theme, opt);
  const blocks = outline(doc, theme);
  const side = theme.wide ? 1474 : 1247;
  const width = 11906 - side * 2;
  const none = { style: D.BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const line = (color, size = 4) => ({ style: D.BorderStyle.SINGLE, size, color });
  const run = (text, o = {}) => new D.TextRun({ text, font: f.body, ...o });
  const runs = rs => rs.flatMap(r => {
    const parts = r.text.split('\n').map((t, i) => run(t, { bold: !!r.b, color: r.color || (r.href ? theme.link : theme.ink), size: 20, ...(i ? { break: 1 } : {}) }));
    return r.href ? [new D.ExternalHyperlink({ link: r.href, children: parts })] : parts;
  });

  const table = (rows, widths) => {
    const last = rows.length - 1;
    return new D.Table({
      width: { size: width, type: D.WidthType.DXA }, columnWidths: widths, layout: D.TableLayoutType.FIXED,
      borders: theme.grid
        ? { top: line(theme.theadLine), bottom: line(theme.theadLine), left: line(theme.theadLine), right: line(theme.theadLine), insideHorizontal: line(theme.hair), insideVertical: line(theme.hair) }
        : { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none },
      rows: rows.map((r, ri) => new D.TableRow({
        tableHeader: ri === 0, cantSplit: true,
        children: r.map((c, ci) => new D.TableCell({
          width: { size: widths[ci], type: D.WidthType.DXA },
          margins: { top: 60, bottom: 60, left: theme.grid ? 100 : 0, right: 140 },
          ...(ri === 0 && theme.theadFill ? { shading: { type: D.ShadingType.CLEAR, fill: theme.theadFill, color: 'auto' } } : {}),
          ...(theme.grid ? {} : { borders: { top: none, left: none, right: none, bottom: ri === 0 ? line(theme.theadLine, theme.theadSz) : ri === last ? none : line(theme.hair) } }),
          children: String(c).split('\n').map(t => new D.Paragraph({ spacing: { before: 0, after: 0, line: 300 }, children: [run(t, ri === 0 ? { bold: true, color: theme.label, size: 16 } : { color: theme.ink, size: 20 })] })),
        })),
      })),
    });
  };
  const autoWidths = rows => {
    const cols = rows[0].length;
    const weight = [...Array(cols).keys()].map(ci => Math.min(30, Math.max(4, ...rows.map(r => Math.max(...String(r[ci]).split('\n').map(s => s.length))))));
    const sum = weight.reduce((a, b) => a + b, 0);
    const w = weight.map(x => Math.floor(width * x / sum));
    w[cols - 1] += width - w.reduce((a, b) => a + b, 0);
    return w;
  };
  const sectionTitle = text => new D.Paragraph({ spacing: { before: 0, after: 200 }, keepNext: true,
    ...(theme.rule ? { border: { bottom: line(theme.rule, theme.ruleSz || 12), } } : {}), children: [run(text, { bold: true, color: theme.ink, size: theme.h1Pt * 2, font: f.head })] });

  const children = [];
  if (opt.cover) {
    const org = run(ORG, { bold: true, color: theme.band ? 'FFFFFF' : theme.label, size: 17, characterSpacing: 60 });
    const align = theme.center ? D.AlignmentType.CENTER : D.AlignmentType.LEFT;
    if (theme.band) {
      children.push(new D.Paragraph({ spacing: { before: 400, after: 0 } }), new D.Table({
        width: { size: width, type: D.WidthType.DXA }, columnWidths: [width],
        borders: { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none },
        rows: [new D.TableRow({ children: [new D.TableCell({ width: { size: width, type: D.WidthType.DXA }, shading: { type: D.ShadingType.CLEAR, fill: theme.accent, color: 'auto' },
          margins: { top: 260, bottom: 520, left: 200, right: 200 }, children: [new D.Paragraph({ children: [org] })] })] })],
      }));
    } else children.push(new D.Paragraph({ alignment: align, spacing: { before: 1400, after: 0 }, children: [org] }));
    children.push(
      new D.Paragraph({ alignment: align, spacing: { before: theme.band ? 1400 : 240, after: 0 }, children: [run(meta.title, { bold: true, color: theme.ink, size: theme.titlePt * 2, font: f.head })] }),
      ...(meta.subtitle ? [new D.Paragraph({ alignment: align, spacing: { before: 140, after: 0 }, children: [run(meta.subtitle, { color: theme.sub, size: 22 })] })] : []),
      new D.Paragraph({ spacing: { before: 200, after: 0 }, border: { bottom: line(theme.accent || theme.ink, theme.center || theme.band ? 8 : 12) } }),
      new D.Paragraph({ spacing: { before: 3600, after: 0 } }),
      table([['항목', '내용'], ...metaRows(meta).map(([k, v]) => [k, v || ''])], [2000, width - 2000]),
      new D.Paragraph({ children: [new D.PageBreak()] }),
    );
  }
  if (opt.front) {
    children.push(
      sectionTitle('개정 이력'),
      table([['버전', '개정일', '개정 내용', '작성자'], [meta.version, meta.date, '최초 작성', meta.author], ['', '', '', '']], [1200, 1500, width - 4400, 1700]),
      new D.Paragraph({ spacing: { before: 600 } }),
      sectionTitle('목차'),
      new D.TableOfContents('목차', { hyperlink: true, headingStyleRange: '1-2' }),
      new D.Paragraph({ children: [new D.PageBreak()] }),
    );
  }
  const HL = [D.HeadingLevel.HEADING_1, D.HeadingLevel.HEADING_2, D.HeadingLevel.HEADING_3];
  let listInstance = 0, lastList = null;
  for (const b of blocks) {
    if (b.t === 'h') {
      if (!b.text || b.repeat) continue;
      const big = theme.number === 'big' && b.level === 1;
      const numRun = b.num ? [run(big ? b.num + '   ' : b.num + ' ', { bold: true, ...(big ? { size: 30 } : {}), color: big ? theme.label : theme.accent || theme.ink })] : [];
      children.push(new D.Paragraph({ heading: HL[b.level - 1], keepNext: true,
        ...(b.level === 1 && theme.rule ? { border: { bottom: { ...line(theme.rule, theme.ruleSz), space: 4 } } } : {}),
        children: [...numRun, run(b.text, { font: b.level < 3 ? f.head : f.body })] }));
    } else if (b.t === 'p') children.push(new D.Paragraph({ children: runs(b.runs) }));
    else if (b.t === 'li') {
      if (b.kind === 'num' && b.list !== lastList) listInstance++;
      lastList = b.list;
      children.push(new D.Paragraph({ spacing: { after: 40 }, numbering: { reference: b.kind, level: b.level, ...(b.kind === 'num' ? { instance: listInstance } : {}) }, children: runs(b.runs) }));
    } else if (b.t === 'note') {
      const n = noteParts(b.runs);
      children.push(new D.Paragraph({ spacing: { before: 320, after: 60 }, indent: { left: 180 }, border: { left: { ...line(theme.accent || theme.ink, 12), space: 10 } },
        children: [run(n.label + '  ', { bold: true, size: 16, color: theme.accent || (theme.key === 'A' ? theme.label : theme.ink) }), ...runs(n.runs)] }));
    } else if (b.t === 'table') {
      children.push(table(b.rows, autoWidths(b.rows)), new D.Paragraph({ spacing: { after: 120 } }));
    } else if (b.t === 'img') {
      const maxW = width / 1440 * 96, maxH = 560;
      const k = Math.min(1, maxW / b.w, maxH / b.h);
      children.push(new D.Paragraph({ spacing: { before: 80, after: 160 }, children: [new D.ImageRun({ type: b.src.startsWith('data:image/png') ? 'png' : 'jpg', data: dataBytes(b.src), transformation: { width: Math.round(b.w * k), height: Math.round(b.h * k) } })] }));
    }
  }
  if (!children.length) children.push(new D.Paragraph(''));

  const bullet = (reference, texts, color) => ({ reference, levels: texts.map((text, level) => ({ level, format: D.LevelFormat.BULLET, text, alignment: D.AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 360 + level * 360, hanging: 260 } }, run: { color, font: f.body } } })) });
  const footer = new D.Footer({ children: [new D.Paragraph({ spacing: { before: 120 }, border: { top: line(theme.hair) }, tabStops: [{ type: D.TabStopType.RIGHT, position: width }],
    children: [run(meta.title, { size: 15, color: theme.label }), new D.TextRun({ children: ['\t', D.PageNumber.CURRENT], size: 15, color: theme.label, font: f.body })] })] });

  const document = new D.Document({
    creator: meta.author || ORG, title: meta.title,
    features: { updateFields: opt.front },
    styles: {
      default: {
        document: { run: { font: f.body, size: 20, color: theme.ink }, paragraph: { spacing: { line: 300, after: 80 } } },
        heading1: { run: { font: f.head, size: theme.h1Pt * 2, bold: true, color: theme.ink }, paragraph: { spacing: { before: 520, after: 200 } } },
        heading2: { run: { font: f.head, size: 19, bold: true, color: theme.ink }, paragraph: { spacing: { before: 280, after: 120 } } },
        heading3: { run: { font: f.body, size: 19, bold: true, color: theme.sub }, paragraph: { spacing: { before: 200, after: 80 } } },
      },
    },
    numbering: { config: [
      bullet('bullet', ['–', '·', '–'], theme.label),
      bullet('check', ['☐', '☐', '☐'], theme.ink),
      { reference: 'num', levels: [0, 1, 2].map(level => ({ level, format: D.LevelFormat.DECIMAL, text: `%${level + 1}.`, alignment: D.AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 360 + level * 360, hanging: 300 } } } })) },
    ] },
    sections: [{
      properties: { titlePage: !!opt.cover, page: { size: { width: 11906, height: 16838 }, margin: { top: theme.wide ? 1588 : 1418, bottom: 1134, left: side, right: side } } },
      footers: { default: footer, first: new D.Footer({ children: [new D.Paragraph('')] }) },
      children,
    }],
  });
  // 목차 필드는 Word가 열 때 채운다. 한컴·미리보기에서도 보이게 제목 목록을 미리 넣어 둔다
  const xmlEsc = s => esc(s).replace(/'/g, '&apos;');
  const tocXml = opt.front ? blocks.filter(b => b.t === 'h' && b.text && !b.repeat && b.level <= 2).map(b =>
    `<w:p><w:pPr><w:spacing w:after="60"/>${b.level === 2 ? '<w:ind w:left="400"/>' : ''}</w:pPr><w:r><w:rPr><w:color w:val="${b.level === 2 ? theme.sub : theme.ink}"/></w:rPr><w:t xml:space="preserve">${xmlEsc((b.num ? b.num + ' ' : '') + b.text)}</w:t></w:r></w:p>`).join('') : '';
  return fixKoreanWrap(await D.Packer.toBlob(document), 'docx', tocXml);
}

// 한글이 글자 단위로 잘리지 않게 어절 단위 줄바꿈을 켠다 (Word wordWrap=0, PowerPoint eaLnBrk=0)
async function fixKoreanWrap(blob, kind, tocXml = '') {
  const zip = await JSZip.loadAsync(blob);
  if (kind === 'docx') {
    if (tocXml) {
      const d = await zip.file('word/document.xml').async('string');
      zip.file('word/document.xml', d.replace(/(<w:instrText[^>]*>TOC [^<]*<\/w:instrText><w:fldChar w:fldCharType="separate"\/><\/w:r><\/w:p>)/, '$1' + tocXml));
    }
    const p = 'word/styles.xml';
    let s = await zip.file(p).async('string');
    s = s.replace(/<w:pPrDefault>\s*<w:pPr>/, '$&<w:kinsoku w:val="1"/><w:wordWrap w:val="0"/><w:overflowPunct w:val="1"/>')
      .replace(/<w:pPrDefault\/>|<w:pPrDefault>\s*<\/w:pPrDefault>/, '<w:pPrDefault><w:pPr><w:kinsoku w:val="1"/><w:wordWrap w:val="0"/><w:overflowPunct w:val="1"/></w:pPr></w:pPrDefault>');
    if (!s.includes('w:wordWrap')) s = s.replace(/<w:docDefaults>/, '$&<w:pPrDefault><w:pPr><w:kinsoku w:val="1"/><w:wordWrap w:val="0"/><w:overflowPunct w:val="1"/></w:pPr></w:pPrDefault>');
    zip.file(p, s);
  } else {
    for (const n of Object.keys(zip.files).filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n))) {
      const s = (await zip.file(n).async('string'))
        .replace(/<a:pPr(?![^>]*LnBrk)(?=[\s/>])/g, '<a:pPr eaLnBrk="0" latinLnBrk="0"')
        .replace(/<a:p>(?!<a:pPr)/g, '<a:p><a:pPr eaLnBrk="0" latinLnBrk="0"/>');
      zip.file(n, s);
    }
  }
  const type = kind === 'docx' ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' : 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
  return zip.generateAsync({ type: 'blob', mimeType: type, compression: 'DEFLATE' });
}

// ---------- 슬라이드 배치 (미리보기와 pptx가 같이 쓴다) ----------
// 단위는 인치, 16:9 = 13.333 × 7.5

const SW = 13.333, SH = 7.5, MX = 0.72, CW = SW - MX * 2;
const BODY_TOP = 1.72, BODY_H = 6.78 - BODY_TOP, BODY_PT = 13, LH = 1.3;

function emWidth(s) {
  let w = 0;
  for (const ch of s) {
    const c = ch.codePointAt(0);
    w += c === 32 ? 0.28 : (c >= 0x1100 && c <= 0x11ff) || (c >= 0x2e80 && c <= 0xd7af) || (c >= 0xf900 && c <= 0xfaff) || (c >= 0xff00 && c <= 0xffef) ? 0.93 : 0.56;
  }
  return w;
}
// ponytail: 글자 폭 추정치로 줄 수를 센다. 넘치는 장이 보이면 0.93/0.56 계수를 올린다
function textHeight(paras, w, pt = BODY_PT) {
  const cap = w * 72 / pt;
  return paras.reduce((sum, p) => sum + runsText(p.runs).split('\n').reduce((n, seg) => n + Math.max(1, Math.ceil(emWidth(seg) * 1.08 / cap)), 0) * pt * LH / 72 + 5 / 72, 0);
}
function splitParas(paras, w, h, pt) {
  const pages = [[]];
  for (const p of paras) {
    const cur = pages[pages.length - 1];
    if (cur.length && textHeight([...cur, p], w, pt) > h) pages.push([p]); else cur.push(p);
  }
  return pages.filter(p => p.length);
}
function fit(img, x, y, w, h, alignX = 'center') {
  const k = Math.min(w / img.w, h / img.h);
  const iw = img.w * k, ih = img.h * k;
  return { k: 'img', src: img.src, x: alignX === 'left' ? x : x + (w - iw) / 2, y: y + (h - ih) / 2, w: iw, h: ih };
}
function tableRowHeights(rows, colW, pt = 11) {
  return rows.map(r => Math.max(0.34, Math.max(...r.map((c, i) => textHeight([{ runs: [{ text: String(c) }] }], colW[i] - 0.12, pt))) + 0.1));
}

function slideLayout(doc, theme, meta, opt) {
  const S = theme.deck; // PPT 매뉴얼 양식(P1~P4)일 때만 값이 있다
  const blocks = outline(doc, theme);
  const f = fonts(theme, opt);
  // 원본 기기매뉴얼: 본문 1.66인치부터 5.02인치. 에디토리얼만 좌우 여백이 넓다
  const mx = S === 'editorial' ? 0.9 : MX, cw = SW - mx * 2;
  const top = S ? 1.66 : BODY_TOP, bodyH = S ? 5.02 : BODY_H;
  const T = (x, y, w, h, paras, o = {}) => ({ k: 'text', x, y, w, h, paras, size: BODY_PT, color: theme.ink, font: f.body, ...o });
  const line = (x, y, w, color, pt) => ({ k: 'line', x, y, w, color, pt });
  const rect = (x, y, w, h, o) => ({ k: 'rect', x, y, w, h, ...o });
  const p = (text, o = {}) => ({ runs: [{ text, ...o }] });
  // 액자형은 사진마다 가는 테두리
  const photo = (img, x, y, w, h) => S === 'frame'
    ? (im => [rect(im.x - 0.1, im.y - 0.1, im.w + 0.2, im.h + 0.2, { stroke: theme.hair, pt: 0.75 }), im])(fit(img, x + 0.1, y + 0.1, w - 0.2, h - 0.2))
    : [fit(img, x, y, w, h)];

  // 장·절 단위로 묶기. 다음 제목이 더 깊은 1수준 제목(아래에 소제목이 있는 장)은 구분 슬라이드가 된다
  const units = [];
  let cur = null, section = '';
  blocks.forEach((b, i) => {
    if (b.t === 'h') {
      const next = blocks.slice(i + 1).find(x => x.t === 'h');
      if (b.level === 1 && b.text && !b.repeat && next?.level > 1) { units.push({ divider: true, text: b.text, num: b.num, label: b.label }); section = b.text; cur = null; }
      else { cur = { title: b.text, num: b.num, repeat: b.repeat, label: b.label, section: b.level === 1 ? '' : section, blocks: [] }; if (b.level === 1) section = ''; units.push(cur); }
    } else {
      if (!cur) { cur = { title: '', num: '', section, blocks: [] }; units.push(cur); }
      cur.blocks.push(b);
    }
  });
  // PPT 양식: 구분 장 바로 뒤의 사진 한 장·짧은 글("12단계")은 구분 장에 붙인다
  if (S) for (let i = units.length - 2; i >= 0; i--) {
    const d = units[i], u = units[i + 1];
    if (!d.divider || u.divider || u.title) continue;
    const imgs = u.blocks.filter(b => b.t === 'img'), texts = u.blocks.filter(b => b.t === 'p');
    if (imgs.length <= 1 && texts.length <= 2 && imgs.length + texts.length === u.blocks.length && texts.every(b => runsText(b.runs).length <= 30)) {
      d.img = imgs[0];
      d.note = texts.map(b => runsText(b.runs).trim()).join(' ');
      units.splice(i + 1, 1);
    }
  }

  // 사진 묶음 + 설명. 1장이면 왼쪽 사진·오른쪽 설명, 여러 장이면 가로로 나열
  const imagePages = (imgs, paras) => {
    const pages = [];
    let rest = paras;
    if (imgs.length === 1) {
      const iw = cw * 0.48, gap = S ? 0.55 : 0.4, tx = mx + iw + gap, tw = cw - iw - gap;
      const chunks = splitParas(paras, tw, bodyH);
      pages.push([...photo(imgs[0], mx, top, iw, bodyH), ...(chunks[0] ? [T(tx, top, tw, bodyH, chunks[0], { valign: 'middle' })] : [])]);
      rest = chunks.slice(1).flat();
    } else {
      const per = imgs.length <= 5 ? imgs.length : 4;
      const captions = paras.length === imgs.length && paras.every(q => runsText(q.runs).length <= 140);
      const below = !captions && paras.length && textHeight(paras, cw) <= bodyH * 0.4 ? paras : null;
      const capPt = S ? 11 : 12;
      // PPT 양식은 사진 틀 높이를 4.06인치로 맞추고 설명을 그 아래 같은 줄(5.88인치)에 둔다
      const frameFor = h => S ? Math.min(4.06, bodyH + 0.18 - h - 0.16) : bodyH - h - 0.15;
      for (let g = 0; g < imgs.length; g += per) {
        const group = imgs.slice(g, g + per);
        const gap = 0.3, cellW = (cw - gap * (per - 1)) / per;
        const items = [];
        if (captions) {
          const caps = paras.slice(g, g + per);
          const capH = Math.max(S ? 0.98 : 0.5, ...caps.map(c => textHeight([c], cellW, capPt)));
          const fh = frameFor(capH);
          group.forEach((im, i) => {
            const x = mx + i * (cellW + gap);
            items.push(...photo(im, x, top, cellW, fh), T(x, top + fh + (S ? 0.16 : 0.15), cellW, capH, [caps[i]], { size: capPt }));
          });
        } else {
          const th = below && g === 0 ? textHeight(below, cw, S ? 12 : BODY_PT) : 0;
          const fh = !th ? bodyH : S ? frameFor(th) : bodyH - th - 0.25;
          group.forEach((im, i) => items.push(...photo(im, mx + i * (cellW + gap), top, cellW, fh)));
          if (th) items.push(T(mx, S ? top + fh + 0.16 : top + bodyH - th, cw, th, below, S ? { size: 12 } : {}));
        }
        pages.push(items);
      }
      rest = captions || below ? [] : paras;
    }
    return { pages, rest };
  };

  // 한 단위를 본문 장들로 나눈다. 반환: [[item...]...]
  const bodyPages = u => {
    const mark = listMarker();
    const toPara = b => b.t === 'li' ? { runs: [{ text: '    '.repeat(b.level) + mark(b) + ' ', color: b.kind === 'bullet' ? theme.label : undefined }, ...b.runs] }
      : b.t === 'note' ? (n => ({ runs: [{ text: n.label + '  ', b: true, color: theme.accent || theme.label }, ...n.runs] }))(noteParts(b.runs))
      : { runs: b.runs };
    // 원래 순서대로 글·사진·표 조각으로 나눈다
    const pieces = [];
    for (const b of u.blocks) {
      const kind = b.t === 'img' ? 'imgs' : b.t === 'table' ? 'table' : 'text';
      const last = pieces[pieces.length - 1];
      if (kind !== 'table' && last?.kind === kind) last.items.push(b); else pieces.push({ kind, items: [b] });
    }
    const imgCount = u.blocks.filter(b => b.t === 'img').length;
    // 슬라이드 한 장 분량(사진 5장 이하, 표 없음)이면 사진과 모든 설명을 한 틀에
    if (imgCount && imgCount <= 5 && !pieces.some(pc => pc.kind === 'table')) {
      pieces.splice(0, pieces.length, { kind: 'imgs', items: u.blocks.filter(b => b.t === 'img') }, { kind: 'text', items: u.blocks.filter(b => b.t !== 'img') });
    }

    const pages = [];
    let page = null, used = 0;
    const newPage = () => { page = []; pages.push(page); used = 0; };
    const room = () => bodyH - used - (used ? 0.3 : 0);
    const place = (item, h) => { page.push({ ...item, y: top + (used ? used + 0.3 : 0) }); used = (used ? used + 0.3 : 0) + h; };
    const flowText = paras => {
      while (paras.length) {
        if (!page || room() < 0.5) newPage();
        const chunk = splitParas(paras, cw, room(), BODY_PT)[0];
        const h = textHeight(chunk, cw);
        if (used && h > room()) { newPage(); continue; }
        place(T(mx, 0, cw, h, chunk), h);
        paras = paras.slice(chunk.length);
      }
    };
    for (let k = 0; k < pieces.length; k++) {
      const pc = pieces[k];
      if (pc.kind === 'imgs') {
        const caps = pieces[k + 1]?.kind === 'text' ? pieces[++k].items.map(toPara) : [];
        const r = imagePages(pc.items, caps);
        pages.push(...r.pages);
        page = null;
        if (r.rest.length) flowText(r.rest);
      } else if (pc.kind === 'text') flowText(pc.items.map(toPara));
      else {
        const tb = pc.items[0];
        const cols = tb.rows[0].length;
        const weight = [...Array(cols).keys()].map(ci => Math.min(30, Math.max(4, ...tb.rows.map(r => String(r[ci]).length))));
        const sum = weight.reduce((a, b) => a + b, 0);
        const colW = weight.map(x => cw * x / sum);
        const heights = tableRowHeights(tb.rows, colW);
        let i = 1;
        for (;;) {
          if (!page || (used && room() < heights[0] + (heights[1] || 0))) newPage();
          let h = heights[0];
          const rows = [tb.rows[0]], hs = [heights[0]];
          while (i < tb.rows.length && h + heights[i] <= room() + 0.01) { h += heights[i]; rows.push(tb.rows[i]); hs.push(heights[i]); i++; }
          if (rows.length === 1 && i < tb.rows.length) { rows.push(tb.rows[i]); hs.push(heights[i]); h += heights[i]; i++; } // 한 줄이 한 장보다 커도 넣는다
          place({ k: 'table', x: mx, w: cw, colW, rows, heights: hs }, h);
          if (i >= tb.rows.length) break;
          newPage();
        }
      }
    }
    return pages.length ? pages : [[]];
  };

  if (S) return deckSlides();

  // ---------- PPT 매뉴얼 양식 (P1~P4) ----------
  function deckSlides() {
    const accent = theme.accent || theme.ink;
    const ed = S === 'editorial';
    const pad = v => String(v).padStart(2, '0');

    const header = (label, counter) => {
      const labelColor = S === 'accent' ? accent : theme.label;
      const items = [];
      if (S === 'frame') items.push(rect(0, 0.34, SW, 0.66, { fill: theme.theadFill }));
      items.push(T(mx, 0.52, 8, 0.3, [p(label, { b: true })], { size: 11, color: labelColor }));
      if (counter && ed) items.push(T(SW - mx - 4, 0.3, 4, 0.52, [{ runs: [{ text: counter.big, size: 22, font: f.head, color: theme.ink }, { text: '  ' + counter.small, color: theme.label }] }], { size: 11, align: 'right', valign: 'bottom' }));
      else if (counter) items.push(T(SW - mx - 4, 0.52, 4, 0.3, [p(counter.text)], { size: 11, color: labelColor, align: 'right' }));
      if (S !== 'frame') items.push(line(mx, 0.92, cw, ed ? theme.hair : accent, S === 'accent' ? 2 : ed ? 0.75 : 1.5));
      return items;
    };
    const footer = (i, from = mx) => [line(from, 6.98, SW - mx - from, theme.hair, 0.75),
      T(from, 7.06, 8, 0.28, [p(meta.title)], { size: 9, color: theme.label }),
      T(SW - mx - 2, 7.06, 2, 0.28, [p(String(i + 1))], { size: 9, color: theme.label, align: 'right' })];

    const cover = () => {
      const center = S === 'frame', al = center ? 'center' : 'left';
      const it = [];
      if (S === 'accent') it.push(rect(0, 0.9, SW * 0.62, 0.9, { fill: accent }), T(mx, 1.2, 7, 0.35, [p(ORG, { b: true })], { size: 13, color: 'FFFFFF' }));
      else it.push(T(mx, 2.25, cw, 0.35, [p(ORG, { b: true })], { size: 13, color: theme.label, align: al }));
      it.push(T(mx, 2.7, cw, 1.1, [p(meta.title, { b: true })], { size: ed ? 44 : 40, font: f.head, align: al }));
      if (meta.subtitle) it.push(T(mx, 3.95, cw, 0.4, [p(meta.subtitle)], { size: 15, color: theme.sub, align: al }));
      const rx = center ? SW / 2 - 3.9 : mx, rw = center ? 7.8 : 10.5;
      it.push(center ? line(rx, 4.55, rw, theme.ink, 1) : line(mx, 4.55, ed ? 3 : cw, ed ? theme.ink : accent, ed ? 0.75 : 1.5));
      const rows = metaRows(meta).filter(([, v]) => v);
      const step = Math.min(0.46, 2.2 / Math.max(rows.length, 1));
      rows.forEach(([k, v], i) => {
        const y = 5.15 + i * step;
        it.push(T(rx, y, 1.7, 0.3, [p(k, { b: true })], { size: 10, color: theme.label }), T(rx + 1.8, y, rw - 1.8, 0.3, [p(v)], { size: 11 }), line(rx, y + 0.32, rw, theme.hair, 0.75));
      });
      return it;
    };

    const divider = (u, no, count) => {
      const num = pad(no);
      const note = u.note && !/^\d+\s*(단계|장)$/.test(u.note) ? u.note : count ? `${count}단계` : '';
      if (S === 'accent') {
        const pw = 4.9, tw = pw - mx - 0.4;
        return [rect(0, 0, pw, SH, { fill: accent }),
          T(mx, 2.2, tw, 0.9, [p(num, { b: true })], { size: 44, color: 'FFFFFF' }),
          T(mx, 3.15, tw, 1.3, [p(u.text, { b: true })], { size: 30, color: 'FFFFFF', font: f.head }),
          ...(note ? [T(mx, 4.55, tw, 0.4, [p(note)], { size: 12, color: 'FFFFFF' })] : []),
          ...(u.img ? [fit(u.img, pw + 0.7, 1.2, SW - pw - 0.7 - mx, 5.2)] : [])];
      }
      if (S === 'frame') {
        return [T(mx, 1.5, cw, 0.75, [p(num, { b: true })], { size: 36, color: theme.label, align: 'center' }),
          T(mx, 2.25, cw, 0.75, [p(u.text, { b: true })], { size: 30, align: 'center', font: f.head }),
          ...(note ? [T(mx, 3.0, cw, 0.4, [p(note)], { size: 12, color: theme.sub, align: 'center' })] : []),
          ...(u.img ? photo(u.img, SW / 2 - 2.3, 3.6, 4.6, 3.1) : [line(SW / 2 - 1.2, 3.6, 2.4, theme.ink, 1)])];
      }
      const tx = mx + 1.5, tw = (u.img ? 8.2 : SW - mx) - tx;
      return [T(mx, ed ? 2.35 : 2.7, 1.4, ed ? 1.3 : 1.0, [p(num, { b: !ed })], { size: ed ? 60 : 44, color: theme.label, font: ed ? f.head : f.body }),
        T(tx, 2.85, tw, 0.8, [p(u.text, { b: true })], { size: ed ? 32 : 30, font: f.head }),
        ...(note ? [T(tx, 3.85, tw, 0.4, [p(note)], { size: 12, color: theme.sub })] : []),
        ...(u.img ? [fit(u.img, 8.41, 1.6, SW - mx - 8.41, 4.6)] : [])];
    };

    const toc = rows => {
      const out = [];
      let items, y;
      const next = () => { items = header('차례'); out.push({ items }); y = 1.35; };
      next();
      let group = null;
      for (const r of rows) {
        if (r.group && r.group !== group) {
          if (y + 0.96 > 6.85) next();
          const part = (doc.parts || []).find(t => t.split(/\s+/).includes(r.group)) || r.group;
          items.push(T(mx, y, 8, 0.35, [p(part, { b: true })], { size: 15, font: f.head }));
          y += 0.5;
        }
        group = r.group;
        if (y + 0.46 > 6.85) next();
        const x = mx + 0.35;
        items.push(T(x, y, 0.55, 0.3, [p(pad(r.no))], { size: 11, color: S === 'accent' ? accent : theme.label }),
          T(x + 0.6, y, 5.4, 0.3, [p(r.text, { b: true })], { size: 12, font: f.head }),
          ...(r.step ? [T(x + 6.05, y, 1.6, 0.3, [p(`${r.count}단계`)], { size: 11, color: theme.sub, align: 'right' })] : []),
          line(x, y + 0.32, 7.7, theme.hair, 0.75));
        y += 0.46;
      }
      return out;
    };

    const out = [];
    if (opt.cover) out.push({ items: cover() });
    const tocAt = out.length;
    const content = [];
    let div = null;
    for (const u of units) {
      if (u.divider) { div = { u, count: 0 }; out.push({ items: [], divider: div }); continue; }
      for (const items of bodyPages(u)) {
        const s = { items, unit: u, title: u.title, div: u.section ? div : null };
        out.push(s); content.push(s);
        if (s.div) s.div.count++;
      }
    }
    // 같은 제목이 이어지는 장은 단계 묶음: 머리말에 제목, 오른쪽에 STEP k / n
    let chain = null;
    for (const s of content) {
      if (chain && s.title && s.title === chain.title && (s.unit === chain.unit || s.unit.repeat)) { chain.slides.push(s); chain.unit = s.unit; }
      else chain = { title: s.title, unit: s.unit, slides: [s] };
      s.chain = chain;
    }
    const singles = content.filter(s => s.chain.slides.length === 1);
    for (const s of content) {
      const { slides, title } = s.chain;
      if (slides.length > 1) {
        const k = slides.indexOf(s) + 1, n = slides.length;
        s.items.unshift(...header(title, { text: `STEP ${pad(k)} / ${pad(n)}`, big: pad(k), small: `/ ${pad(n)}` }));
      } else {
        const k = singles.indexOf(s) + 1, n = singles.length;
        s.items.unshift(...header(s.unit.label || s.unit.section || meta.title, { text: `${pad(k)} / ${pad(n)}`, big: pad(k), small: `/ ${pad(n)}` }),
          ...(s.title ? [T(mx, ed ? 1.02 : 1.12, cw, 0.5, [p(s.title, { b: true })], { size: ed ? 20 : 15, font: f.head })] : []));
      }
    }
    let no = 0;
    for (const s of out) if (s.divider) s.items = [...(S === 'accent' ? [] : header(s.divider.u.label || meta.title)), ...divider(s.divider.u, ++no, s.divider.count)];
    if (opt.front) {
      const rows = [];
      let group = null, gno = 0;
      for (const s of out) {
        const r = s.divider ? { text: s.divider.u.text, count: s.divider.count, step: true, group: s.divider.u.label || '' }
          : s.chain && !s.div && s.chain.slides[0] === s && s.title ? { text: s.title, count: s.chain.slides.length, step: s.chain.slides.length > 1, group: '' } : null;
        if (!r) continue;
        if (r.group !== group) { group = r.group; gno = 0; }
        r.no = ++gno;
        rows.push(r);
      }
      if (rows.length > 1) out.splice(tocAt, 0, ...toc(rows));
    }
    out.forEach((s, i) => { if (!(i === 0 && opt.cover)) s.items.push(...footer(i, s.divider && S === 'accent' ? 5.6 : mx)); });
    return out;
  }

  const slides = [];
  const content = [];
  if (opt.cover) {
    const al = theme.center ? 'center' : 'left';
    const s = [];
    if (theme.band) s.push({ k: 'rect', x: 0, y: 0.9, w: SW * 0.62, h: 0.9, fill: theme.accent }, T(MX, 1.2, 7, 0.35, [p(ORG, { b: true })], { size: 13, color: 'FFFFFF' }));
    else s.push(T(MX, 2.05, CW, 0.35, [p(ORG, { b: true })], { size: 13, color: theme.label, align: al }));
    s.push(T(MX, 2.5, CW, 1.2, [p(meta.title, { b: true })], { size: Math.min(40, theme.titlePt + 12), font: f.head, align: al, valign: 'bottom' }));
    if (meta.subtitle) s.push(T(MX, 3.8, CW, 0.45, [p(meta.subtitle)], { size: 15, color: theme.sub, align: al }));
    s.push(line(MX, 4.45, CW, theme.accent || theme.ink, 1.5));
    const rows = metaRows(meta).filter(([, v]) => v);
    rows.forEach(([k, v], i) => {
      const y = 4.85 + i * 0.36;
      s.push(T(MX, y, 1.7, 0.3, [p(k, { b: true })], { size: 10, color: theme.label }), T(MX + 1.8, y, 6, 0.3, [p(v)], { size: 11 }), line(MX, y + 0.33, 7.8, theme.hair, 0.75));
    });
    slides.push({ items: s });
  }
  let tocAt = -1;
  if (opt.front && units.length > 1) { tocAt = slides.length; slides.push(null); }

  const tocRows = [];
  // 글만 있는 짧은 소제목들은 같은 장 안에서 한 슬라이드에 이어 담는다
  const SUB_H = 0.4, GAP = 0.3;
  const sub = (u, y) => T(MX, y, CW, SUB_H, [{ runs: [...(u.num ? [{ text: u.num + ' ', color: theme.accent || theme.label }] : []), { text: u.title }] }], { size: 14, bold: true, font: f.head });
  let pack = null, dividerNo = 0;
  for (const u of units) {
    if (u.divider) {
      pack = null;
      slides.push({ items: [
        T(MX, 2.2, CW, 1, [p(String(++dividerNo).padStart(2, '0'))], { size: 44, color: theme.label, bold: true }),
        T(MX, 3.2, CW, 1, [p(u.text, { b: true })], { size: 30, font: f.head }),
        line(MX, 4.35, 2.4, theme.rule || theme.ink, 2),
      ] });
      tocRows.push({ text: u.text, num: u.num, count: 0, level: 1 });
      continue;
    }
    const pages = bodyPages(u);
    // 빈 소제목(목차 골격)도 글만 있는 소제목처럼 묶는다
    const textOnly = u.title && pages.length === 1 && (pages[0].length === 0 || (pages[0].length === 1 && pages[0][0].k === 'text'));
    const body = textOnly && pages[0][0];
    const bodyH = body ? body.h : 0;
    if (textOnly && pack && pack.section === u.section && pack.used + GAP + SUB_H + bodyH <= BODY_H) {
      const s = pack.slide;
      if (!pack.merged) {
        const first = s.items[0];
        s.items = [sub(s, BODY_TOP), ...(first ? [{ ...first, y: BODY_TOP + SUB_H }] : [])];
        pack.used = SUB_H + (first ? first.h : 0);
        s.title = u.section; s.num = ''; s.section = '';
        pack.merged = true;
      }
      const y = BODY_TOP + pack.used + (pack.used ? GAP : 0);
      s.items.push(sub(u, y), ...(body ? [{ ...body, y: y + SUB_H }] : []));
      pack.used = y - BODY_TOP + SUB_H + bodyH;
      tocRows.push({ text: u.title, num: u.num, count: 0, level: u.section ? 2 : 1 });
      continue;
    }
    if (u.title && !u.repeat) tocRows.push({ text: u.title, num: u.num, count: pages.length, level: u.section ? 2 : 1 });
    else if (tocRows.length) tocRows[tocRows.length - 1].count += pages.length;
    pack = null;
    for (const items of pages) {
      const s = { items: [], content: true, section: u.section, title: u.title, num: u.num };
      s.items.push(...items);
      slides.push(s); content.push(s);
    }
    if (textOnly) pack = { slide: content[content.length - 1], section: u.section, used: bodyH, merged: false };
  }

  // 머리말·꼬리말은 전체 장 수가 정해진 뒤에 붙인다
  content.forEach((s, i) => {
    s.items.unshift(
      T(MX, 0.5, 8, 0.3, [p(s.section || meta.title, { b: true })], { size: 11, color: theme.label }),
      T(SW - MX - 3, 0.5, 3, 0.3, [p(`${String(i + 1).padStart(2, '0')} / ${String(content.length).padStart(2, '0')}`)], { size: 11, color: theme.label, align: 'right' }),
      line(MX, 0.92, CW, theme.rule || theme.ink, theme.accent ? 2 : 1.5),
      ...(s.title ? [T(MX, 1.1, CW, 0.45, [{ runs: [...(s.num ? [{ text: s.num + (theme.number === 'big' ? '   ' : ' '), color: theme.number === 'big' ? theme.label : theme.accent || theme.ink }] : []), { text: s.title }] }], { size: 16, bold: true, font: f.head })] : []),
    );
  });
  if (tocAt >= 0) {
    // 줄이 많으면 큰 장만 남기고(아래 소제목 장 수는 합친다), 그래도 넘치면 차례를 여러 장으로
    let rows = tocRows;
    if (rows.length > 13 && rows.some(r => r.level === 1)) {
      rows = [];
      for (const r of tocRows) { if (r.level === 1 || !rows.length) rows.push({ ...r }); else rows[rows.length - 1].count += r.count; }
    }
    const tocSlides = [];
    for (let at = 0; at < rows.length; at += 13) {
      const items = [T(MX, 0.5, 8, 0.3, [p('차례', { b: true })], { size: 11, color: theme.label }), line(MX, 0.92, CW, theme.rule || theme.ink, theme.accent ? 2 : 1.5)];
      rows.slice(at, at + 13).forEach((r, i) => {
        const y = 1.3 + i * 0.4;
        const ind = r.level === 2 ? 0.4 : 0;
        items.push(
          T(MX + ind, y, 0.6, 0.32, [p(r.num || String(at + i + 1).padStart(2, '0'))], { size: 11, color: theme.label }),
          T(MX + ind + 0.65, y, 7, 0.32, [p(r.text, r.level === 1 ? { b: true } : {})], { size: 12 }),
          ...(r.count ? [T(MX + 8, y, 1.2, 0.32, [p(r.count + '장')], { size: 11, color: theme.label, align: 'right' })] : []),
          line(MX, y + 0.36, 9.2, theme.hair, 0.75),
        );
      });
      tocSlides.push({ items });
    }
    slides.splice(tocAt, 1, ...tocSlides);
  }
  // 전체 쪽번호 꼬리말
  slides.forEach((s, i) => {
    if (i === 0 && opt.cover) return;
    s.items.push(line(MX, 6.98, CW, theme.hair, 0.75),
      T(MX, 7.05, 8, 0.25, [p(meta.title)], { size: 9, color: theme.label }),
      T(SW - MX - 2, 7.05, 2, 0.25, [p(String(i + 1))], { size: 9, color: theme.label, align: 'right' }));
  });
  return slides;
}

function slidesHtml(doc, theme, meta, opt) {
  return slideLayout(doc, theme, meta, opt).map(s => slideHtml(s, theme, opt)).join('');
}

function slideHtml(s, theme, opt) {
  const px = v => (v * 96).toFixed(1) + 'px';
  const f = fonts(theme, opt);
  return `<div class="slide" style="font-family:'${f.body}',Pretendard,sans-serif">${s.items.map(it => {
    const box = `left:${px(it.x)};top:${px(it.y)};width:${px(it.w)};`;
    if (it.k === 'line') return `<div class="ln" style="${box}border-top:${it.pt}pt solid #${it.color}"></div>`;
    if (it.k === 'rect') return `<div style="${box}height:${px(it.h)};background:${it.fill ? '#' + it.fill : 'none'};${it.stroke ? `border:${it.pt}pt solid #${it.stroke}` : ''}"></div>`;
    if (it.k === 'img') return `<img style="${box}height:${px(it.h)}" src="${it.src}" alt="">`;
    if (it.k === 'table') {
      return `<table class="stbl${theme.grid ? ' grid' : ''}" style="${box}--thead-line:#${theme.theadLine};--hair:#${theme.hair};--thead-fill:${theme.theadFill ? '#' + theme.theadFill : 'transparent'};color:#${theme.ink}">
        <colgroup>${it.colW.map(w => `<col style="width:${px(w)}">`).join('')}</colgroup>
        ${it.rows.map((r, ri) => `<tr style="height:${px(it.heights[ri])}">${r.map(c => ri === 0 ? `<th style="color:#${theme.label}">${esc(c).replace(/\n/g, '<br>')}</th>` : `<td>${esc(c).replace(/\n/g, '<br>')}</td>`).join('')}</tr>`).join('')}</table>`;
    }
    const valign = { middle: 'center', bottom: 'flex-end' }[it.valign] || 'flex-start';
    return `<div class="tx" style="${box}height:${px(it.h)};justify-content:${valign};text-align:${it.align || 'left'};font-size:${it.size}pt;color:#${it.color};${it.bold ? 'font-weight:700;' : ''}font-family:'${it.font}',Pretendard,sans-serif">${it.paras.map(q => `<p>${runsHtml(q.runs)}</p>`).join('')}</div>`;
  }).join('')}</div>`;
}

async function toPptx(doc, theme, meta, opt) {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_WIDE';
  pres.title = meta.title;
  const f = fonts(theme, opt);
  for (const s of slideLayout(doc, theme, meta, opt)) {
    const sl = pres.addSlide();
    sl.background = { color: 'FFFFFF' };
    for (const it of s.items) {
      if (it.k === 'line') sl.addShape(pres.ShapeType.line, { x: it.x, y: it.y, w: it.w, h: 0, line: { color: it.color, width: it.pt } });
      else if (it.k === 'rect') sl.addShape(pres.ShapeType.rect, { x: it.x, y: it.y, w: it.w, h: it.h,
        ...(it.fill ? { fill: { color: it.fill } } : {}), line: it.stroke ? { color: it.stroke, width: it.pt } : { color: it.fill, width: 0 } });
      else if (it.k === 'img') sl.addImage({ data: it.src.slice(5), x: it.x, y: it.y, w: it.w, h: it.h });
      else if (it.k === 'table') {
        const no = { type: 'none' }, ln = (c, pt = 0.75) => ({ type: 'solid', pt, color: c });
        const last = it.rows.length - 1;
        sl.addTable(it.rows.map((r, ri) => r.map(c => ({ text: String(c), options: {
          bold: ri === 0, color: ri === 0 ? theme.label : theme.ink, fontSize: ri === 0 ? 10 : 11,
          ...(ri === 0 && theme.theadFill ? { fill: { color: theme.theadFill } } : {}),
          border: theme.grid ? [ln(theme.hair), ln(theme.hair), ln(theme.hair), ln(theme.hair)] : [no, no, ri === 0 ? ln(theme.theadLine, theme.theadSz / 8) : ri === last ? no : ln(theme.hair), no],
        } }))), { x: it.x, y: it.y, w: it.w, colW: it.colW, rowH: it.heights, fontFace: f.body, valign: 'middle', margin: [0.03, 0.08, 0.03, theme.grid ? 0.08 : 0] });
      } else {
        const runs = it.paras.flatMap((q, qi) => q.runs.map((r, ri) => ({ text: r.text, options: {
          bold: !!(r.b || it.bold), color: r.color || (r.href ? theme.link : it.color),
          ...(r.size ? { fontSize: r.size } : {}), ...(r.font ? { fontFace: r.font } : {}),
          ...(r.href ? { hyperlink: { url: r.href } } : {}),
          ...(ri === q.runs.length - 1 && qi < it.paras.length - 1 ? { breakLine: true } : {}),
        } })));
        sl.addText(runs, { x: it.x, y: it.y, w: it.w, h: Math.max(it.h, 0.2), fontFace: it.font, fontSize: it.size, color: it.color, align: it.align || 'left',
          valign: it.valign === 'middle' ? 'middle' : it.valign === 'bottom' ? 'bottom' : 'top', margin: 0, lineSpacingMultiple: LH, paraSpaceAfter: 5, fit: 'none' });
      }
    }
  }
  return fixKoreanWrap(await pres.write({ outputType: 'blob' }), 'pptx');
}
