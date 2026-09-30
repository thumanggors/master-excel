// Build eBook MASTER RUMUS EXCEL: content/*.txt -> dist/*.html -> dist/*.pdf (Chrome headless)
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { execFileSync } from 'node:child_process';

const root = path.dirname(fileURLToPath(import.meta.url));
const dist = path.join(root, 'dist');
fs.mkdirSync(dist, { recursive: true });

const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const usedIds = new Set();
const slug = t => {
  let b = t.toLowerCase().replace(/<[^>]+>/g, '').replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 's';
  let s = b, n = 2;
  while (usedIds.has(s)) s = `${b}-${n++}`;
  usedIds.add(s);
  return s;
};

// ---------- Excel syntax highlight ----------
const TOK = /("(?:[^"]|"")*")|(\/\/.*$)|([A-Za-z_][A-Za-z0-9_.]*)(?=\()|((?<![A-Za-z0-9_])(?:[A-Za-z_][A-Za-z0-9_]*!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?(?![A-Za-z0-9_])|(?<![A-Za-z0-9_])\$?[A-Z]{1,3}:\$?[A-Z]{1,3}(?![A-Za-z0-9_]))|((?<![A-Za-z0-9_])\d+(?:\.\d+)?%?)/g;
function hl(src) {
  let out = '', last = 0;
  src.replace(TOK, (m, str, com, fn, ref, num, off) => {
    out += esc(src.slice(last, off));
    last = off + m.length;
    out += `<span class="${str ? 's' : com ? 'c' : fn ? 'f' : ref ? 'r' : 'n'}">${esc(m)}</span>`;
    return m;
  });
  return out + esc(src.slice(last));
}
const code = lines => `<pre class="xl">${lines.map(hl).join('\n')}</pre>`;

// ---------- inline ----------
function inline(s) {
  return s.split(/(`[^`]+`)/).map(p => {
    if (/^`[^`]+`$/.test(p)) {
      const c = p.slice(1, -1);
      if (c.startsWith('=')) { stats.inline++; return `<code class="xl">${hl(c)}</code>`; }
      return `<code>${esc(c)}</code>`;
    }
    return esc(p)
      .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
      .replace(/\[([^\]]+)\]\((#[^)]+)\)/g, '<a href="$2">$1</a>');
  }).join('');
}

// ---------- tables ----------
const splitRow = r => r.replace(/^\|/, '').replace(/\|$/, '').split(/(?<!\\)\|/).map(c => c.trim().replace(/\\\|/g, '|'));
const isNum = c => /^(Rp\s?)?[-+]?[\d.,]+\s?%?$/.test(c) || /^#[A-Z/0!?]+$/.test(c);
const colName = n => { let s = ''; n++; while (n) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };
function table(rows, sheet) {
  if (!rows.length) return '';
  const td = (c, tag = 'td') => `<${tag}${isNum(c) ? ' class="num"' : ''}>${inline(c)}</${tag}>`;
  if (sheet) {
    const m = sheet.match(/^([A-Z]+)(\d+)$/) || ['', 'A', '1'];
    const c0 = [...m[1]].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1, r0 = +m[2];
    const w = Math.max(...rows.map(r => r.length));
    let h = `<div class="tw"><table class="sheet${w > 6 ? ' wide' : ''}"><thead><tr><th class="corner"></th>${Array.from({ length: w }, (_, k) => `<th class="col">${colName(c0 + k)}</th>`).join('')}</tr></thead><tbody>`;
    rows.forEach((r, k) => { r = [...r, ...Array(w - r.length).fill('')]; h += `<tr${k === 0 ? ' class="hdr"' : ''}><th class="rown">${r0 + k}</th>${r.map(c => td(c)).join('')}</tr>`; });
    return h + '</tbody></table></div>';
  }
  const [head, ...body] = rows;
  return `<div class="tw${rows.length > 12 ? ' long' : ''}"><table class="grid"><thead><tr>${head.map(c => td(c, 'th')).join('')}</tr></thead><tbody>${body.map(r => `<tr>${r.map(c => td(c)).join('')}</tr>`).join('')}</tbody></table></div>`;
}

// ---------- callouts ----------
const CALL = {
  tip: ['💡', 'PRO TIP'], warn: ['⚠️', 'WARNING'], mistake: ['⚠️', 'COMMON MISTAKE'], important: ['📌', 'IMPORTANT'],
  advanced: ['🚀', 'ADVANCED'], real: ['💼', 'REAL WORLD'], challenge: ['🧠', 'CHALLENGE'], note: ['📝', 'CATATAN'], version: ['🧩', 'KOMPATIBILITAS'],
};
const callout = (type, title, lines) => {
  const [ico, label] = CALL[type] || CALL.note;
  return `<aside class="call ${type}"><div class="call-h"><span class="ico">${ico}</span>${esc(title || label)}</div>${blocks(lines)}</aside>`;
};

// ---------- directive fields ----------
function fields(buf, keys) {
  const f = {}; let cur = null;
  for (const l of buf) {
    const m = l.match(/^([a-z]+):\s?(.*)$/);
    if (m && keys.includes(m[1])) { cur = m[1]; f[cur] = m[2] ? [m[2]] : []; }
    else if (cur) f[cur].push(l);
  }
  return f;
}
const nb = a => (a || []).filter(l => l.trim());
const lbl = t => `<div class="lbl">${t}</div>`;

let fxCount = 0;
const stats = { fx: 0, formulas: 0, inline: 0 };
function fx(arg, buf) {
  const [name, ver] = arg.split('|').map(s => s.trim());
  const f = fields(buf, ['fungsi', 'syntax', 'args', 'data', 'formula', 'hasil', 'jelas', 'tip', 'salah', 'real']);
  stats.fx++;
  const args = nb(f.args).map(l => { const i = l.indexOf('|'); return [l.slice(0, i).trim(), l.slice(i + 1).trim()]; });
  const note = (t, a) => a ? callout(t, '', a) : '';
  const formula = nb(f.formula);
  stats.formulas += formula.filter(l => l.trim().startsWith('=')).length;
  return `<article class="fx" id="fx-${slug(name)}">
<div class="fx-card"><div class="fx-top"><span class="fx-kicker">FORMULA CARD</span>${ver ? `<span class="badge">${esc(ver)}</span>` : ''}</div>
<h3 class="fx-name">${esc(name)}</h3>${lbl('Fungsi')}<div class="fx-desc">${blocks(f.fungsi || [])}</div>${lbl('Syntax')}${code(nb(f.syntax))}</div>
${args.length ? `<div class="blk">${lbl('Penjelasan Argument')}${table([['Argument', 'Penjelasan'], ...args], false)}</div>` : ''}
${f.data ? `<div class="blk">${lbl('Contoh Data')}${blocks(f.data, 'A1')}</div>` : ''}
<div class="blk">${lbl('Formula')}${code(formula)}</div>
${f.hasil ? `<div class="blk">${lbl('Hasil')}<div class="result">${blocks(f.hasil)}</div></div>` : ''}
${f.jelas ? `<div class="blk">${lbl('Penjelasan')}${blocks(f.jelas)}</div>` : ''}
<div class="fx-notes">${note('tip', f.tip)}${note('mistake', f.salah)}${note('real', f.real)}</div>
</article>`;
}

function kase(arg, buf) {
  const f = fields(buf, ['problem', 'data', 'formula', 'result', 'explain', 'challenge', 'answer']);
  const part = (k, t, sheet) => f[k] ? `<div class="case-part">${lbl(t)}${blocks(f[k], sheet)}</div>` : '';
  const id = slug('case ' + arg);
  toc.push({ lvl: 2, text: 'Real World Case Study: ' + arg, id });
  return `<section class="case" id="${id}"><div class="case-h"><span>REAL WORLD CASE STUDY</span><h2>${inline(arg)}</h2></div>
${part('problem', 'Problem')}${part('data', 'Raw Data', 'A1')}${part('formula', 'Formula')}${part('result', 'Result')}${part('explain', 'Explanation')}
${f.challenge ? callout('challenge', 'CHALLENGE', f.challenge) : ''}${f.answer ? `<div class="case-part answer">${lbl('Answer')}${blocks(f.answer)}</div>` : ''}</section>`;
}

let curFxNames = [];
function chapter(arg, buf) {
  const [num, title] = arg.split('|').map(s => s.trim());
  const f = fields(buf, ['desc', 'goals', 'extra']);
  const id = 'bab-' + num;
  usedIds.add(id);
  toc.push({ lvl: 1, text: `BAB ${num} · ${title}`, id, chapter: true });
  const chips = curFxNames.map(n => `<span>${esc(n)}</span>`).join('');
  return `<section class="full-page chapter" id="${id}"><div class="ch-grid"></div><div class="ch-in">
<div class="ch-bab">BAB</div><div class="ch-num">${esc(num)}</div><h1 class="ch-title">${esc(title)}</h1>
<div class="ch-desc">${blocks(f.desc || [])}</div>
<div class="ch-cols"><div class="ch-goals"><div class="ch-lbl">Tujuan Pembelajaran</div>${blocks(f.goals || [])}</div>
<div class="ch-stat"><div class="ch-lbl">Formula Count</div><div class="ch-count">${curFxNames.length}</div><div class="ch-sub">formula card di BAB ini${f.extra ? '<br>' + inline(f.extra.join(' ')) : ''}</div></div></div>
<div class="ch-lbl">Rumus yang Dibahas</div><div class="chips">${chips}</div></div></section>`;
}

// ---------- block parser ----------
const toc = [];
let fileKind = 'front';
function heading(n, text) {
  const id = slug(text);
  if (n === 1) toc.push({ lvl: 1, text, id });
  if (n === 2 && fileKind === 'chapter') toc.push({ lvl: 2, text, id });
  return `<h${n} id="${id}">${inline(text)}</h${n}>`;
}
function blocks(lines, defSheet = false) {
  let html = '', i = 0, sheet = null;
  const para = [];
  const flush = () => { if (para.length) { html += `<p>${inline(para.join(' '))}</p>`; para.length = 0; } };
  while (i < lines.length) {
    const l = lines[i];
    let m;
    if (!l.trim()) { flush(); i++; continue; }
    if ((m = l.match(/^@@(\w+)\s*(.*)$/))) {
      flush(); i++;
      if (m[1] === 'page') { html += '<div class="pb"></div>'; continue; }
      if (m[1] === 'toc') { html += '<!--TOC-->'; continue; }
      const buf = [];
      // nested directives (e.g. @@raw inside @@case) close with their own @@end
      for (let depth = 1; i < lines.length; i++) {
        const t = lines[i].trim();
        if (t === '@@end' && --depth === 0) break;
        if (/^@@(?!end$|page$|toc$)\w/.test(t)) depth++;
        buf.push(lines[i]);
      }
      i++;
      if (m[1] === 'fx') html += fx(m[2], buf);
      else if (m[1] === 'case') html += kase(m[2], buf);
      else if (m[1] === 'chapter') html += chapter(m[2], buf);
      else if (m[1] === 'raw') html += buf.join('\n');
      else if (m[1] === 'card') { // mini formula card: first line formula, rest description
        const [nm, ...rest] = m[2].split('|');
        html += `<div class="mini"><div class="mini-n">${esc(nm.trim())}</div>${code([buf[0]])}${blocks(buf.slice(1))}</div>`;
      }
      continue;
    }
    if ((m = l.match(/^(#{1,4}) (.*)$/))) { flush(); html += heading(m[1].length, m[2]); i++; continue; }
    if (l.startsWith('```')) {
      flush(); const lang = l.slice(3).trim(), buf = []; i++;
      while (i < lines.length && !lines[i].startsWith('```')) buf.push(lines[i++]);
      i++;
      if (lang === 'excel') { stats.formulas += buf.filter(x => x.trim().startsWith('=')).length; html += code(buf); }
      else html += `<pre class="plain">${esc(buf.join('\n'))}</pre>`;
      continue;
    }
    if ((m = l.match(/^:::(\w+)\s*(.*)$/))) {
      flush(); const buf = []; let depth = 1; i++;
      while (i < lines.length) {
        if (/^:::\w/.test(lines[i])) depth++;
        else if (lines[i].trim() === ':::' && --depth === 0) break;
        buf.push(lines[i++]);
      }
      i++; html += callout(m[1], m[2], buf); continue;
    }
    if ((m = l.match(/^@(sheet|plain)\s*(\S*)\s*$/))) { flush(); sheet = m[1] === 'sheet' ? (m[2] || 'A1') : false; i++; continue; }
    if ((m = l.match(/^@cap\s+(.*)$/))) { flush(); html += `<div class="cap">${inline(m[1])}</div>`; i++; continue; }
    if (l.trim().startsWith('|')) {
      flush(); const rows = [];
      while (i < lines.length && lines[i].trim().startsWith('|')) {
        const r = lines[i++].trim();
        if (/^\|[\s:|-]+\|$/.test(r)) continue;
        rows.push(splitRow(r));
      }
      html += table(rows, sheet ?? defSheet); sheet = null; continue;
    }
    const li = /^\s*(?:[-*]|\d+\.) /;
    if (li.test(l)) {
      flush();
      const ordered = /^\s*\d+\. /.test(l), start = ordered ? parseInt(l) : 1, items = [];
      while (i < lines.length && li.test(lines[i])) {
        items.push(lines[i++].replace(li, ''));
        while (i < lines.length && /^\s{2,}\S/.test(lines[i]) && !li.test(lines[i])) items[items.length - 1] += ' ' + lines[i++].trim();
      }
      const lis = items.map(t => `<li>${inline(t)}</li>`).join('');
      html += ordered ? `<ol${start !== 1 ? ` start="${start}"` : ''}>${lis}</ol>` : `<ul>${lis}</ul>`;
      continue;
    }
    para.push(l.trim()); i++;
  }
  flush();
  return html;
}

// ---------- assemble ----------
const files = fs.readdirSync(path.join(root, 'content')).filter(f => f.endsWith('.txt')).sort();
let body = '';
for (const f of files) {
  const src = fs.readFileSync(path.join(root, 'content', f), 'utf8').replace(/\r/g, '');
  fileKind = /^bab/.test(f) ? 'chapter' : 'front';
  curFxNames = [...src.matchAll(/^@@fx\s+([^|\n]+)/gm)].map(m => m[1].trim());
  body += `<div class="file kind-${fileKind}" data-src="${f}">${blocks(src.split('\n'))}</div>\n`;
}
const tocHtml = `<nav class="toc">${toc.filter(t => t.id !== 'daftar-isi').map(t =>
  `<a class="t${t.lvl}${t.chapter ? ' tch' : ''}" href="#${t.id}">${inline(t.text)}</a>`).join('')}</nav>`;
body = body.replace('<!--TOC-->', tocHtml);

const css = fs.readFileSync(path.join(root, 'style.css'), 'utf8');
const head = title => `<!doctype html><html lang="id"><head><meta charset="utf-8"><title>${title}</title>
<meta name="author" content="Pictor Roito Tumanggor">
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&family=JetBrains+Mono:wght@400;600;700&display=block" rel="stylesheet">
<style>${css}</style></head>`;
const htmlPath = path.join(dist, 'master-rumus-excel.html');
fs.writeFileSync(htmlPath, `${head('Master Rumus Excel')}<body><main>${body}</main></body></html>`);

// cover-only page for PNG export
const coverSrc = fs.readFileSync(path.join(root, 'content', files[0]), 'utf8').replace(/\r/g, '');
fs.writeFileSync(path.join(dist, 'cover.html'), `${head('Cover')}<body class="cover-only"><main>${blocks(coverSrc.split('\n'))}</main></body></html>`);

console.log(`files=${files.length} formulaCards=${stats.fx} formulaLines=${stats.formulas} inlineFormulas=${stats.inline} tocEntries=${toc.length}`);

if (process.argv.includes('--pdf')) {
  const chrome = 'C:/Program Files/Google/Chrome/Application/chrome.exe';
  const common = ['--headless=new', '--disable-gpu', '--hide-scrollbars', '--virtual-time-budget=30000'];
  execFileSync(chrome, [...common, '--no-pdf-header-footer', '--generate-pdf-document-outline',
    `--print-to-pdf=${path.join(dist, 'MASTER-RUMUS-EXCEL.pdf')}`, pathToFileURL(htmlPath).href], { stdio: 'inherit' });
  execFileSync(chrome, [...common, '--window-size=794,1123', '--force-device-scale-factor=2',
    `--screenshot=${path.join(dist, 'cover.png')}`, pathToFileURL(path.join(dist, 'cover.html')).href], { stdio: 'inherit' });
  console.log('PDF + cover.png written to dist/');
}
