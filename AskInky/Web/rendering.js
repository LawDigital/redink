/* Part of Red Ink (Ask Inky). Copyright (c) LawDigital Ltd., Switzerland.
 * File: rendering.js
 * Purpose: Embedded desktop web user interface behavior.
 * License: Red Ink License; third-party attribution in embedded LICENSE.txt.
 */
/* Deterministic display/clipboard rendering. Untrusted model content never becomes executable UI. */
(() => {
    'use strict';
    const allowed = new Set('p br strong b em i s del code pre ul ol li blockquote h1 h2 h3 h4 h5 h6 table thead tbody tfoot tr th td span div sup sub dl dt dd hr a abbr kbd samp mark'.split(' '));
    const forbidden = new Set('script style iframe frame frameset object embed svg math form button select textarea video audio source canvas link meta base noscript'.split(' '));
    const block = new Set('p div h1 h2 h3 h4 h5 h6 blockquote pre ul ol dl table hr'.split(' '));
    const cellStyle = 'border:1px solid #d1cfcc;padding:6px 8px;vertical-align:top;';
    function safeLink(value) {
        const raw = value.trim();
        if (/[\u0000-\u0020\u007f]/.test(raw)) return '';
        if (raw.startsWith('#')) return '#answer-' + raw.slice(1).replace(/[^\w.:-]/g, '-');
        try {
            const uri = new URL(raw);
            return ['https:', 'http:', 'mailto:'].includes(uri.protocol) && !uri.username && !uri.password ? uri.href : '';
        } catch { return ''; }
    }
    function prepare(source, fontSize = 12, imageLabel = 'Image') {
        const template = document.createElement('template');
        template.innerHTML = source || '';
        const root = document.createElement('div');
        let images = 0;
        let removed = 0;
        function clean(node) {
            if (node.nodeType === Node.TEXT_NODE) return document.createTextNode(node.data);
            if (node.nodeType !== Node.ELEMENT_NODE) return document.createDocumentFragment();
            const tag = node.localName.toLowerCase();
            if (tag === 'img') {
                images++;
                const replacement = document.createElement('span');
                replacement.textContent = `[${imageLabel}${node.getAttribute('alt') ? ': ' + node.getAttribute('alt') : ''}]`;
                return replacement;
            }
            if (tag === 'input') {
                if (node.getAttribute('type') === 'checkbox') return document.createTextNode(node.hasAttribute('checked') ? '☑ ' : '☐ ');
                removed++;
                return document.createDocumentFragment();
            }
            if (forbidden.has(tag) || node.namespaceURI !== 'http://www.w3.org/1999/xhtml') {
                removed++;
                return document.createDocumentFragment();
            }
            const result = allowed.has(tag) ? document.createElement(tag) : document.createDocumentFragment();
            if (result.nodeType === Node.ELEMENT_NODE) {
                if (tag === 'a') {
                    const href = safeLink(node.getAttribute('href') || '');
                    if (href) result.setAttribute('href', href);
                }
                if (node.hasAttribute('title')) result.title = node.getAttribute('title');
                if (node.id) result.id = 'answer-' + node.id.replace(/[^\w.:-]/g, '-');
                if (tag === 'ol' && /^-?\d{1,6}$/.test(node.getAttribute('start') || '')) result.setAttribute('start', node.getAttribute('start'));
                if (tag === 'td' || tag === 'th') {
                    for (const name of ['colspan', 'rowspan']) {
                        const value = Number(node.getAttribute(name));
                        if (Number.isInteger(value) && value > 0 && value <= 100) result.setAttribute(name, String(value));
                    }
                    const align = node.style.textAlign || node.getAttribute('align');
                    if (['left', 'right', 'center'].includes(align)) result.style.textAlign = align;
                }
                if (tag === 'code') {
                    for (const name of node.classList) if (/^language-[\w+-]{1,40}$/.test(name)) result.classList.add(name);
                }
            }
            for (const child of node.childNodes) result.appendChild(clean(child));
            return result;
        }
        for (const child of template.content.childNodes) root.appendChild(clean(child));
        // Protect literal code from outer whitespace normalization (including leading indentation).
        const literals = [];
        let marker = '\uE000';
        while (root.textContent.includes(marker)) marker += '\uE000';
        function literal(value) { literals.push(value); return marker + (literals.length - 1) + marker; }
        function plain(node, depth = 0) {
            if (node.nodeType === Node.TEXT_NODE) return node.data.replace(/\s+/g, ' ');
            const tag = node.localName;
            if (tag === 'pre') return '\n' + literal(node.textContent.replace(/\r\n?/g, '\n')) + '\n';
            if (tag === 'code') return literal(node.textContent);
            if (tag === 'br') return '\n';
            if (tag === 'hr') return '\n';
            if (tag === 'ul' || tag === 'ol') {
                let index = Number(node.getAttribute('start') || 1);
                return '\n' + Array.from(node.children).filter(x => x.localName === 'li').map(li => {
                    const value = Array.from(li.childNodes).map(x => plain(x, depth + 1)).join('').trim();
                    return '  '.repeat(depth) + (tag === 'ol' ? `${index++}. ` : '• ') + value;
                }).join('\n') + '\n';
            }
            if (tag === 'table') {
                return '\n' + Array.from(node.querySelectorAll('tr')).map(row => Array.from(row.children).map(cell =>
                    Array.from(cell.childNodes).map(x => plain(x, depth)).join('').trim().replace(/\n+/g, ' ')).join('\t')).join('\n') + '\n';
            }
            const text = Array.from(node.childNodes).map(child => plain(child, depth)).join('');
            return block.has(tag) || tag === 'dt' || tag === 'dd' ? '\n' + text.trim() + '\n' : text;
        }
        const skeleton = Array.from(root.childNodes).map(x => plain(x)).join('').replace(/^[ \t]+\n/gm, '\n').trim();
        const text = skeleton.replace(new RegExp(marker + '(\\d+)' + marker, 'g'), (_, index) => literals[Number(index)]);
        const copy = root.cloneNode(true);
        const px = Number.isInteger(fontSize) && fontSize >= 6 && fontSize <= 16 ? fontSize : 12;
        copy.setAttribute('style', `font-family:Segoe UI,Arial,sans-serif;font-size:${px}px;line-height:1.6;color:#252423;`);
        for (const el of copy.querySelectorAll('*')) {
            el.removeAttribute('class');
            const tag = el.localName;
            let style = '';
            if (tag === 'table') style = 'border-collapse:collapse;margin:10px 0;';
            if (tag === 'th' || tag === 'td') style = cellStyle + (tag === 'th' ? 'font-weight:600;background-color:#f6f6f7;' : '') + (el.style.textAlign ? `text-align:${el.style.textAlign};` : '');
            if (tag === 'pre') style = 'white-space:pre-wrap;font-family:Consolas,monospace;padding:10px;background-color:#f6f6f7;border:1px solid #e7e5e3;';
            if (tag === 'code' || tag === 'kbd' || tag === 'samp') style = 'font-family:Consolas,monospace;';
            if (tag === 'p') style = 'margin:0 0 10px;';
            if (/^h[1-6]$/.test(tag)) style = `font-weight:600;font-size:${Math.round(px * ({h1:1.45,h2:1.25,h3:1.1}[tag] || 1))}px;margin:12px 0 6px;`;
            if (tag === 'blockquote') style = 'margin:10px 0;padding-left:10px;border-left:3px solid #d1cfcc;';
            if (tag === 'a') style = 'color:#b91312;text-decoration:underline;';
            if (tag === 'ul' || tag === 'ol') style = 'margin:8px 0;padding-left:24px;';
            if (style) el.setAttribute('style', style);
        }
        return { html: root.innerHTML, copyHtml: copy.outerHTML, text, images, removed };
    }
    window.redInkRendering = Object.freeze({ prepare });
})();
