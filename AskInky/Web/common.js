/* Local, versioned WebView2 message contract. No fetch, storage, remote script or host object. */
(() => {
    'use strict';
    let strings = Object.create(null);
    const listeners = new Set();
    const bridge = window.chrome && window.chrome.webview;
    const t = key => strings[key] || key;
    function post(action, payload = {}) {
        if (!bridge) return false;
        bridge.postMessage({ version: '1', action, ...payload });
        return true;
    }
    function translate() {
        for (const el of document.querySelectorAll('[data-text]')) el.textContent = t(el.dataset.text);
        for (const el of document.querySelectorAll('[data-title]')) {
            el.title = t(el.dataset.title);
            el.setAttribute('aria-label', t(el.dataset.title));
        }
        for (const el of document.querySelectorAll('[data-placeholder]')) el.placeholder = t(el.dataset.placeholder);
    }
    function preferences(message) {
        if (message.strings) strings = message.strings;
        if (message.language) document.documentElement.lang = message.language;
        if (typeof message.darkMode === 'boolean') document.documentElement.dataset.theme = message.darkMode ? 'dark' : 'light';
        if (Number.isInteger(message.fontSize) && message.fontSize >= 6 && message.fontSize <= 16) {
            document.documentElement.style.setProperty('--base-font-size', `${message.fontSize}px`);
        }
        translate();
    }
    function status(text, error = false) {
        const el = document.getElementById('status');
        el.textContent = text || '';
        el.classList.toggle('error', error);
    }
    function receive(message) {
        if (!message || typeof message !== 'object' || typeof message.type !== 'string') return;
        for (const fn of listeners) fn(message);
    }
    if (bridge) bridge.addEventListener('message', event => receive(event.data));
    window.redInkUi = Object.freeze({ t, post, preferences, translate, status, connected: !!bridge, on: fn => listeners.add(fn) });
    document.addEventListener('click', event => {
        const link = event.target.closest && event.target.closest('a[href]');
        if (!link) return;
        event.preventDefault();
        const href = link.getAttribute('href');
        if (href.startsWith('#')) {
            const target = document.getElementById(href.slice(1));
            if (target && document.getElementById('answer')?.contains(target)) target.scrollIntoView({ block: 'nearest' });
        } else if (/^(https?:|mailto:)/i.test(href)) post('openLink', { url: href });
    });
})();
