/* Part of Red Ink (Ask Inky). Copyright (c) LawDigital Ltd., Switzerland.
 * File: settings.js
 * Purpose: Embedded desktop web user interface behavior.
 * License: Red Ink License; third-party attribution in embedded LICENSE.txt.
 */
(() => {
    'use strict';
    const ui = window.redInkUi;
    const $ = id => document.getElementById(id);
    let ready = false, busy = false, recording = false, recordingOcr = false, apiChanged = false, registered = false;
    let registeredUser = '', savedOrigin = '', language = 'de', bundles = {}, probe = null, registryManaged = false;
    const editable = Array.from(document.querySelectorAll('input, textarea, select, button'));
    let savedModel = '';
    const keyOptions = [];
    for (let k = 65; k <= 90; k++) keyOptions.push([k, String.fromCharCode(k)]);
    for (let k = 48; k <= 57; k++) keyOptions.push([k, String.fromCharCode(k)]);
    for (let n = 1; n <= 24; n++) if (n !== 12) keyOptions.push([111 + n, `F${n}`]);
    for (const [key, text] of keyOptions) { $('hotkey-key').add(new Option(text, String(key))); $('ocr-hotkey-key').add(new Option(text, String(key))); }
    function rebuildModels(models, error) {
        const selector = $('settings-model');
        const wanted = modelValue();
        selector.replaceChildren();
        selector.add(new Option(ui.t('desktop.defaultModel'), ''));
        const seen = new Set(['']);
        for (const entry of Array.isArray(models) ? models : []) {
            const id = String(entry.id || '').trim();
            if (!id || id.length > 256 || seen.has(id)) continue;
            seen.add(id);
            const label = String(entry.name || id).trim();
            selector.add(new Option(label && label !== id ? `${label} (${id})` : id, id));
        }
        selector.add(new Option(ui.t('desktop.customModel'), '__custom__'));
        if (wanted && !seen.has(wanted)) {
            selector.value = '__custom__';
            $('settings-model-custom').value = wanted;
        } else selector.value = wanted;
        showCustomModel();
        $('model-status').textContent = error ? ui.t('desktop.modelsUnavailable') : ui.t('desktop.modelsLoaded');
    }
    function modelValue() {
        return $('settings-model').value === '__custom__' ? $('settings-model-custom').value.trim() : $('settings-model').value;
    }
    function showCustomModel() {
        $('settings-model-custom').hidden = $('settings-model').value !== '__custom__';
    }
    function ocrShortcut() {
        return { modifiers: ($('ocr-hotkey-ctrl').checked ? 2 : 0) | ($('ocr-hotkey-shift').checked ? 4 : 0) | ($('ocr-hotkey-alt').checked ? 1 : 0), key: Number($('ocr-hotkey-key').value) };
    }
    function ocrShortcutText() {
        const {modifiers, key} = ocrShortcut();
        return [(modifiers & 2) ? 'Ctrl' : '', (modifiers & 4) ? 'Shift' : '', (modifiers & 1) ? 'Alt' : '', keyOptions.find(x => x[0] === key)?.[1]].filter(Boolean).join(' + ');
    }
    function setOcrShortcut(modifiers, key) {
        $('ocr-hotkey-ctrl').checked = !!(modifiers & 2);
        $('ocr-hotkey-shift').checked = !!(modifiers & 4);
        $('ocr-hotkey-alt').checked = !!(modifiers & 1);
        $('ocr-hotkey-key').value = String(key);
        $('ocr-hotkey-display').value = ocrShortcutText();
    }
    function origin(raw) {
        try {
            let text = raw.trim() || 'http://localhost';
            if (!/^[a-z][a-z0-9+.-]*:\/\//i.test(text)) {
                const host = new URL('http://' + text).hostname.toLowerCase();
                const local = host === 'localhost' || host === '[::1]' || /^127\.(\d+\.){2}\d+$/.test(host);
                text = (local ? 'http://' : 'https://') + text;
            }
            return new URL(text).origin;
        } catch { return ''; }
    }
    function isLocal(raw) {
        try {
            const host = new URL(origin(raw)).hostname.toLowerCase();
            return host === 'localhost' || host === '[::1]' || /^127\.(\d+\.){2}\d+$/.test(host);
        } catch { return false; }
    }
    function shortcut() {
        const modifiers = ($('hotkey-ctrl').checked ? 2 : 0) | ($('hotkey-shift').checked ? 4 : 0) | ($('hotkey-alt').checked ? 1 : 0);
        return { modifiers, key: Number($('hotkey-key').value) };
    }
    function shortcutText() {
        const { modifiers, key } = shortcut();
        return [(modifiers & 2) ? 'Ctrl' : '', (modifiers & 4) ? 'Shift' : '', (modifiers & 1) ? 'Alt' : '', keyOptions.find(x => x[0] === key)?.[1]].filter(Boolean).join(' + ');
    }
    function setShortcut(modifiers, key) {
        $('hotkey-ctrl').checked = !!(modifiers & 2);
        $('hotkey-shift').checked = !!(modifiers & 4);
        $('hotkey-alt').checked = !!(modifiers & 1);
        $('hotkey-key').value = String(key);
        $('hotkey-display').value = shortcutText();
    }
    function setRecording(active) {
        recording = active && ready && !busy;
        $('hotkey-display').classList.toggle('recording', recording);
        $('hotkey-display').value = recording ? ui.t('desktop.recording') : shortcutText();
        ui.post('recording', { active: recording, target: 'text' });
    }
    function state() {
        for (const el of editable) el.disabled = !ready || busy;
        const locked = registryManaged && !$('settings-ignore-registry-link').checked;
        for (const id of ['settings-helper-url', 'settings-helper-apikey', 'settings-license-user', 'settings-registration-link']) $(id).disabled = !ready || busy || locked;
        $('btn-register-device').disabled = !ready || busy || locked || !$('settings-registration-link').value.trim();
        $('btn-unregister-device').disabled = !ready || busy || locked || !registered;
        $('settings-registry-managed-hint').hidden = !locked;
        document.body.classList.toggle('is-busy', busy);
        document.querySelector('main').setAttribute('aria-busy', String(busy));
    }
    function licenseText() {
        return probe && probe.licenseKey ? ui.t(probe.licenseKey).replace('{product}', probe.licenseProduct || '') : ui.t('desktop.licenseUnknown');
    }
    function connectionStatus() {
        $('connection-state').removeAttribute('data-text');
        $('connection-dot').className = 'connection-dot';
        if (probe) {
            $('connection-state').textContent = ui.t('desktop.connected') + (probe.version ? ' · ' + probe.version : '');
            $('connection-dot').classList.add(probe.authorized ? 'connected' : 'failed');
            $('license-status').textContent = licenseText();
        } else {
            $('connection-state').textContent = ui.t('desktop.notTested');
            $('license-status').textContent = ui.t('desktop.licenseUnknown');
        }
    }
    function registrationStatus() {
        const changed = origin($('settings-helper-url').value) !== savedOrigin;
        if (changed && savedOrigin) $('registration-state').textContent = ui.t('desktop.serverChanged');
        else if (registered) $('registration-state').textContent = ui.t('settings.deviceRegisteredAs').replace('{user}', registeredUser);
        else $('registration-state').textContent = ui.t(isLocal($('settings-helper-url').value) ? 'desktop.localRegistration' : 'settings.deviceNotRegistered');
    }
    function preview() {
        language = $('language-select').value;
        ui.preferences({ strings: bundles[language], language, darkMode: $('dark-mode-toggle').checked, fontSize: Number($('font-size').value) });
        $('font-size-value').textContent = `${$('font-size').value} px`;
        $('temperature-value').textContent = (Number($('temperature').value) / 100).toLocaleString(language, { maximumFractionDigits: 2 });
        $('settings-llm-debug-panel').hidden = !$('settings-show-llm-debug').checked;
        $('hotkey-display').value = recording ? ui.t('desktop.recording') : shortcutText();
        $('ocr-hotkey-display').value = recordingOcr ? ui.t('desktop.recording') : ocrShortcutText();
        connectionStatus(); registrationStatus();
        ui.post('preview', { language });
    }
    function values() {
        const key = shortcut();
        return {
            helperUrl: $('settings-helper-url').value, apiKey: $('settings-helper-apikey').value, apiKeyChanged: apiChanged,
            licenseUser: $('settings-license-user').value, language: $('language-select').value,
            ignoreRegistryHelperLink: $('settings-ignore-registry-link').checked,
            darkMode: $('dark-mode-toggle').checked, fontSize: Number($('font-size').value), temperature: Number($('temperature').value) / 100,
            model: modelValue(), systemPrompt: $('system-prompt').value, showLlmDebug: $('settings-show-llm-debug').checked,
            hotkeyModifiers: key.modifiers, hotkeyKey: key.key, ocrHotkeyModifiers: ocrShortcut().modifiers, ocrHotkeyKey: ocrShortcut().key
        };
    }
    function commit(action) {
        if (!ready || busy) return;
        const candidate = values();
        if (!(candidate.hotkeyModifiers & 3) || !keyOptions.some(x => x[0] === candidate.hotkeyKey)) {
            ui.status(ui.t('desktop.invalidShortcut'), true); return;
        }
        if (!(candidate.ocrHotkeyModifiers & 3) || !keyOptions.some(x => x[0] === candidate.ocrHotkeyKey) || (candidate.ocrHotkeyModifiers === candidate.hotkeyModifiers && candidate.ocrHotkeyKey === candidate.hotkeyKey)) { ui.status(ui.t('desktop.invalidShortcut'), true); return; }
        if (action === 'register' && !$('settings-registration-link').value.trim()) {
            ui.status(ui.t('settings.registrationLinkMissing'), true); return;
        }
        setRecording(false);
        busy = true; state();
        ui.status(action === 'test' ? ui.t('desktop.testing') : action === 'register' ? ui.t('settings.registrationInProgress') : '');
        ui.post(action, { values: candidate, ...(action === 'register' ? { link: $('settings-registration-link').value.trim() } : {}) });
    }
    ui.on(message => {
        switch (message.type) {
            case 'settings': {
                const v = message.values;
                bundles = message.bundles; language = message.language;
                ui.preferences(message);
                $('language-select').value = language;
                $('dark-mode-toggle').checked = v.darkMode;
                $('settings-helper-url').value = v.helperUrl;
                $('settings-helper-apikey').value = v.apiKey;
                $('settings-helper-apikey').placeholder = v.apiKeyReadable ? '' : ui.t('desktop.keyUnreadable');
                $('settings-license-user').value = v.licenseUser;
                $('font-size').value = v.fontSize;
                $('temperature').value = Math.round(v.temperature * 100);
                $('system-prompt').value = v.systemPrompt;
                savedModel = v.model || '';
                rebuildModels([], false);
                if (savedModel) { $('settings-model').value = '__custom__'; $('settings-model-custom').value = savedModel; showCustomModel(); }
                $('settings-show-llm-debug').checked = v.showLlmDebug;
                $('settings-ignore-registry-link').checked = !!v.ignoreRegistryHelperLink;
                registryManaged = !!v.registryManaged;
                $('debug-sent').textContent = message.debugSent || '';
                $('debug-received').textContent = message.debugReceived || '';
                registered = v.registered; registeredUser = v.userId; savedOrigin = origin(v.helperUrl);
                apiChanged = false; recording = false; probe = null;
                setShortcut(v.hotkeyModifiers, v.hotkeyKey);
                setOcrShortcut(v.ocrHotkeyModifiers, v.ocrHotkeyKey);
                ready = true; preview(); state();
                $('registration-setup-help').hidden = !message.needsRegistration;
                $('settings-registration-link').classList.toggle('registration-needs-input', !!message.needsRegistration);
                if (message.needsRegistration) {
                    const linkField = $('settings-registration-link');
                    linkField.scrollIntoView({ block: 'center' });
                    linkField.focus({ preventScroll: true });
                } else ui.post('loadModels', { values: values() });
                ui.status(message.feedback || (message.needsRegistration ? ui.t('settings.initialSetupInstruction') : ''), !!message.failure);
                break;
            }
            case 'busy': busy = !!message.busy; state(); break;
            case 'error': busy = false; ui.status(message.message, true); state(); break;
            case 'probe': probe = message; connectionStatus(); ui.status(message.authorized ? ui.t('desktop.connected') : licenseText(), !message.authorized); break;
            case 'models': rebuildModels(message.models, !!message.error); break;
            case 'clearInvite': $('settings-registration-link').value = ''; state(); break;
            case 'hotkey': if (message.ocr) { setOcrShortcut(message.modifiers, message.key); recordingOcr = false; $('ocr-hotkey-display').blur(); } else { setShortcut(message.modifiers, message.key); setRecording(false); $('hotkey-display').blur(); } break;
        }
    });
    $('hotkey-display').addEventListener('focus', () => setRecording(true));
    $('hotkey-display').addEventListener('click', () => { if (!recording) setRecording(true); });
    $('hotkey-display').addEventListener('blur', () => setRecording(false));
    $('hotkey-display').addEventListener('keydown', event => {
        if (!recording) return;
        if (event.key === 'Tab') { setRecording(false); return; }
        event.preventDefault();
        if (event.key === 'Escape') { setRecording(false); event.target.blur(); return; }
        if (['Control', 'Shift', 'Alt', 'Meta'].includes(event.key)) return;
        let key = 0;
        if (/^[a-z]$/i.test(event.key)) key = event.key.toUpperCase().charCodeAt(0);
        else if (/^Digit[0-9]$/.test(event.code)) key = event.code.charCodeAt(5);
        else if (/^F\d{1,2}$/.test(event.code)) key = 111 + Number(event.code.slice(1));
        const modifiers = (event.ctrlKey ? 2 : 0) | (event.shiftKey ? 4 : 0) | (event.altKey ? 1 : 0);
        if (event.metaKey || !(modifiers & 3) || !keyOptions.some(x => x[0] === key)) { ui.status(ui.t('desktop.invalidShortcut'), true); return; }
        setShortcut(modifiers, key); setRecording(false); event.target.blur(); ui.status('');
    });
    $('ocr-hotkey-display').addEventListener('focus', () => { recordingOcr = true; $('ocr-hotkey-display').value = ui.t('desktop.recording'); ui.post('recording', {active: true, target: 'ocr'}); });
    $('ocr-hotkey-display').addEventListener('blur', () => { recordingOcr = false; $('ocr-hotkey-display').value = ocrShortcutText(); ui.post('recording', {active: false, target: 'ocr'}); });
    $('ocr-hotkey-display').addEventListener('keydown', event => {
        if (!recordingOcr) return;
        if (event.key === 'Tab') return;
        event.preventDefault();
        if (event.key === 'Escape') { event.target.blur(); return; }
        if (['Control', 'Shift', 'Alt', 'Meta'].includes(event.key)) return;
        let key = 0;
        if (/^[a-z]$/i.test(event.key)) key = event.key.toUpperCase().charCodeAt(0);
        else if (/^Digit[0-9]$/.test(event.code)) key = event.code.charCodeAt(5);
        else if (/^F\d{1,2}$/.test(event.code)) key = 111 + Number(event.code.slice(1));
        const modifiers = (event.ctrlKey ? 2 : 0) | (event.shiftKey ? 4 : 0) | (event.altKey ? 1 : 0);
        if (event.metaKey || !(modifiers & 3) || !keyOptions.some(x => x[0] === key)) { ui.status(ui.t('desktop.invalidShortcut'), true); return; }
        setOcrShortcut(modifiers, key); recordingOcr = false; event.target.blur();
    });
    $('reset-ocr-hotkey').addEventListener('click', () => setOcrShortcut(3, 79));
    for (const id of ['ocr-hotkey-ctrl','ocr-hotkey-shift','ocr-hotkey-alt','ocr-hotkey-key']) $(id).addEventListener('change', () => { recordingOcr = false; $('ocr-hotkey-display').value = ocrShortcutText(); });
    $('reset-hotkey').addEventListener('click', () => { setRecording(false); setShortcut(3, 75); ui.status(''); });
    for (const id of ['hotkey-ctrl', 'hotkey-shift', 'hotkey-alt', 'hotkey-key']) $(id).addEventListener('change', () => { setRecording(false); $('hotkey-display').value = shortcutText(); });
    for (const id of ['language-select', 'dark-mode-toggle', 'settings-show-llm-debug']) $(id).addEventListener('change', preview);
    for (const id of ['font-size', 'temperature']) $(id).addEventListener('input', preview);
    $('settings-model').addEventListener('change', showCustomModel);
    $('settings-helper-apikey').addEventListener('input', () => { apiChanged = true; probe = null; connectionStatus(); });
    $('settings-helper-url').addEventListener('input', () => { probe = null; connectionStatus(); registrationStatus(); });
    $('settings-registration-link').addEventListener('input', state);
    $('settings-ignore-registry-link').addEventListener('change', () => { state(); connectionStatus(); registrationStatus(); });
    $('settings-helper-url').addEventListener('change', () => { if (ready && !busy) ui.post('loadModels', { values: values() }); });
    $('save-settings').addEventListener('click', () => commit('save'));
    $('btn-test-connection').addEventListener('click', () => commit('test'));
    $('btn-register-device').addEventListener('click', () => commit('register'));
    $('btn-unregister-device').addEventListener('click', () => commit('unregister'));
    $('company-website').addEventListener('click', event => { event.preventDefault(); ui.post('openWebsite'); });
    $('third-party-licenses').addEventListener('click', event => { event.preventDefault(); ui.post('openLicenses'); });
    $('cancel-settings').addEventListener('click', () => { if (!busy) ui.post('cancel'); });
    state();
    if (ui.connected) ui.post('ready');
    else ui.status('This page is hosted by the Windows desktop application.', true);
})();
