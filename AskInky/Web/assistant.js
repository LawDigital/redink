/* Part of Red Ink (Ask Inky). Copyright (c) LawDigital AG, Switzerland.
 * File: assistant.js
 * Purpose: Embedded desktop web user interface behavior.
 * License: Red Ink License; third-party attribution in embedded LICENSE.txt.
 */
(() => {
    'use strict';
    const ui = window.redInkUi;
    const source = document.getElementById('selected-text');
    const question = document.getElementById('question');
    const answer = document.getElementById('answer');
    const send = document.getElementById('send-button');
    const copy = document.getElementById('copy-button');
    const copySelection = document.getElementById('copy-selection-button');
    const copySource = document.getElementById('copy-source-button');
    const settings = document.getElementById('settings-button');
    const historyOption = document.getElementById('history-option');
    const navigation = document.getElementById('answer-navigation');
    const previous = document.getElementById('previous-answer');
    const next = document.getElementById('next-answer');
    const remove = document.getElementById('delete-answer');
    const clear = document.getElementById('clear-answers');
    const counter = document.getElementById('answer-count');
    const turnQuestion = document.getElementById('turn-question');
    const answerScroll = document.getElementById('answer-scroll');
    const historyToggle = document.getElementById('include-history');
    // P18 W367: options of newer Helpers, paste-back and file drop.
    const modeRow = document.getElementById('mode-row');
    // 2.0.3 (DARO UID 572): no agent mode in Ask Inky (quick, short requests only).
    const webOption = document.getElementById('web-option'), webMode = document.getElementById('web-mode');
    const thinkingOption = document.getElementById('thinking-option'), thinkingLevel = document.getElementById('thinking-level');
    const replace = document.getElementById('replace-button');
    const dropHint = document.getElementById('drop-hint');
    const fileOffer = document.getElementById('file-offer'), sendFile = document.getElementById('send-file-button');
    let modes = { modern: false, replaceEnabled: false };
    let ready = false, busy = false, locked = false, fontSize = 12, rawHtml = '';
    let payload = { html: '', text: '' };
    let selectedPayload = null;

    function selectedAnswerPayload() {
        const selection = window.getSelection();
        if (!selection || selection.isCollapsed || selection.rangeCount === 0) return null;
        const range = selection.getRangeAt(0);
        if (!answer.contains(range.startContainer) || !answer.contains(range.endContainer)) return null;
        const text = selection.toString();
        if (!text.trim() || text.length > 16000000) return null;
        const fragment = document.createElement('div');
        fragment.appendChild(range.cloneContents());
        const rendered = window.redInkRendering.prepare(fragment.innerHTML, fontSize, ui.t('desktop.imagePlaceholder'));
        if (!rendered.copyHtml || rendered.copyHtml.length > 16000000) return null;
        return { html: rendered.copyHtml, text };
    }

    function updateSelectedAnswer() {
        selectedPayload = rawHtml && ready && !busy && !locked ? selectedAnswerPayload() : null;
        copySelection.hidden = !selectedPayload;
        copySelection.disabled = !selectedPayload;
    }
    let turnIndex = -1, turnCount = 0;
    function state() {
        const disabled = !ready || busy || locked;
        source.disabled = question.disabled = send.disabled = settings.disabled = disabled;
        copy.disabled = disabled || !rawHtml;
        updateSelectedAnswer();
        copySource.disabled = disabled || !source.value;
        historyToggle.disabled = disabled;
        historyOption.hidden = turnCount === 0;
        navigation.hidden = false;
        previous.disabled = disabled || turnIndex <= 0;
        next.disabled = disabled || turnIndex < 0 || turnIndex >= turnCount - 1;
        remove.disabled = disabled || turnIndex < 0;
        clear.disabled = disabled || turnCount === 0;
        replace.hidden = !modes.replaceEnabled;
        replace.disabled = disabled || !rawHtml || !modes.replaceEnabled;
        webMode.disabled = thinkingLevel.disabled = sendFile.disabled = disabled;
        webOption.hidden = !(modes.modern && modes.webAvailable);
        thinkingOption.hidden = !(modes.modern && Array.isArray(modes.thinkingLevels) && modes.thinkingLevels.length > 0);
        document.body.classList.toggle('is-busy', busy);
        requestSize();
        document.querySelector('main').setAttribute('aria-busy', String(busy));
    }
    let sizeQueued = false;
    function requestSize() {
        if (sizeQueued) return;
        sizeQueued = true;
        requestAnimationFrame(() => {
            sizeQueued = false;
            if (!ready || !ui.connected) return;
            // Report the total client height needed to read the answer without scrolling.
            // The native window constrains it to the monitor work area.
            const deficit = Math.max(0, answerScroll.scrollHeight - answerScroll.clientHeight);
            const preferred = Math.ceil(window.innerHeight + deficit + 6);
            ui.post('preferredHeight', { height: preferred });
        });
    }
    function render() {
        selectedPayload = null;
        copySelection.hidden = true;
        copySelection.disabled = true;
        if (!rawHtml) {
            answer.replaceChildren();
            const placeholder = document.createElement('p');
            placeholder.className = 'empty-answer';
            placeholder.textContent = ui.t('desktop.answerPlaceholder');
            answer.appendChild(placeholder);
            payload = { html: '', text: '' };
            requestSize();
            return;
        }
        const prepared = window.redInkRendering.prepare(rawHtml, fontSize, ui.t('desktop.imagePlaceholder'));
        answer.innerHTML = prepared.html;
        payload = { html: prepared.copyHtml, text: prepared.text };
        requestSize();
        if (prepared.images || prepared.removed) {
            ui.post('renderWarning', { images: prepared.images, removed: prepared.removed });
            ui.status(ui.t(prepared.images ? 'desktop.imagesBlocked' : 'desktop.contentFiltered'));
        }
    }
    function setTurn(message) {
        turnIndex = Number.isInteger(message.turnIndex) ? message.turnIndex : -1;
        turnCount = Number.isInteger(message.turnCount) ? message.turnCount : 0;
        const label = message.turnQuestion || '';
        turnQuestion.textContent = label;
        turnQuestion.hidden = !label;
        requestSize();
        counter.textContent = turnCount ? `${turnIndex + 1} / ${turnCount}` : '';
        counter.title = ui.t('desktop.answerNavigationHint');
    }
    function preferences(message) {
        ui.preferences(message);
        if (message.status !== undefined) ui.status(message.status);
        if (message.fontSize) fontSize = message.fontSize;
        if (typeof message.shortcut === 'string') document.getElementById('shortcut').textContent = message.shortcut;
        render();
    }
    function submit() {
        if (!ready || busy || locked) return;
        if (!question.value.trim()) { ui.status(ui.t('desktop.invalidQuestion'), true); question.focus(); return; }
        busy = true;
        ui.status(ui.t('desktop.sending'));
        state();
        ui.post('send', { question: question.value });
    }
    ui.on(message => {
        switch (message.type) {
            case 'init':
                rawHtml = message.html || '';
                setTurn(message);
                historyToggle.checked = message.includeHistory !== false;
                source.value = message.selectedText || '';
                question.value = message.question || '';
                busy = !!message.busy; locked = !!message.locked; ready = true;
                ui.status(message.status);
                preferences(message); state();
                if (!busy && !locked) question.focus();
                break;
            case 'preferences': preferences(message); break;
            case 'ocrCapability': { const hint = document.getElementById('ocr-hint'); hint.hidden = !message.available; hint.textContent = message.available ? `OCR · ${message.shortcut || ''}` : ''; break; }
            case 'selection': source.value = message.selectedText || ''; ui.status(message.status); state(); break;
            case 'newConversation': source.value = message.selectedText || ''; rawHtml = ''; question.value = ''; setTurn({ turnIndex: -1, turnCount: 0 }); historyToggle.checked = message.includeHistory !== false; render(); state(); break;
            case 'answer': rawHtml = message.html || ''; setTurn(message); answerScroll.scrollTop = 0; question.value = message.question || ''; historyToggle.checked = message.includeHistory !== false; ui.status(message.status); render(); state(); break;
            case 'busy': busy = !!message.busy; if (message.status !== undefined) ui.status(message.status); state(); break;
            case 'locked': locked = !!message.locked; state(); break;
            // 2.0.4 (DARO UID 578): the selection is the path of a local file -> 'Datei senden'.
            case 'fileOffer': fileOffer.hidden = !message.name; document.getElementById('file-offer-text').textContent = message.name ? ui.t('desktop.fileOffer').replace('{name}', message.name) : ''; state(); break;
            case 'error': busy = false; ui.status(message.message, true); state(); break;
            case 'status': ui.status(message.status); break;
            case 'focus': if (ready && !busy && !locked) question.focus(); break;
            case 'modes': applyModes(message); break;
        }
    });
    function applyModes(message) {
        modes = message;
        modeRow.hidden = !message.modern;
        webMode.checked = !!message.web;
        const levels = Array.isArray(message.thinkingLevels) ? message.thinkingLevels : [];
        // 2.0.3: exactly the levels the Helper offers for this model; no 'Standard' entry.
        thinkingLevel.replaceChildren();
        for (const level of levels) thinkingLevel.add(new Option(ui.t('desktop.thinking.' + level) === 'desktop.thinking.' + level ? level : ui.t('desktop.thinking.' + level), level));
        if (levels.length) thinkingLevel.value = levels.includes(message.thinking) ? message.thinking : levels[0];
        state();
    }
    function modeChanged() {
        if (!ready || busy || locked) return;
        ui.post('modeChanged', { agent: false, web: webMode.checked, thinking: thinkingLevel.value });
        state();
    }
    webMode.addEventListener('change', modeChanged);
    thinkingLevel.addEventListener('change', modeChanged);
    replace.addEventListener('click', () => { if (!replace.disabled) ui.post('replace'); });
    sendFile.addEventListener('click', () => { if (!sendFile.disabled) { fileOffer.hidden = true; ui.post('sendFile'); } });
    // A file dropped on the window: only the WebView2 file object goes to the host (no content in JS).
    let dragDepth = 0;
    document.addEventListener('dragenter', event => { if (event.dataTransfer && Array.from(event.dataTransfer.types || []).includes('Files')) { dragDepth++; dropHint.hidden = !ready || busy || locked; } });
    document.addEventListener('dragleave', () => { dragDepth = Math.max(0, dragDepth - 1); if (!dragDepth) dropHint.hidden = true; });
    document.addEventListener('dragover', event => { event.preventDefault(); if (event.dataTransfer) event.dataTransfer.dropEffect = 'copy'; });
    document.addEventListener('drop', event => {
        event.preventDefault();
        dragDepth = 0; dropHint.hidden = true;
        const files = event.dataTransfer ? Array.from(event.dataTransfer.files || []) : [];
        if (!files.length || !ready || busy || locked) return;
        const bridge = window.chrome && window.chrome.webview;
        if (bridge && typeof bridge.postMessageWithAdditionalObjects === 'function') {
            bridge.postMessageWithAdditionalObjects({ version: '1', action: 'dropFile' }, [files[0]]);
        }
    });
    question.addEventListener('input', () => ui.post('questionChanged', { question: question.value }));
    question.addEventListener('keydown', event => {
        if (event.key === 'Enter' && !event.shiftKey && !event.isComposing && event.keyCode !== 229) {
            event.preventDefault(); submit();
        }
    });
    send.addEventListener('click', submit);
    settings.addEventListener('click', () => { if (!settings.disabled) ui.post('settings'); });
    copy.addEventListener('click', () => { if (!copy.disabled) ui.post('copy'); });
    // Preserve an answer selection when clicking the copy icon (browser focus changes).
    copySelection.addEventListener('pointerdown', event => { if (!copySelection.disabled) event.preventDefault(); });
    copySelection.addEventListener('click', () => { if (!copySelection.disabled && selectedPayload) ui.post('copySelection'); });
    document.addEventListener('selectionchange', updateSelectedAnswer);
    copySource.addEventListener('click', () => { if (!copySource.disabled) ui.post('copySource'); });
    historyToggle.addEventListener('change', () => { if (!historyToggle.disabled) ui.post('historyChanged', { includeHistory: historyToggle.checked }); });
    function navigate(direction) {
        if (busy || locked || !ready) return;
        if (direction === 'previous' && turnIndex > 0) ui.post('navigateAnswer', { direction });
        if (direction === 'next' && turnIndex >= 0 && turnIndex < turnCount - 1) ui.post('navigateAnswer', { direction });
    }
    previous.addEventListener('click', () => navigate('previous'));
    next.addEventListener('click', () => navigate('next'));
    remove.addEventListener('click', () => { if (!remove.disabled) ui.post('deleteAnswer'); });
    clear.addEventListener('click', () => {
        if (!clear.disabled && window.confirm(ui.t('desktop.confirmClearAnswers'))) ui.post('clearAnswers');
    });
    document.addEventListener('keydown', event => {
        if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
        if (event.ctrlKey || event.metaKey || event.shiftKey) return;
        const target = event.target;
        const editing = target && (target.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(target.tagName));
        if (editing || busy || locked || turnCount < 2) return;
        if (!event.altKey && !answerScroll.contains(target) && !navigation.contains(target) && target !== document.body) return;
        event.preventDefault();
        navigate(event.key === 'ArrowLeft' ? 'previous' : 'next');
    });
    // Native window resizing must not request a new height, or it creates a feedback loop.
    // Content changes still schedule the desired height through requestSize().
    window.redInkAnswer = Object.freeze({ getCopyPayload: () => ({ ...payload }),
        getSelectedCopyPayload: () => selectedPayload ? { ...selectedPayload } : null });
    state();
    if (ui.connected) ui.post('ready');
    else ui.status('This page is hosted by the Windows desktop application.', true);
})();
