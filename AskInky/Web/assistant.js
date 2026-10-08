(() => {
    'use strict';
    const ui = window.redInkUi;
    const source = document.getElementById('selected-text');
    const question = document.getElementById('question');
    const answer = document.getElementById('answer');
    const send = document.getElementById('send-button');
    const copy = document.getElementById('copy-button');
    const settings = document.getElementById('settings-button');
    let ready = false, busy = false, locked = false, fontSize = 12, rawHtml = '';
    let payload = { html: '', text: '' };
    function state() {
        const disabled = !ready || busy || locked;
        source.disabled = question.disabled = send.disabled = settings.disabled = disabled;
        copy.disabled = disabled || !rawHtml;
        document.body.classList.toggle('is-busy', busy);
        document.querySelector('main').setAttribute('aria-busy', String(busy));
    }
    function render() {
        if (!rawHtml) {
            answer.replaceChildren();
            const placeholder = document.createElement('p');
            placeholder.className = 'empty-answer';
            placeholder.textContent = ui.t('desktop.answerPlaceholder');
            answer.appendChild(placeholder);
            payload = { html: '', text: '' };
            return;
        }
        const prepared = window.redInkRendering.prepare(rawHtml, fontSize, ui.t('desktop.imagePlaceholder'));
        answer.innerHTML = prepared.html;
        payload = { html: prepared.copyHtml, text: prepared.text };
        if (prepared.images || prepared.removed) {
            ui.post('renderWarning', { images: prepared.images, removed: prepared.removed });
            ui.status(ui.t(prepared.images ? 'desktop.imagesBlocked' : 'desktop.contentFiltered'));
        }
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
                source.value = message.selectedText || '';
                question.value = message.question || '';
                busy = !!message.busy; locked = !!message.locked; ready = true;
                ui.status(message.status);
                preferences(message); state();
                if (!busy && !locked) question.focus();
                break;
            case 'preferences': preferences(message); break;
            case 'selection': source.value = message.selectedText || ''; ui.status(message.status); break;
            case 'answer': rawHtml = message.html || ''; ui.status(message.status); render(); state(); break;
            case 'busy': busy = !!message.busy; if (message.status !== undefined) ui.status(message.status); state(); break;
            case 'locked': locked = !!message.locked; state(); break;
            case 'error': busy = false; ui.status(message.message, true); state(); break;
            case 'status': ui.status(message.status); break;
            case 'focus': if (ready && !busy && !locked) question.focus(); break;
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
    window.redInkAnswer = Object.freeze({ getCopyPayload: () => ({ ...payload }) });
    state();
    if (ui.connected) ui.post('ready');
    else ui.status('This page is hosted by the Windows desktop application.', true);
})();
