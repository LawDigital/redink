#!/usr/bin/env python3
"""Embedded browser UI regression tests (Chromium simulation, not Windows/WebView2)."""
import os
import pathlib
import shutil
from playwright.sync_api import sync_playwright

root = pathlib.Path(__file__).resolve().parents[1]
with sync_playwright() as pw:
    executable = shutil.which('chromium') or shutil.which('google-chrome')
    launch = {'headless': True, 'args': ['--no-sandbox']}
    if executable:
        launch['executable_path'] = executable
    browser = pw.chromium.launch(**launch)
    page = browser.new_page(viewport={'width': 760, 'height': 540}, bypass_csp=True)
    page.set_content((root / 'Web/picker.html').read_text(encoding='utf-8'))
    page.add_style_tag(content=(root / 'Web/picker.css').read_text(encoding='utf-8'))
    page.evaluate('''() => {
      window.__posted = [];
      window.chrome = {webview:{
        postMessage: value => window.__posted.push(value),
        addEventListener: (name, cb) => { if (name === 'message') window.__onmodel = cb; }
      }};
    }''')
    page.add_script_tag(content=(root / 'Web/picker.js').read_text(encoding='utf-8'))
    fixtures = [
      {'Id':0,'Title':'Textkorrektur','Prompt':'Text korrigieren','Category':'Allgemein','Source':'central'},
      {'Id':1,'Title':'Brief (local)','Prompt':'Freundlicher Brief','Category':'Korrespondenz','Source':'local'},
      {'Id':2,'Title':'Abschluss','Prompt':'Fazit formulieren','Category':'Allgemein','Source':'central'},
    ]
    def model(selection=False, local=True, preview=55):
        page.evaluate('arg => window.__onmodel({data:arg})', {
            'type':'model','items':fixtures,'hasSelection':selection,
            'hasLocal':local,'canSave':selection and local,'previewPercent':preview,
        })
    def posted():
        return page.evaluate('window.__posted')
    def clear():
        page.evaluate('window.__posted = []')

    model(False, False)
    assert page.title() == 'Red Ink Prompt Picker'
    assert page.locator('header, .intro, .brand').count() == 0
    assert page.locator('#search').evaluate('(el) => document.activeElement === el')
    assert page.locator('#search').get_attribute('placeholder') == 'Prompts durchsuchen'
    assert page.locator('#saveSelection').is_disabled()
    assert page.locator('.item').count() == 3
    assert page.locator('#saveHint').evaluate('(el) => el.getBoundingClientRect().width <= 1')

    # Arrow keys move the selection while typing focus stays inside the filter.
    assert page.locator('#previewTitle').inner_text() == 'Textkorrektur'
    page.locator('#search').press('ArrowDown')
    assert page.locator('#search').evaluate('(el) => document.activeElement === el')
    assert page.locator('#previewTitle').inner_text() == 'Brief (local)'
    page.locator('#search').press('ArrowUp')
    assert page.locator('#previewTitle').inner_text() == 'Textkorrektur'
    page.locator('#search').press('ArrowDown')
    page.locator('#search').press('Enter')
    assert posted() == [{'type':'select','id':1}], posted()
    clear()
    model()

    # Pointer double click commits once, including underlying pointerup handling.
    page.locator('.item').nth(2).dblclick(delay=105)
    assert posted() == [{'type':'select','id':2}], posted()
    clear()
    model()

    # Clicking once then Enter inserts the selected row; single click never inserts.
    page.locator('.item').nth(1).click()
    assert posted() == []
    assert page.locator('#previewTitle').inner_text() == 'Brief (local)'
    page.locator('#search').focus()
    page.locator('#search').press('Enter')
    assert posted() == [{'type':'select','id':1}], posted()
    clear()

    model(selection=True)
    page.locator('#search').fill('korrekt')
    assert page.locator('.item').count() == 1
    page.locator('#search').press('Enter')
    assert posted() == [{'type':'select','id':0}], posted()
    clear()
    model(selection=True)
    page.locator('#saveSelection').click()
    assert posted()[-1]['type'] == 'save'
    clear()
    page.evaluate('window.dispatchEvent(new Event("focus"))')
    page.wait_for_timeout(50)
    assert page.locator('#search').evaluate('(el) => document.activeElement === el')

    # Initial preview percentage + pointer drag + keyboard width adjustment.
    model(preview=60)
    assert page.locator('#splitter').get_attribute('aria-valuenow') == '60'
    original_width = page.locator('.previewPane').bounding_box()['width']
    splitter = page.locator('#splitter')
    rect = splitter.bounding_box()
    page.mouse.move(rect['x'] + rect['width'] / 2, rect['y'] + rect['height'] / 2)
    page.mouse.down()
    page.mouse.move(rect['x'] - 65, rect['y'] + rect['height'] / 2, steps=5)
    page.mouse.up()
    assert page.locator('.previewPane').bounding_box()['width'] > original_width + 45
    signals = [x for x in posted() if x.get('type') == 'split']
    assert len(signals) >= 1 and 32 <= signals[-1]['previewPercent'] <= 68, signals
    clear()
    splitter.focus()
    before = int(splitter.get_attribute('aria-valuenow'))
    splitter.press('ArrowRight')
    after = int(splitter.get_attribute('aria-valuenow'))
    assert after == max(32, before - 2) and posted()[-1] == {'type':'split','previewPercent':after}, posted()
    clear()

    # Search-focus changes must not silently alter user-selected splitter geometry.
    model(preview=after)
    assert int(splitter.get_attribute('aria-valuenow')) == after
    assert page.locator('#search').evaluate('(el) => document.activeElement === el')
    for width, height in [(600,390),(760,540),(900,700),(525,320)]:
        page.set_viewport_size({'width':width,'height':height})
        dims = page.evaluate('''() => ({scrollW:document.documentElement.scrollWidth,
           width:innerWidth, scrollH:document.body.scrollHeight, height:innerHeight,
           preview:document.querySelector('.previewPane').getBoundingClientRect().width,
           list:document.querySelector('.results').clientHeight})''')
        assert dims['scrollW'] <= dims['width'] and dims['scrollH'] <= dims['height'], dims
        assert dims['preview'] >= 180 and dims['list'] >= 180, dims
    if os.environ.get('PROMPT_PICKER_QA_SCREENSHOT'):
        page.set_viewport_size({'width':760,'height':540})
        model(preview=55)
        page.screenshot(path=os.environ['PROMPT_PICKER_QA_SCREENSHOT'])
    browser.close()
    print('PASS Chromium: keyboard filter/arrows/Enter; doubleclick once; click + Enter; splitter drag/keyboard/persistence messages; responsive layouts; no local path write controls')
