#!/usr/bin/env python3
"""Source/integration checks. No claim of Windows compilation or runtime testing."""
from pathlib import Path
import subprocess
import xml.etree.ElementTree as ET

root = Path(__file__).resolve().parents[1]
folder = root.parent

def verify(value, message):
    assert value, message
    print('PASS', message)

project = ET.parse(root / 'PromptPicker.vbproj').getroot()
attrs = {x.tag: x.text for x in project.iter() if x.tag in ('TargetFramework', 'UseWindowsForms', 'RootNamespace', 'StartupObject')}
verify(attrs == {'TargetFramework':'net8.0-windows','UseWindowsForms':'true','RootNamespace':'PromptPicker','StartupObject':'PromptPicker.Program'}, 'VB.NET Windows Forms net8 project attributes')
verify(any(x.attrib.get('Include') == 'Microsoft.Web.WebView2' for x in project.iter('PackageReference')), 'WebView2 dependency')
for emb in project.iter('EmbeddedResource'):
    path = root / emb.attrib['Include'].replace('\\', '/')
    verify(path.exists(), f'Embedded resource exists: {path.relative_to(root)}')

reader=(root/'WordConfigReader.vb').read_text()
writer=(root/'PromptLibraryStore.vb').read_text()
context=(root/'PickerContext.vb').read_text()
services=(root/'SelectionClipboardService.vb').read_text()
verify('"%AppData%\\Microsoft\\Word\\redink.ini"' in reader, 'Word default configuration path')
verify('"Software\\Red Ink"' in reader and '"IniPath"' in reader, 'Registry directory fallback')
verify(reader.index('If System.IO.File.Exists(wordDefault)') < reader.index('If Not System.String.IsNullOrWhiteSpace(regDirectory)'), 'Word config priority before registry')
verify('RedInkPathExpander.Expand(central)' in reader and 'RedInkPathExpander.Expand(local)' in reader, 'Word library path expansion delegates to Red Ink equivalent')
verify('dict.TryGetValue("PromptLib", central)' in reader and 'dict.TryGetValue("PromptLibLocal", local)' in reader, 'INI keys are read')
verify(all(x not in reader for x in ('WriteAllText', 'AppendAllText', 'SetValue(', 'CreateDirectory', 'File.Move')), 'INI resolver never writes')
verify(all(marker in (root/'RedInkPathExpander.vb').read_text() for marker in ['%DOCUMENTS%', '%DOWNLOADS%', '%APPSTARTUPPATH%', 'Path.IsPathRooted', 'expanded.Trim']), 'Red Ink path aliases and quotes are supported')
verify('If Not config.HasLocal Then Return False' in writer and 'If Not CanSaveLocal(config) Then Throw' in writer, 'Local path absent: save impossible')
verify('CanonicalFilePath(config.LocalPath)' in writer and 'CanonicalFilePath(config.CentralPath)' in writer, 'Central/local path alias guard')
verify('System.IO.File.AppendAllText(config.LocalPath' in writer and 'System.IO.Directory.CreateDirectory(folder)' in writer, 'Only configured local path written')
verify('AppendAllText(config.CentralPath' not in writer and 'WriteAllText(config.CentralPath' not in writer, 'Central library never written')
verify('ReadFile(config.LocalPath' in writer and 'ReadFile(config.CentralPath' in writer and writer.index('ReadFile(config.LocalPath') < writer.index('ReadFile(config.CentralPath'), 'Local-first loading')
verify('" (local)"' in writer and '"(Uncategorized)"' in writer, 'Title suffix and uncategorized category')
verify('Replace(vbCrLf, "\\n")' in writer and '.Replace("|", "¦")' in writer, 'Prompt format newlines and separators')
verify('Native.GetForegroundWindow() <> window Then Return False' in services and 'Native.SetForegroundWindow(window)' in services, 'Insertion verifies original foreground window')
verify('GetClipboardSequenceNumber() = pastedSequence' in services, 'Clipboard only restored if unchanged')
verify('Ctrl+Alt+P' in (root/'README.md').read_text() and 'Current' not in (root/'README.md').read_text()[:100], 'Shortcut docs')
verify('Native.MOD_CONTROL Or Native.MOD_ALT' in (root/'PickerSettings.vb').read_text() and 'Keys.P' in (root/'PickerSettings.vb').read_text(), 'Hotkey differs from Ask Inky')
verify('If Not IsExternal(target)' in context and 'Await WordConfigReader.LoadAsync()' in context, 'Original app targeting / runtime INI read')
verify('New SavePromptForm(selectedText' in (root/'PickerForm.vb').read_text(), 'Capture to local prompt dialog')
verify('ShowNativeFallback()' in (root/'PickerForm.vb').read_text(), 'Native fallback available')
verify('Content-Security-Policy' in (root/'Web/picker.html').read_text() and 'default-src \'none\'' in (root/'Web/picker.html').read_text(), 'Webview CSP')
verify('textContent =' in (root/'Web/picker.js').read_text() and 'innerHTML' not in (root/'Web/picker.js').read_text(), 'Browser text-only rendering of library values')
result=subprocess.run(['node','--check',str(root/'Web/picker.js')],capture_output=True,text=True)
verify(result.returncode == 0, 'JS syntax parse: '+(result.stderr or 'OK'))
if (folder/'Red Ink.sln').exists():
    sln=(folder/'Red Ink.sln').read_text(encoding='utf-8-sig')
    verify('PromptPicker\\PromptPicker.vbproj' in sln and sln.count('E7C890ED-5E95-4878-8CF1-F1EB10D85E9B') == 9, 'Solution project/config entries')

# UI and input regression contracts. No reliance on the rendering technology
# to enforce write access: the store remains the single authority.
picker=(root/'PickerForm.vb').read_text()
html=(root/'Web/picker.html').read_text()
js=(root/'Web/picker.js').read_text()
css=(root/'Web/picker.css').read_text()
native=(root/'Native.vb').read_text()
save=(root/'SavePromptForm.vb').read_text()
shortcut=(root/'ShortcutSettingsForm.vb').read_text()
ui=(root/'PickerUi.vb').read_text()
settings=(root/'PickerSettings.vb').read_text()
verify('PickerClientWidth As Integer = 1140' in settings and 'PickerClientHeight As Integer = 540' in settings, 'Default width enlarged 50 percent from 760 to 1140 while height stays 540')
verify('CurrentPickerLayoutVersion As Integer = 2' in settings and 'PickerLayoutVersion As Integer = CurrentPickerLayoutVersion' in settings, 'New picker settings carry a layout migration marker')
verify('TryGetProperty(NameOf(PickerSettings.PickerLayoutVersion)' in settings and 'isLegacy AndAlso data.PickerClientWidth = 760' in settings and 'data.PickerClientWidth = 1140' in settings, 'Migrate only legacy 760px default, not customised widths')
verify('data.PickerLayoutVersion = PickerSettings.CurrentPickerLayoutVersion' in settings and 'If Not System.IO.File.Exists(ConfigPath) Then Return New PickerSettings()' in settings, 'Migration is one-time and missing settings get new first-run defaults')
verify('FitWindowToWorkArea(System.Windows.Forms.Screen.FromControl(Me).WorkingArea)' in picker and 'System.Math.Min(Width, workArea.Width)' in picker, 'Wider default fits current display on first opening')
verify('menu.Items.Add("Lizenzhinweise …"' in context and 'menu.Items.Add(New System.Windows.Forms.ToolStripSeparator())' in context and 'New LicenseInfoForm()' in context, 'Discrete offline license viewer accessible via tray')
verify('ShowLicenseInformation()' in context and 'If exiting Then Return' in context and 'Using license As New LicenseInfoForm()' in context, 'License viewer lifecycle safely scoped')
legal=(root/'LegalResources.vb').read_text()
licenseform=(root/'LicenseInfoForm.vb').read_text()
licensebytes=(root/'LICENSE.txt').read_bytes()
verify('GetManifestResourceStream("PromptPicker.LICENSE.txt")' in legal and 'reader.ReadToEnd()' in legal, 'License text loaded from embedded resource, not downloaded')
verify('LegalResources.LicenseText' in licenseform and 'ReadOnly = True' in licenseform and 'ScrollBars.Vertical' in licenseform, 'License text visible and read-only with scrolling')
verify('PickerUi.ActionButton("Schliessen", "close"' in licenseform and 'TopMost = True' in licenseform and 'SizeType.Percent, 100.0F' in licenseform, 'License dialog matches compact native UI and resizes')
verify(b'RED INK PROMPT PICKER' in licensebytes and b'RED INK \r\n=======' in licensebytes and b'ASK INKY \xe2\x80\x93 RED INK COMPONENT NOTICE' not in licensebytes, 'Notices identify Prompt Picker rather than its Ask Inky source wrapper')
verify('EmbeddedResource Include="LICENSE.txt" LogicalName="PromptPicker.LICENSE.txt"' in (root/'PromptPicker.vbproj').read_text(), 'Legal resource remains bundled for single-EXE publication')

verify('ClientSize = New System.Drawing.Size(System.Math.Clamp(pickerSettings.PickerClientWidth' in picker, 'Picker restores saved client size')
verify('Text = "Red Ink Prompt Picker"' in picker and '<title>Red Ink Prompt Picker</title>' in html, 'Dialog/window title')
verify('RestoreWindowPosition()' in picker and 'Screen.AllScreens' in picker and 'visibleScreen.WorkingArea' in picker, 'Window bounds restored only on a visible monitor')
verify('Protected Overrides Sub OnLoad' in picker and picker.index('RestoreWindowPosition()', picker.index('Protected Overrides Sub OnLoad')) > picker.index('ClientSize = New System.Drawing.Size(', picker.index('Protected Overrides Sub OnLoad')), 'DPI scaling finished before persisted size and position reapplied')
verify('PickerX As Integer?' in settings and 'PickerY As Integer?' in settings and 'CaptureWindowSettings()' in picker, 'Backward-compatible persisted coordinates and geometry capture')
verify('SettingsStore.Save(currentSettings)' in context and 'desired.PreviewPercent = currentSettings.PreviewPercent' in context, 'Window and splitter geometry persists across hotkey edits')
verify('PreviewPercent As Integer = 55' in settings and 'value >= 32 AndAlso value <= 68' in picker, 'Preview default 55 percent; untrusted split messages validated')
verify('System.Windows.Forms.SplitContainer' in picker and 'nativeSplitter.SplitterDistance' in picker, 'Native fallback has adjustable split width')

verify('TopMost = True' in picker and 'TopMost = True' in save and 'TopMost = True' in shortcut, 'All picker dialogs TopMost')
verify('<header' not in html and 'class="intro"' not in html and 'logo.png' not in html, 'No logo/header/intro in browser picker')
verify('placeholder="Prompts durchsuchen"' in html, 'Search begins picker page')
verify('id="splitter"' in html and 'role="separator"' in html and 'gridTemplateColumns' in js, 'Browser preview separator is mouse and keyboard adjustable')
verify('resizePanels(model.previewPercent)' in js and 'post({type:"split",previewPercent})' in js, 'Preview width round-trips to host settings')
verify('window.addEventListener("focus"' in js and 'view.CoreWebView2.ExecuteScriptAsync' in picker, 'Search focus restored on browser or host activation')
verify('e.key === "ArrowDown"' in js and 'document.activeElement !== category' in js and 'e.preventDefault(); commit();' in js, 'Up/down navigation from search and Enter commit')

verify('results.addEventListener("pointerup"' in js and 'results.addEventListener("dblclick"' in js and 'if (!p || committing) return;' in js, 'WebView2 pointer/double-click converges on one commit')
verify('BeginInvoke(New System.Action(Sub() CommitSelection(id)))' in picker and 'picker.selection.commit' in picker, 'Deferred host commit avoids disposing WebView2 mid-click')
verify('CaptureFocusedControl(target)' in context and 'RestoreFocusedControl(window, originalFocus)' in services, 'External editor focus captured and restored generically')
verify('insertion.dispatch;' in services and 'insertion.activate.attempt=' in services and 'insertion.keys.dispatched' in services, 'Insertion stages logged for Windows runtime diagnosis')

verify('WaitUntilReleasedAsync(True)' in services and 'SendPaste()' in services and 'Native.SendInput(' in services, 'Mouse-up guarded Win32 paste injection')
verify('Public Mouse As MouseRecord' in native and 'Public Keyboard As KeyboardRecord' in native and 'Public Function SendInput' in native, 'INPUT union correct for mouse-size x64 ABI')
verify('System.Windows.Forms.SendKeys.Send("^v")' not in services, 'No focus-sensitive SendKeys paste')
verify('System.Windows.Forms.TableLayoutPanel' in save and 'System.Windows.Forms.SizeType.AutoSize' in save, 'DPI-resilient save dialog rows')
verify('PromptLibLocal' not in save and 'redink.ini' not in save, 'Save dialog contains no internal settings names')
verify('KeyOption' in shortcut and 'TopMost = True' in shortcut, 'Hotkey dialog: readable key names and same window behavior')
verify('ClientSize = New System.Drawing.Size(515, 218)' in shortcut and '.RowCount = 5' in shortcut and '.MaximumSize = New System.Drawing.Size(465, 0)' in shortcut, 'Hotkey dialog expanded with unconstrained hint row')

verify('PickerUi.ActionButton' in save and 'PickerUi.ActionButton' in shortcut and 'PickerUi.ActionButton' in picker, 'Consistent vector-icon action buttons')
verify('System.Drawing.Bitmap' in ui and 'System.Drawing.Graphics' in ui, 'Native line icons require no icon font')
verify('WriteAllText' not in picker and 'WriteAllText' not in save and 'WriteAllText' not in shortcut, 'UI never writes INI or library files')

# Pure-format spec fixtures (independent Python implementation). These do not
# execute VB code; static invariants above tie it to the intended implementation.
def read_lines(content, suffix):
    result=[]; category=None
    for line in content.splitlines():
        line=line.strip()
        if not line or line.startswith(';'): continue
        if line.startswith('</') and line.endswith('>'): category=None; continue
        if line.startswith('<') and line.endswith('>') and '|' not in line:
            cat=line[1:-1].strip()
            if cat and not cat.startswith('/'):
                category=cat; continue
        if '|' in line:
            title,prompt=line.split('|',1)
            result.append((title.strip()+suffix,prompt.strip(),category or '(Uncategorized)'))
    return result
local = '<Lokal>\nA|Mehr|Pipe\n</Lokal>\nX|eins\\nzwei'
central = '; Kommentar\n<Zentral>\nB|Weitere Auswahl\n</Zentral>'
items=read_lines(local,' (local)')+read_lines(central,'')
verify(items==[('A (local)','Mehr|Pipe','Lokal'),('X (local)','eins\\nzwei','(Uncategorized)'),('B','Weitere Auswahl','Zentral')], 'Prompt library format parity fixture')
print('All static checks passed. Windows build/runtime NOT tested.')
