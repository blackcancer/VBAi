import * as monaco from '../../../artifacts/monaco-dependencies/monaco/package/esm/vs/editor/editor.api.js';
import '../../../artifacts/monaco-dependencies/monaco/package/esm/vs/editor/editor.all.js';
import { installLanguage } from './language.js';

self.MonacoEnvironment = { getWorker: () => new Worker('./editor.worker.js', { type: 'module' }) };
const keywords = 'Option Explicit Private Public Friend Static Dim Const Sub Function Property Get Let Set End If Then Else ElseIf Select Case For Each Next Do Loop While Wend With New Nothing As ByVal ByRef Optional ParamArray ReDim Preserve Return Exit On Error Resume GoTo GoSub Declare PtrSafe Lib Alias Type Enum Implements WithEvents Call Me And Or Not Xor Mod Is Like True False Null Empty Boolean Byte Integer Long LongLong LongPtr Single Double Currency Date String Object Variant Debug Stop'.split(' ');
monaco.languages.register({ id: 'vba' });
monaco.languages.setMonarchTokensProvider('vba', {
  ignoreCase: true, keywords: keywords.map(k => k.toLowerCase()),
  tokenizer: { root: [
    [/'[^\n]*/, 'comment'], [/\bRem\s.*$/, 'comment'], [/"([^"\n]|"")*"?/, 'string'],
    [/^\s*#(?:If|ElseIf|Else|End\s+If|Const)\b/, 'keyword'], [/#.*?#/, 'number'],
    [/&[Hh][0-9a-fA-F]+|\b\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?/, 'number'],
    [/[a-zA-Z_\u0080-\uFFFF][\w\u0080-\uFFFF]*/, { cases: { '@keywords': 'keyword', '@default': 'identifier' } }],
    [/[()[\]]/, '@brackets'], [/[+\-*/=<>:&]/, 'operator']
  ] }
});
monaco.languages.setLanguageConfiguration('vba', {
  comments: { lineComment: "'" }, brackets: [['(', ')']],
  autoClosingPairs: [{ open: '(', close: ')' }, { open: '"', close: '"', notIn: ['string', 'comment'] }],
  surroundingPairs: [{ open: '(', close: ')' }, { open: '"', close: '"' }]
});
monaco.languages.registerCompletionItemProvider('vba', {
  provideCompletionItems(model, position) {
    const word = model.getWordUntilPosition(position);
    return { suggestions: keywords.map(label => ({ label, kind: monaco.languages.CompletionItemKind.Keyword, insertText: label,
      range: { startLineNumber: position.lineNumber, endLineNumber: position.lineNumber, startColumn: word.startColumn, endColumn: word.endColumn } })) };
  }
});
monaco.editor.defineTheme('vbai-dark', { base: 'vs-dark', inherit: true, rules: [], colors: {
  'editor.background': '#171b21', 'editor.foreground': '#e5e7eb', 'editorLineNumber.foreground': '#8997aa',
  'editor.selectionBackground': '#264f78', 'editorWidget.background': '#20252d', 'editorGutter.background': '#171b21'
} });
monaco.editor.defineTheme('vbai-light', { base: 'vs', inherit: true, rules: [], colors: { 'editor.background': '#ffffff' } });
const models = new Map();
let active = null, suppress = false, diff = null;
const editor = monaco.editor.create(document.getElementById('editor'), {
  theme: 'vbai-dark', language: 'vba', automaticLayout: true, fontFamily: 'Cascadia Code, Consolas, monospace', fontSize: 14,
  minimap: { enabled: false }, scrollBeyondLastLine: false, renderWhitespace: 'selection', wordBasedSuggestions: 'currentDocument',
  tabSize: 4, insertSpaces: true, readOnly: false, glyphMargin: true
});
function send(value) { window.chrome.webview.postMessage(value); }
const language = installLanguage(monaco, editor, models, send);
let commandLabels = {}, nativeActions = [];
function snapshot(entry) { return { id: entry.id, text: entry.model.getValue(), version: entry.model.getVersionId() }; }
function select(id) {
  if (active && models.has(active)) models.get(active).view = editor.saveViewState();
  active = id; const entry = models.get(id); editor.setModel(entry.model);
  if (entry.view) editor.restoreViewState(entry.view);
  hideDiff(); editor.focus();
}
function hideDiff() {
  if (diff) { const old = diff.getModel(); diff.dispose(); old.original.dispose(); diff = null; }
  document.getElementById('diff').hidden = true; document.getElementById('editor').hidden = false;
}
window.vbai = {
  languageReply: language.reply,
  labels(value) { commandLabels = value; installNativeActions(); installAssistantActions(); },
  languageInspect(id, line, column) { return language.inspect(models.get(id).model, { lineNumber: line, column }); },
  diagnostics(id, version, markers) { const entry = models.get(id); if (entry && entry.model.getVersionId() === version) monaco.editor.setModelMarkers(entry.model, 'VBA compiler', markers); },
  execution(id, line, reveal = true) { const entry = models.get(id); if (!entry) return; entry.execution = entry.model.deltaDecorations(entry.execution || [], line > 0 ? [{ range: new monaco.Range(line, 1, line, 1), options: { isWholeLine: true, className: 'vbai-execution-line', glyphMarginClassName: 'vbai-execution', glyphMarginHoverMessage: { value: 'VBE: Show Next Statement' } } }] : []); if (line > 0 && reveal && active === id) editor.revealLineInCenter(line); },
  breakpointRequested(id, line) { const entry = models.get(id); if (!entry) return; entry.breakpoints ||= new Map(); if (entry.breakpoints.has(line)) { entry.model.deltaDecorations(entry.breakpoints.get(line), []); entry.breakpoints.delete(line); return; } entry.breakpoints.set(line, entry.model.deltaDecorations([], [{ range: new monaco.Range(line, 1, line, 1), options: { glyphMarginClassName: 'vbai-breakpoint-pending', glyphMarginHoverMessage: { value: commandLabels['Breakpoint request sent; verify in VBE.'] || 'Breakpoint request sent; verify in VBE.' } } }])); },
  open(id, text) {
    if (!models.has(id)) {
      const model = monaco.editor.createModel(text, 'vba', monaco.Uri.parse('vbai://module/' + id));
      const entry = { id, model, view: null }; models.set(id, entry);
      model.onDidChangeContent(() => {
        monaco.editor.setModelMarkers(model, 'VBA compiler', []);
        entry.execution = model.deltaDecorations(entry.execution || [], []);
        for (const decorations of entry.breakpoints?.values() || []) model.deltaDecorations(decorations, []);
        entry.breakpoints?.clear();
        if (!suppress) send({ type: 'change', ...snapshot(entry) });
      });
    }
    select(id); return models.get(id).model.getVersionId();
  },
  close(id) { const entry = models.get(id); if (!entry) return; if (active === id) { hideDiff(); editor.setModel(null); active = null; } language.close(id); entry.model.dispose(); models.delete(id); },
  select, snapshots: () => [...models.values()].map(snapshot),
  apply(id, version, text) {
    const entry = models.get(id); if (!entry || entry.model.getVersionId() !== version) return 0;
    if (entry.model.getValue() !== text) {
      suppress = true;
      try { entry.model.pushStackElement(); entry.model.pushEditOperations([], [{ range: entry.model.getFullModelRange(), text }], () => null); entry.model.pushStackElement(); }
      finally { suppress = false; }
    }
    return entry.model.getVersionId();
  },
  theme(dark, contrast) { monaco.editor.setTheme(contrast ? (dark ? 'hc-black' : 'hc-light') : dark ? 'vbai-dark' : 'vbai-light'); document.body.style.background = dark ? '#171b21' : '#ffffff'; },
  compare(original) {
    hideDiff(); if (!active) return;
    document.getElementById('editor').hidden = true; document.getElementById('diff').hidden = false;
    diff = monaco.editor.createDiffEditor(document.getElementById('diff'), { automaticLayout: true, renderSideBySide: true, readOnly: true, originalEditable: false });
    diff.setModel({ original: monaco.editor.createModel(original, 'vba'), modified: models.get(active).model });
  },
  hideDiff,
  selection() { const range = editor.getSelection(); return range ? { startLine: range.startLineNumber, startColumn: range.startColumn, endLine: range.endLineNumber, endColumn: range.endColumn, text: editor.getModel().getValueInRange(range) } : null; },
  read(id) { const entry = models.get(id); if (!entry) return null; const range = active === id ? editor.getSelection() : null; return { ...snapshot(entry), selection: range ? { startLine: range.startLineNumber, startColumn: range.startColumn, endLine: range.endLineNumber, endColumn: range.endColumn, text: entry.model.getValueInRange(range) } : null }; },
  selectRange(id, version, startLine, startColumn, endLine, endColumn) { const entry = models.get(id); if (!entry || entry.model.getVersionId() !== version) return false; if (active !== id) select(id); hideDiff(); const range = new monaco.Range(startLine, startColumn, endLine, endColumn); editor.setSelection(range); editor.revealRangeInCenter(range); editor.focus(); return true; },
  position() { const p = editor.getPosition(); return p ? { line: p.lineNumber, column: p.column } : { line: 1, column: 1 }; },
  reveal(line, column) { editor.setPosition({ lineNumber: line, column }); editor.revealLineInCenter(line); editor.focus(); },
  command(name) { const action = editor.getAction(name); if (action) action.run(); else editor.trigger('vbai', name, {}); },
  insert(text) { editor.trigger('vbai', 'type', { text }); },
  testInfo() { return { language: editor.getModel()?.getLanguageId(), models: models.size, theme: document.body.style.background, version: monaco.editor?.getModels().length, diff: !!diff, pendingBreakpoints: models.get(active)?.breakpoints?.size || 0, executionMarkers: models.get(active)?.execution?.length || 0, markers: editor.getModel() ? monaco.editor.getModelMarkers({ resource: editor.getModel().uri }).length : 0 }; }
};
editor.addAction({ id: 'vbai.save', label: 'VBA: Save', keybindings: [monaco.KeyMod.CtrlCmd | monaco.KeyCode.KeyS],
  run: () => { if (active) send({ type: 'command', name: 'save', id: active }); } });
function nativeCommand(name) { if (!active) return; const entry = models.get(active); send({ type: 'editorCommand', name, id: active, version: entry.model.getVersionId(), line: editor.getPosition()?.lineNumber || 1 }); }
function installNativeActions() {
  for (const action of nativeActions) action.dispose(); nativeActions = [];
  for (const [name, label, binding] of [
    ['compile', 'Compile project', monaco.KeyMod.CtrlCmd | monaco.KeyMod.Shift | monaco.KeyCode.KeyB],
    ['toggle_breakpoint', 'Toggle breakpoint', monaco.KeyCode.F9],
    ['show_next_statement', 'Show next statement', 0],
    ['step_into', 'Step into', monaco.KeyCode.F8],
    ['step_over', 'Step over', monaco.KeyMod.Shift | monaco.KeyCode.F8],
    ['step_out', 'Step out', monaco.KeyMod.CtrlCmd | monaco.KeyMod.Shift | monaco.KeyCode.F8]
  ]) nativeActions.push(editor.addAction({ id: 'vbai.' + name, label: 'VBA: ' + (commandLabels[label] || label), keybindings: binding ? [binding] : [], contextMenuGroupId: 'vba', run: () => nativeCommand(name) }));
}
installNativeActions();
let assistantActions = [];
function installAssistantActions() {
for (const action of assistantActions) action.dispose(); assistantActions = [];
for (const [name, label] of [['expliquer', 'Explain'], ['corriger', 'Fix'], ['refactoriser', 'Refactor']]) {
  assistantActions.push(editor.addAction({ id: 'vbai.' + name, label: 'VBAi: ' + (commandLabels[label] || label), contextMenuGroupId: 'vbai',
    run: () => { if (!active) return; const entry = models.get(active), selection = editor.getSelection();
      const text = selection && !selection.isEmpty() ? entry.model.getValueInRange(selection) : entry.model.getValue();
      send({ type: 'assistantAction', name: '/' + name, ...snapshot(entry), selectedText: text, line: selection && !selection.isEmpty() ? selection.startLineNumber : 1 }); }
  }));
}
}
installAssistantActions();
editor.onMouseDown(e => { if ([monaco.editor.MouseTargetType.GUTTER_GLYPH_MARGIN, monaco.editor.MouseTargetType.GUTTER_LINE_NUMBERS].includes(e.target.type) && e.target.position) { editor.setPosition(e.target.position); nativeCommand('toggle_breakpoint'); } });
const style = document.createElement('style'); style.textContent = '.vbai-breakpoint-pending::before { content:"○";color:#ef5350;font-size:20px;font-weight:bold; } .vbai-execution::before { content:"➜";color:#f7c948; } .vbai-execution-line { background:#f7c94826; }'; document.head.appendChild(style);
send({ type: 'ready' });
