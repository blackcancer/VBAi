// Run with: node tools/tests/Test-MonacoLanguage.mjs
// Exercises the actual language provider with immutable symbol fixtures; no Office or npm required.
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { test } from 'node:test';

const source = await readFile(new URL('../../assets/editor/src/language.js', import.meta.url), 'utf8');
const { installLanguage } = await import('data:text/javascript;base64,' + Buffer.from(source).toString('base64'));
const symbol = (Name, Module, TypeName, extras = {}) => ({ Name, Module, TypeName, Kind: 'Property', Scope: 'Module',
  Line: 1, Column: 1, EndLine: 100, Parameters: [], Declaration: Name, ...extras });
const external = (name, owner, type, library) => symbol(name, owner, type, { External: true, Library: library });
const symbols = [
  symbol('Test', 'Main', 'Variant', { Kind: 'Procedure' }),
  symbol('sheet', 'Main', 'Excel.Worksheet', { Kind: 'Variable', Scope: 'Test' }),
  symbol('app', 'Main', 'Word.Application', { Kind: 'Variable' }),
  symbol('app', 'Main', 'Excel.Application', { Kind: 'Variable', Scope: 'Test' }),
  symbol('wordApp', 'Main', 'Word.Application', { Kind: 'Variable', Scope: 'Test' }),
  symbol('hiddenOther', 'Other', 'String', { Private: true }),
  external('Range', 'Worksheet', 'Excel.Range', 'Excel'),
  external('Cells', 'Worksheet', 'Excel.Range', 'Excel'),
  external('Font', 'Range', 'Excel.Font', 'Excel'),
  external('Value', 'Range', 'Variant', 'Excel'),
  external('Bold', 'Font', 'Boolean', 'Excel'),
  external('Workbooks', 'Application', 'Excel.Workbooks', 'Excel'),
  external('Documents', 'Application', 'Word.Documents', 'Word'),
  external('Item', 'Workbooks', 'Excel.Workbook', 'Excel'),
  external('Worksheets', 'Workbook', 'Excel.Worksheets', 'Excel'),
  external('Count', 'Documents', 'Long', 'Word'),
];

async function complete(text, fixture = symbols) {
  const lines = text.split('\n');
  const position = { lineNumber: lines.length, column: lines.at(-1).length + 1 };
  const model = { getVersionId: () => 1, isDisposed: () => false,
    getLineContent: line => lines[line - 1],
    getWordUntilPosition: () => ({ startColumn: position.column, endColumn: position.column }) };
  let provider, language;
  const monaco = {
    Range: class { constructor(...values) { this.values = values; } },
    languages: { registerCompletionItemProvider: (_, value) => { provider = value; },
      registerHoverProvider() {}, registerDefinitionProvider() {}, registerSignatureHelpProvider() {} },
    editor: { registerEditorOpener() {} },
  };
  language = installLanguage(monaco, {}, new Map([['doc', { id: 'doc', model }]]), message => {
    queueMicrotask(() => language.reply(message.request, { id: 'doc', module: 'Main', symbols: fixture, sources: [] }));
  });
  const result = await provider.provideCompletionItems(model, position);
  return result.suggestions.map(item => item.label).sort();
}

test('typed receiver lists only its members', async () => {
  assert.deepEqual(await complete('sheet.'), ['Cells', 'Range']);
});
test('property and call chains resolve return types', async () => {
  assert.deepEqual(await complete('sheet.Range("A1").Font.'), ['Bold']);
});
test('nested arguments and punctuation in strings do not alter receiver', async () => {
  assert.deepEqual(await complete('sheet.Range(Choose(1, "A1", "B2")).Font.'), ['Bold']);
});
test('nested With uses its parent receiver', async () => {
  assert.deepEqual(await complete('With sheet\n  With .Range("A1")\n    .Font.'), ['Bold']);
});
test('End With restores the outer receiver', async () => {
  assert.deepEqual(await complete('With sheet\n  With .Range("A1")\n  End With\n  .'), ['Cells', 'Range']);
});
test('qualified Excel Application excludes Word Application members', async () => {
  assert.deepEqual(await complete('app.'), ['Workbooks']);
});
test('qualified Word Application excludes Excel Application members', async () => {
  assert.deepEqual(await complete('wordApp.'), ['Documents']);
});
test('qualification survives chained COM return types', async () => {
  assert.deepEqual(await complete('app.Workbooks.Item(1).'), ['Worksheets']);
});
test('locals shadow module declarations and other private declarations stay hidden', async () => {
  const names = await complete('');
  assert.equal(names.filter(name => name === 'app').length, 1);
  assert.equal(names.includes('hiddenOther'), false);
  assert.deepEqual(await complete('app.'), ['Workbooks']);
});
test('apostrophes inside With string arguments are not VBA comments', async () => {
  assert.deepEqual(await complete('With sheet.Range("A\'1")\n  .Font.'), ['Bold']);
});
test('With receiver also resolves on assignment right hand side', async () => {
  assert.deepEqual(await complete('With sheet\n  result = .'), ['Cells', 'Range']);
});

test('Debug exposes VBA Print and Assert without irrelevant keywords', async () => {
  assert.deepEqual(await complete('Debug.'), ['Assert', 'Print']);
});
test('library namespaces expose their referenced types', async () => {
  const library = [symbol('Scripting', 'Scripting', 'Scripting', { Kind: 'Library', External: true, Library: 'Scripting' }),
    symbol('Dictionary', 'Scripting', 'Scripting.Dictionary', { Kind: 'Class', External: true, Library: 'Scripting' }),
    external('Add', 'Dictionary', 'Void', 'Scripting')];
  assert.deepEqual(await complete('Scripting.', library), ['Dictionary']);
  assert.deepEqual(await complete('Scripting.Dictionary.', library), ['Add']);
  assert.equal((await complete('', library)).includes('Dictionary'), true);
});
test('library global functions and enum values are available unqualified', async () => {
  const fixture = [external('MsgBox', 'Interaction', 'Long', 'VBA'), external('vbTextCompare', 'VbCompareMethod', 'Long', 'VBA')].map(s => ({ ...s, Global: true }));
  const names = await complete('', fixture); assert.equal(names.includes('MsgBox'), true); assert.equal(names.includes('vbTextCompare'), true);
});
test('default collection indexing resolves its item and property collections', async () => {
  assert.deepEqual(await complete('app.Workbooks(1).', symbols), ['Worksheets']);
  const fixture = [...symbols, symbol('books', 'Main', 'Excel.Workbooks', { Kind: 'Variable', Scope: 'Test' })];
  assert.deepEqual(await complete('books(1).', fixture), ['Worksheets']);
});
test('late-bound variables assigned New resolve metadata without executing objects', async () => {
  const fixture = [symbol('Test', 'Main', '', { Kind: 'Procedure' }), symbol('d', 'Main', 'Object', { Kind: 'Variable', Scope: 'Test' }), external('Add', 'Dictionary', 'Void', 'Scripting')];
  assert.deepEqual(await complete('Set d = New Scripting.Dictionary\nd.', fixture), ['Add']);
  assert.deepEqual(await complete('Set d = New Scripting.Dictionary\nSet d = Nothing\nd.', fixture), []);
});
test('unqualified referenced types follow the first library while explicit qualification remains exact', async () => {
  const fixture = [symbol('Test', 'Main', '', { Kind: 'Procedure' }), symbol('a', 'Main', 'Application', { Kind: 'Variable', Scope: 'Test' }),
    symbol('Application', 'Excel', 'Excel.Application', { Kind: 'Class', External: true, Library: 'Excel' }),
    symbol('Application', 'Word', 'Word.Application', { Kind: 'Class', External: true, Library: 'Word' }), ...symbols];
  assert.deepEqual(await complete('a.', fixture), ['Workbooks']);
});
test('property getters take precedence over setters during member chaining', async () => {
  const fixture = [...symbols]; fixture.unshift(external('Font', 'Range', 'Void', 'Excel'));
  assert.deepEqual(await complete('sheet.Range("A1").Font.', fixture), ['Bold']);
});
test('completion is suppressed in strings, apostrophe comments and Rem comments', async () => {
  for (const text of ['\' Debug.', 'Rem Debug.', 'If ready Then Rem Debug.', 'Debug.Print "Debug.', 'text = "an escaped ""Debug.'])
    assert.deepEqual(await complete(text), []);
});
test('conditional members and other-class private members stay hidden on typed receivers', async () => {
  const fixture = [...symbols, external('PrivateMethod', 'Worksheet', '', 'Excel'), external('ConditionalMethod', 'Worksheet', '', 'Excel')];
  fixture.at(-2).Private = true; fixture.at(-1).Conditional = true;
  assert.deepEqual(await complete('sheet.', fixture), ['Cells', 'Range']);
  assert.deepEqual(await complete('sheet.Missing.'), []);
});

function service(text, fixture = symbols) {
  const lines = text.split('\n'); let provider, hover, signature, language, requests = 0, current = fixture, disposed = false, version = 1;
  const model = { getVersionId: () => version, isDisposed: () => disposed, getLineContent: line => lines[line - 1],
    getWordUntilPosition: position => ({ startColumn: position.column, endColumn: position.column }),
    getWordAtPosition: position => {
      for (const match of lines[position.lineNumber - 1].matchAll(/[\p{L}_][\p{L}\p{N}_]*\$?/gu))
        if (position.column >= match.index + 1 && position.column <= match.index + match[0].length + 1)
          return { word: match[0], startColumn: match.index + 1, endColumn: match.index + match[0].length + 1 };
      return null;
    } };
  const monaco = { Range: class { constructor(...values) { this.values = values; } },
    languages: { registerCompletionItemProvider: (_, value) => provider = value, registerHoverProvider: (_, value) => hover = value,
      registerDefinitionProvider() {}, registerSignatureHelpProvider: (_, value) => signature = value }, editor: { registerEditorOpener() {} } };
  language = installLanguage(monaco, {}, new Map([['doc', { id: 'doc', model }]]), message => {
    requests++; queueMicrotask(() => language.reply(message.request, { id: 'doc', module: 'Main', symbols: current, sources: [] }));
  });
  return { model, provider, hover, signature, language, position: { lineNumber: lines.length, column: lines.at(-1).length + 1 },
    get requests() { return requests; }, set fixture(value) { current = value; }, set disposed(value) { disposed = value; }, set version(value) { version = value; } };
}
test('reference additions and removals refresh without changing the model revision', async () => {
  const state = service('sheet.');
  assert.equal((await state.provider.provideCompletionItems(state.model, state.position)).suggestions.length, 2);
  state.fixture = [...symbols, external('Name', 'Worksheet', 'String', 'Excel')];
  assert.equal((await state.provider.provideCompletionItems(state.model, state.position)).suggestions.length, 3);
  state.fixture = []; assert.equal((await state.provider.provideCompletionItems(state.model, state.position)).suggestions.length, 0);
  assert.equal(state.requests, 3);
});
test('simultaneous completion and hover coalesce only their in-flight request', async () => {
  const state = service('sheet.Range');
  const results = await Promise.all([state.provider.provideCompletionItems(state.model, state.position), state.hover.provideHover(state.model, state.position)]);
  assert.equal(state.requests, 1); assert.ok(results[1]);
  await state.hover.provideHover(state.model, state.position); assert.equal(state.requests, 2);
});
test('hover supplies declaration, native documentation, library identity and safe Markdown', async () => {
  const fixture = [external('Run', 'Worksheet', 'Long', 'Excel')];
  Object.assign(fixture[0], { Declaration: 'Worksheet.Run(ByVal value As Long) As Long', Documentation: 'Run *this* <script>', LibraryDescription: 'Excel object library', LibraryPath: 'C:\\Libraries\\Excel.exe', HelpFile: 'Excel.chm', HelpContext: 123 });
  const state = service('sheet.Run', [...symbols.filter(s => s.Name !== 'Range'), ...fixture]);
  const result = await state.hover.provideHover(state.model, state.position);
  assert.ok(result.contents.some(part => part.value.includes('ByVal value As Long')));
  assert.ok(result.contents.some(part => part.value.includes('Library: **Excel**')));
  assert.ok(result.contents.some(part => part.value.includes('Excel object library')));
  assert.ok(result.contents.some(part => part.value.includes('Excel.chm (123)')));
  assert.ok(result.contents.every(part => part.isTrusted === false));
  assert.ok(result.contents.some(part => part.value.includes('\\<script\\>')));
  assert.deepEqual(result.range.values, [1, 7, 1, 10]);
});
test('Debug hover exposes the actual operation description', async () => {
  const state = service('Debug.Print'); const result = await state.hover.provideHover(state.model, state.position);
  assert.ok(result.contents.some(part => part.value.includes('Immediate window')));
});
test('signature help retains typed parameter descriptions and handles nested arguments', async () => {
  const member = external('Run', 'Worksheet', 'Long', 'Excel'); member.Parameters = ['ByVal value As Long', 'Optional ByVal text As String'];
  member.Declaration = 'Worksheet.Run(ByVal value As Long, Optional ByVal text As String) As Long'; member.Documentation = 'Native operation';
  const state = service('sheet.Run(Choose(1, "a,b", "c"), ', [...symbols, member]);
  const result = await state.signature.provideSignatureHelp(state.model, state.position);
  assert.equal(result.value.activeParameter, 1); assert.equal(result.value.signatures[0].documentation, 'Native operation');
  assert.equal(result.value.signatures[0].label, member.Declaration);
});
test('stale and disposed models do not publish obsolete language results', async () => {
  for (const mode of ['version', 'disposed']) {
    const state = service('sheet.'); const request = state.provider.provideCompletionItems(state.model, state.position);
    state[mode] = mode === 'version' ? 2 : true;
    assert.equal((await request).suggestions.length, 0);
  }
});

test('native VBA aliases expose usable Variant and String names without rewriting other libraries', async () => {
  const fixture = ['str', 'var'].map(kind => ({ ...external('_B_' + kind + '_Left', 'Strings', kind === 'str' ? 'String' : 'Variant', 'VBA'), Global: true, Declaration: 'Strings._B_' + kind + '_Left(value, length)' }));
  fixture.push(external('_B_var_Custom', 'Strings', 'Variant', 'Other'));
  for (const library of ['VBA', 'Other']) fixture.push({ ...external('Strings', library, library + '.Strings', library), Kind: 'Type' });
  const names = await complete('VBA.Strings.', fixture);
  assert.deepEqual(names, ['Left', 'Left$']);
  assert.ok((await complete('', fixture)).includes('_B_var_Custom') === false);
  const state = service('VBA.Strings.Left$', fixture);
  const hover = await state.hover.provideHover(state.model, state.position);
  assert.ok(hover.contents.some(part => part.value.includes('Strings.Left$(value, length)')));
  assert.ok(hover.contents.every(part => !part.value.includes('_B_str_')));
  const signature = service('VBA.Strings.Left$("text", ', fixture.map(s => ({ ...s, Parameters: ['value', 'length'] })));
  assert.equal((await signature.signature.provideSignatureHelp(signature.model, signature.position)).value.signatures[0].label, 'Strings.Left$(value, length)');
  assert.deepEqual(await complete('Other.Strings.', fixture), ['_B_var_Custom']);
});
