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
