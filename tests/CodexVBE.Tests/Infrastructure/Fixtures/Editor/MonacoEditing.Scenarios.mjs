import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { test } from 'node:test';
const root = new URL('../../../../../', import.meta.url);
const source = await readFile(new URL('assets/editor/src/editing.js', root), 'utf8');
const { codeMask, formatLines, enterPlan, installEditing } = await import('data:text/javascript;base64,' + Buffer.from(source).toString('base64'));

for (const [header, end] of [
  ['Sub Example()', 'End Sub'], ['Private Static Sub Example()', 'End Sub'], ['Public Function Compute() As Long', 'End Function'],
  ['Property Get Caption() As String', 'End Property'], ['Private Property Let Caption(ByVal value As String)', 'End Property'],
  ['Property Set Item(ByVal value As Object)', 'End Property'], ['If ready Then', 'End If'],
  ['For i = 1 To 10', 'Next i'], ['For Each item In collection', 'Next item'], ['For', 'Next'],
  ['Do', 'Loop'], ['Do While ready', 'Loop'], ['While ready', 'Wend'], ['With application', 'End With'],
  ['Select Case value', 'End Select'], ['Public Type Record', 'End Type'], ['Enum Mode', 'End Enum'], ['#If VBA7 Then', '#End If']
]) test('Enter closes ' + header, () => {
  const plan = enterPlan([header, ''], 2);
  assert.equal(plan.text, '    \n' + end); assert.equal(plan.column, 5);
  const existing = enterPlan([header, '', end], 2);
  assert.equal(existing.text, '    ');
});

test('completed multiline procedure gets an aligned closing block', () => {
  const plan = enterPlan(['Private Sub Work( _', 'ByVal x As Long)', ''], 3);
  assert.equal(plan.header, '    ByVal x As Long)'); assert.equal(plan.text, '    \nEnd Sub');
});
test('continuations, inline If and strings do not introduce closures', () => {
  for (const text of ['If ready Then Debug.Print ready', "' Sub Example()", 'Rem For i = 1 To 3', 'Debug.Print "Sub Example()"', 'Sub Work( _'])
    assert.equal(enterPlan([text, ''], 2).text.includes('End Sub'), false);
  assert.equal(enterPlan(['If ready Then Debug.Print ready', ''], 2).text, '');
  assert.equal(enterPlan(['Sub Work( _', ''], 2).text, '    ');
});
test('existing nested blocks are not mistaken for an outer closure', () => {
  assert.equal(enterPlan(['If a Then', '', 'If b Then', 'End If'], 2).text, '    \nEnd If');
  assert.equal(enterPlan(['If a Then', '', 'If b Then', 'End If', 'End If'], 2).text, '    ');
  assert.equal(enterPlan(['Sub First()', '', 'Sub Second()', 'End Sub'], 2).text, '    \nEnd Sub');
});
test('empty first line and splitting a populated line are left alone', () => {
  assert.equal(enterPlan([''], 1), null); assert.equal(enterPlan(['Sub Work()', 'remaining code'], 2), null);
});
test('indentation covers procedures, branches, loops, select cases and labels', () => {
  const input = ['Sub Work()', 'If a Then', 'Debug.Print " a "', 'ElseIf b Then', 'For i = 1 To 3', 'Debug.Print i', 'Next i', 'Else', 'Select Case value', 'Case 1', 'Debug.Print "one"', 'Case Else', 'Debug.Print "other"', 'End Select', 'End If', 'exitPoint:', 'End Sub'];
  assert.deepEqual(formatLines(input).lines, ['Sub Work()', '    If (a) Then', '        Debug.Print " a "', '    ElseIf (b) Then', '        For i = 1 To 3', '            Debug.Print i', '        Next i', '    Else', '        Select Case value', '            Case 1', '                Debug.Print "one"', '            Case Else', '                Debug.Print "other"', '        End Select', '    End If', 'exitPoint:', 'End Sub']);
});
test('multiple Next variables close the matching nested loops', () => {
  assert.deepEqual(formatLines(['For i = 1 To 3', 'For j = 1 To 3', 'Next j, i', 'Debug.Print 1']).lines,
    ['For i = 1 To 3', '    For j = 1 To 3', 'Next j, i', 'Debug.Print 1']);
});
test('tabs and nondefault space widths follow Monaco options', () => {
  assert.equal(enterPlan(['Sub Work()', ''], 2, { insertSpaces: false, tabSize: 2 }).text, '\t\nEnd Sub');
  assert.equal(enterPlan(['Sub Work()', ''], 2, { insertSpaces: true, tabSize: 2 }).text, '  \nEnd Sub');
});
test('conditions are formatted without touching calls, comments, dates or existing parentheses', () => {
  const lines = ['If(value = "Then")then', 'If (a) And (b) Then', 'If ((ready)) Then', "If ready Then ' keep this", 'If ready Then Rem do not change', 'If ready Then Run 1', 'Call Work(value)', 'Work value', 'Debug.Print #1/1/2026#', 'End If'];
  const result = formatLines(lines).lines.map(s => s.trimStart());
  assert.equal(result[0], 'If (value = "Then") Then'); assert.equal(result[1], 'If ((a) And (b)) Then');
  assert.equal(result[2], 'If ((ready)) Then'); assert.equal(result[3], "If (ready) Then ' keep this");
  assert.equal(result[4], 'If (ready) Then Rem do not change'); assert.deepEqual(result.slice(5, 9), lines.slice(5, 9));
  assert.equal(formatLines(['If (broken Then']).lines[0], 'If (broken Then');
});
test('formatting is idempotent including directives and continuations', () => {
  const source = ['#If VBA7 Then', 'Sub Work( _', 'ByVal x As Long)', 'If x > 0 Then', 'Debug.Print "a""b", _', '"next"', 'End If', 'End Sub', '#Else', "' old runtime", '#End If'];
  const once = formatLines(source).lines; assert.deepEqual(formatLines(once).lines, once);
});
test('lexical mask retains offsets and does not parse escaped quotes, comments or date punctuation', () => {
  for (const line of ['Debug.Print "a""Then\'b"', "x = 1 ' If a Then", 'If ready Then Rem End Sub', 'x = #1/1/2026 12:30#', 'x = "unterminated'])
    assert.equal(codeMask(line).length, line.length);
  assert.equal(codeMask('Debug.Print "End Sub"').includes('End Sub'), false);
});
test('registered formatters restrict edits to the selected range and Enter retains the body cursor', () => {
  let document, range, typed, applied;
  const lines = ['Sub Work()', 'Debug.Print 1', 'End Sub'];
  const model = { getLinesContent: () => lines, getOptions: () => ({ tabSize: 4 }), getLanguageId: () => 'vba', getLineMaxColumn: n => lines[n - 1].length + 1 };
  const editor = { onDidType: callback => typed = callback, getRawOptions: () => ({}), getModel: () => model,
    getPosition: () => ({ lineNumber: 2 }), getSelections: () => [{}], executeEdits: (...args) => applied = args, pushUndoStop() {} };
  const monaco = { Range: class { constructor(...values) { this.values = values; } }, Selection: class { constructor(...values) { this.values = values; } },
    languages: { registerDocumentFormattingEditProvider: (_, provider) => document = provider, registerDocumentRangeFormattingEditProvider: (_, provider) => range = provider } };
  installEditing(monaco, editor);
  assert.equal(document.provideDocumentFormattingEdits(model, {}).length, 1);
  assert.equal(range.provideDocumentRangeFormattingEdits(model, { startLineNumber: 3, endLineNumber: 3 }, {}).length, 0);
  lines[1] = ''; lines.pop(); typed('\n'); assert.equal(applied[1][1].text, '    \nEnd Sub'); assert.deepEqual(applied[2][0].values, [2, 5, 2, 5]);
  applied = null; typed('pasted\ntext'); assert.equal(applied, null);
  editor.getRawOptions = () => ({ readOnly: true }); typed('\n'); assert.equal(applied, null);
  editor.getRawOptions = () => ({}); editor.getSelections = () => [{}, {}]; typed('\n'); assert.equal(applied, null);
});
