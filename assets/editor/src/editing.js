// VBA formatting changes layout only; calls, literals and inline If statements retain their syntax.
export function codeMask(text) {
  let result = '', quoted = false, date = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quoted) {
      result += ' ';
      if (c === '"') { if (text[i + 1] === '"') { result += ' '; i++; } else quoted = false; }
    } else if (date) { result += ' '; if (c === '#') date = false; }
    else if (c === '"') { quoted = true; result += ' '; }
    else if (c === "'") { result += ' '.repeat(text.length - i); break; }
    else if (c === '#' && /[\d/]/.test(text[i + 1] || '')) { date = true; result += ' '; }
    else result += c;
  }
  const rem = /(?:^|:|\bThen\b)\s*Rem\b/i.exec(result);
  if (rem) { const start = rem.index + rem[0].search(/\bRem\b/i); result = result.slice(0, start) + ' '.repeat(result.length - start); }
  return result;
}

function block(text) {
  const code = codeMask(text).trim();
  if (/_\s*$/.test(code)) return null;
  const procedure = /^(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function|Property\s+(?:Get|Let|Set))\b/i.exec(code);
  if (procedure) { const kind = procedure[1].split(/\s/)[0].toLowerCase(); return { kind, close: 'End ' + kind[0].toUpperCase() + kind.slice(1) }; }
  if (/^#If\b.*\bThen\s*$/i.test(code)) return { kind: '#if', close: '#End If' };
  if (/^If\b.*\bThen\s*$/i.test(code)) return { kind: 'if', close: 'End If' };
  if (/^Select\s+Case\b/i.test(code)) return { kind: 'select', close: 'End Select' };
  const loop = /^For\b(?:\s+Each)?\s*([\p{L}_][\p{L}\p{N}_]*)?/iu.exec(code);
  if (loop) return { kind: 'for', close: 'Next' + (loop[1] ? ' ' + loop[1] : '') };
  for (const [kind, close] of [['do', 'Loop'], ['while', 'Wend'], ['with', 'End With'], ['type', 'End Type'], ['enum', 'End Enum']])
    if (new RegExp('^(?:(?:Public|Private)\\s+)?' + kind + '\\b', 'i').test(code)) return { kind, close };
  return null;
}

function closing(text) {
  const code = codeMask(text).trim();
  const end = /^(#?End)\s+(Sub|Function|Property|If|Select|With|Type|Enum)\b/i.exec(code);
  if (end) return (end[1][0] === '#' ? '#' : '') + end[2].toLowerCase();
  if (/^Next\b/i.test(code)) return 'for';
  if (/^Loop\b/i.test(code)) return 'do';
  if (/^Wend\b/i.test(code)) return 'while';
  return null;
}

function statements(text) {
  const code = codeMask(text).trim();
  // Everything following Then in a single-line If belongs to that conditional statement.
  return /^If\b.*\bThen\s*\S/i.test(code) ? [code] : code.split(/:(?!=)/).map(part => part.trim());
}

function closeBlocks(stack, code) {
  const kind = closing(code);
  if (!kind) return;
  const count = kind === 'for' ? Math.max(1, code.split(',').length) : 1;
  for (let n = 0; n < count; n++) { const at = stack.map(b => b.kind).lastIndexOf(kind); if (at >= 0) stack.splice(at); }
}

function unclosed(text) {
  const stack = [];
  for (const statement of statements(text)) {
    closeBlocks(stack, statement);
    const opener = block(statement); if (opener) stack.push(opener);
  }
  return stack.at(-1) || null;
}

function condition(text) {
  const mask = codeMask(text), match = /^(\s*)(If|ElseIf)\s*(.*?)\bThen\b/i.exec(mask);
  if (!match || !match[3].trim() || mask.slice(match[0].length).trim()) return text;
  const start = match[1].length + match[2].length, end = match[0].length - 4;
  const expression = text.slice(start, end).trim();
  let depth = 0, wrapped = expression[0] === '(';
  const code = codeMask(expression);
  for (let i = 0; i < code.length; i++) {
    if (code[i] === '(') depth++;
    if (code[i] === ')') depth--;
    if (depth < 0 || (depth === 0 && i < code.length - 1)) wrapped = false;
  }
  if (depth !== 0 || /_\s*$/.test(code)) return text;
  return match[1] + (match[2].toLowerCase() === 'if' ? 'If' : 'ElseIf') + ' ' + (wrapped ? expression : '(' + expression + ')') + ' Then' + text.slice(match[0].length);
}

export function formatLines(lines, options = {}) {
  const unit = options.insertSpaces === false ? '\t' : ' '.repeat(Math.max(1, options.tabSize || 4));
  const stack = [], result = [], indents = [];
  let continuation = false, logical = '';
  for (let index = 0; index < lines.length; index++) {
    const original = lines[index], code = codeMask(original).trim();
    if (!continuation) closeBlocks(stack, statements(code)[0]);
    const branch = /^(#?ElseIf|#?Else)\b/i.test(code), caseLine = /^Case\b/i.test(code);
    if (caseLine && stack.at(-1)?.kind === 'case') stack.pop();
    let depth = stack.length - (branch && stack.some(b => b.kind === (code[0] === '#' ? '#if' : 'if')) ? 1 : 0) + (continuation ? 1 : 0);
    const label = /^[\p{L}_][\p{L}\p{N}_]*\s*:(?![=])/u.test(code) || /^\d+\s/.test(code);
    depth = label ? 0 : Math.max(0, depth);
    const indent = unit.repeat(depth); indents.push(indent);
    result.push(original.trim() ? indent + condition(original.trimStart()) : '');
    logical += (logical ? ' ' : '') + code.replace(/_\s*$/, '').trim();
    const continued = /_\s*$/.test(code);
    if (!continued) {
      const parts = statements(logical);
      for (let part = 0; part < parts.length; part++) {
        if (part > 0) closeBlocks(stack, parts[part]);
        const opener = block(parts[part]); if (opener) stack.push(opener);
      }
      if (caseLine) stack.push({ kind: 'case' });
      logical = '';
    }
    continuation = continued;
  }
  return { lines: result, indents, nextIndent: unit.repeat(stack.length + (continuation ? 1 : 0)), unit };
}

function hasCloser(lines, start, opener) {
  let depth = 0;
  for (let i = start; i < lines.length; i++) {
    for (const statement of statements(lines[i])) {
      const next = block(statement), close = closing(statement);
      // A following procedure belongs to another body, rather than closing this one.
      if (next && ['sub', 'function', 'property'].includes(next.kind) && ['sub', 'function', 'property'].includes(opener.kind)) return false;
      if (next?.kind === opener.kind) depth++;
      if (close === opener.kind) { if (depth === 0) return true; depth--; }
      if (close && ['sub', 'function', 'property'].includes(close) && !['sub', 'function', 'property'].includes(opener.kind)) return false;
    }
  }
  return false;
}

export function enterPlan(lines, lineNumber, options = {}) {
  const previous = lineNumber - 2, current = lineNumber - 1;
  if (previous < 0 || current >= lines.length || lines[current].trim()) return null;
  const formatted = formatLines(lines.slice(0, current), options);
  let start = previous;
  while (start > 0 && /_\s*$/.test(codeMask(lines[start - 1]))) start--;
  const statement = lines.slice(start, previous + 1).map(line => codeMask(line).trim().replace(/_\s*$/, '')).join(' ');
  const header = formatted.lines[previous], opener = /_\s*$/.test(codeMask(lines[previous])) ? null : unclosed(statement);
  const closure = opener && !hasCloser(lines, current + 1, opener) ? '\n' + formatted.nextIndent.slice(0, -formatted.unit.length) + opener.close : '';
  return { previous, current, header, text: formatted.nextIndent + closure, column: formatted.nextIndent.length + 1 };
}

export function installEditing(monaco, editor) {
  const edits = (model, options, range) => {
    const original = model.getLinesContent(), formatted = formatLines(original, options).lines, result = [];
    for (let i = 0; i < original.length; i++)
      if ((!range || i + 1 >= range.startLineNumber && i + 1 <= range.endLineNumber) && original[i] !== formatted[i])
        result.push({ range: new monaco.Range(i + 1, 1, i + 1, original[i].length + 1), text: formatted[i] });
    return result;
  };
  monaco.languages.registerDocumentFormattingEditProvider('vba', { provideDocumentFormattingEdits: (model, options) => edits(model, options) });
  monaco.languages.registerDocumentRangeFormattingEditProvider('vba', { provideDocumentRangeFormattingEdits: (model, range, options) => edits(model, options, range) });
  editor.onDidType(text => {
    if (text !== '\n' && text !== '\r\n' || editor.getRawOptions().readOnly) return;
    const model = editor.getModel(), position = editor.getPosition();
    if (!model || model.getLanguageId() !== 'vba' || editor.getSelections().length !== 1) return;
    const plan = enterPlan(model.getLinesContent(), position.lineNumber, model.getOptions());
    if (!plan) return;
    editor.executeEdits('vbai.enter', [
      { range: new monaco.Range(plan.previous + 1, 1, plan.previous + 1, model.getLineMaxColumn(plan.previous + 1)), text: plan.header },
      { range: new monaco.Range(plan.current + 1, 1, plan.current + 1, model.getLineMaxColumn(plan.current + 1)), text: plan.text }
    ], [new monaco.Selection(position.lineNumber, plan.column, position.lineNumber, plan.column)]);
    editor.pushUndoStop();
  });
}
