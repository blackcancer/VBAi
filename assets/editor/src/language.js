// Project declarations are supplied by the host from immutable VBA snapshots.
export function installLanguage(monaco, editor, models, send) {
  let sequence = 0;
  const pending = new Map();
  const definitions = new Map();
  function request(model) {
    const entry = [...models.values()].find(e => e.model === model);
    if (!entry) return Promise.resolve(null);
    const version = model.getVersionId();
    if (entry.languageVersion === version && Date.now() - entry.languageTime < 1500) return entry.languagePromise;
    const request = ++sequence; entry.languageVersion = version; entry.languageTime = Date.now();
    return entry.languagePromise = new Promise(resolve => {
      const timeout = setTimeout(() => { pending.delete(request); resolve(null); }, 10000);
      pending.set(request, data => { clearTimeout(timeout); resolve(!model.isDisposed() && model.getVersionId() === version ? data : null); });
      send({ type: 'language', id: entry.id, version, request });
    });
  }
  function visible(data, position, prefix) {
    if (!data) return [];
    const all = data.symbols, local = all.filter(s => s.Module === data.module);
    const procedure = local.find(s => ['Procedure', 'Property'].includes(s.Kind) && s.Line <= position.lineNumber && s.EndLine >= position.lineNumber);
    let symbols = all.filter(s => !s.Conditional && (s.Module === data.module || !s.Private) &&
      (s.Scope === 'Module' || (s.Module === data.module && s.Scope === procedure?.Name)));
    const match = /([\p{L}_][\p{L}\p{N}_]*)\s*\.\s*[\p{L}\p{N}_]*$/u.exec(prefix);
    if (match) {
      const receiver = match[1].toLowerCase();
      const variable = symbols.find(s => s.Name.toLowerCase() === receiver && s.Scope === procedure?.Name) || symbols.find(s => s.Name.toLowerCase() === receiver);
      const type = variable?.TypeName?.split('.').pop()?.toLowerCase();
      return all.filter(s => !s.Conditional && (!s.Private || s.Module === data.module) &&
        ((s.Module.toLowerCase() === (type || receiver) && s.Scope === 'Module' && !['Module', 'Class'].includes(s.Kind)) ||
         (s.Scope?.toLowerCase() === (type || receiver) && ['Field', 'EnumMember'].includes(s.Kind))));
    }
    // Local declarations shadow module and project declarations.
    symbols = symbols.filter(s => !s.External && (s.Module === data.module || ['Module', 'Class'].includes(s.Kind) || !all.some(t => t.Name === s.Module && t.Kind === 'Class')));
    symbols.sort((a, b) => Number(b.Scope === procedure?.Name) - Number(a.Scope === procedure?.Name) || Number(b.Module === data.module) - Number(a.Module === data.module));
    const seen = new Set();
    return symbols.filter(s => { const key = s.Name.toLowerCase(); if (seen.has(key)) return false; seen.add(key); return true; });
  }
  const range = s => new monaco.Range(s.Line, s.Column, s.Line, s.Column + s.Name.length);
  const kind = s => ({ Procedure: 1, Property: 9, Variable: 4, Parameter: 4, Constant: 14, Module: 8, Class: 5, Type: 5, Enum: 15, EnumMember: 16, Field: 3 }[s.Kind] ?? 4);
  async function resolve(model, position) {
    const data = await request(model), word = model.getWordAtPosition(position);
    if (!word) return null;
    const prefix = model.getLineContent(position.lineNumber).slice(0, word.endColumn - 1);
    const symbol = visible(data, position, prefix).find(s => s.Name.toLowerCase() === word.word.toLowerCase());
    return symbol ? { data, symbol } : null;
  }
  monaco.languages.registerCompletionItemProvider('vba', {
    triggerCharacters: ['.'],
    async provideCompletionItems(model, position) {
      const data = await request(model), word = model.getWordUntilPosition(position);
      const prefix = model.getLineContent(position.lineNumber).slice(0, position.column - 1);
      return { suggestions: visible(data, position, prefix).map(s => ({ label: s.Name, kind: kind(s), detail: s.Declaration, insertText: s.Name,
        range: new monaco.Range(position.lineNumber, word.startColumn, position.lineNumber, word.endColumn) })) };
    }
  });
  monaco.languages.registerHoverProvider('vba', { async provideHover(model, position) {
    const found = await resolve(model, position); if (!found) return null;
    return { contents: [{ value: '```vba\n' + found.symbol.Declaration.replace(/`/g, '') + '\n```' }] };
  } });
  monaco.languages.registerDefinitionProvider('vba', { async provideDefinition(model, position) {
    const found = await resolve(model, position); if (!found || found.symbol.External) return null;
    if (found.symbol.Module === found.data.module) return { uri: model.uri, range: range(found.symbol) };
    const source = found.data.sources.find(s => s.Module === found.symbol.Module); if (!source) return null;
    const uri = monaco.Uri.parse('vbai-definition://module/' + encodeURIComponent(found.data.id) + '/' + encodeURIComponent(found.symbol.Module));
    let definition = definitions.get(uri.toString());
    if (!definition) { definition = monaco.editor.createModel(source.Text, 'vba', uri); definitions.set(uri.toString(), definition); }
    else if (definition.getValue() !== source.Text) definition.setValue(source.Text);
    return { uri, range: range(found.symbol) };
  } });
  monaco.editor.registerEditorOpener({ openCodeEditor(source, resource, selection) {
    if (resource.scheme !== 'vbai-definition') return false;
    const parts = resource.path.split('/');
    send({ type: 'definition', id: decodeURIComponent(parts[1]), module: decodeURIComponent(parts[2]), line: selection?.startLineNumber || selection?.lineNumber || 1, column: selection?.startColumn || selection?.column || 1 });
    return true;
  } });
  monaco.languages.registerSignatureHelpProvider('vba', { signatureHelpTriggerCharacters: ['(', ','], async provideSignatureHelp(model, position) {
    const prefix = model.getLineContent(position.lineNumber).slice(0, position.column - 1);
    // Walk backward to the current invocation, ignoring nested arguments and quoted commas.
    let depth = 0, quoted = false, count = 0, opening = -1;
    for (let i = prefix.length - 1; i >= 0; i--) {
      const c = prefix[i]; if (c === '"') { quoted = !quoted; continue; } if (quoted) continue;
      if (c === ')') depth++; else if (c === '(') { if (!depth) { opening = i; break; } depth--; } else if (c === ',' && !depth) count++;
    }
    if (opening < 0) return null;
    const name = /([\p{L}_][\p{L}\p{N}_]*)\s*$/u.exec(prefix.slice(0, opening)); if (!name) return null;
    const data = await request(model), s = visible(data, position, prefix.slice(0, opening)).find(s => s.Name.toLowerCase() === name[1].toLowerCase() && s.Parameters.length);
    if (!s) return null;
    return { value: { signatures: [{ label: s.Name + '(' + s.Parameters.join(', ') + ')', parameters: s.Parameters.map(label => ({ label })) }], activeSignature: 0, activeParameter: Math.min(count, s.Parameters.length - 1) }, dispose() {} };
  } });
  return { close(id) { for (const [uri, model] of definitions) if (model.uri.path.split('/')[1] === encodeURIComponent(id)) { model.dispose(); definitions.delete(uri); } }, reply(request, data) { const resolve = pending.get(request); if (resolve) { pending.delete(request); resolve(data); } }, async inspect(model, position) { return visible(await request(model), position, model.getLineContent(position.lineNumber).slice(0, position.column - 1)); } };
}
