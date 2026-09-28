// Project declarations are supplied by the host from immutable VBA snapshots.
export function installLanguage(monaco, editor, models, send) {
  let sequence = 0;
  const pending = new Map();
  const definitions = new Map();
  const intrinsic = [
    { Name: 'Debug', Module: 'Global', Kind: 'Object', TypeName: 'VBA.Debug', Declaration: 'VBA.Debug', Documentation: 'VBA debugging output and assertions.' },
    { Name: 'Print', Module: 'Debug', Kind: 'Procedure', TypeName: '', Parameters: ['expression As Variant'], Declaration: 'Debug.Print expression', Documentation: 'Displays expressions in the Immediate window.' },
    { Name: 'Assert', Module: 'Debug', Kind: 'Procedure', TypeName: '', Parameters: ['condition As Boolean'], Declaration: 'Debug.Assert condition', Documentation: 'Suspends execution when the condition is False.' }
  ].map(s => ({ Scope: 'Module', External: true, Library: 'VBA', Parameters: [], ...s }));
  function request(model) {
    const entry = [...models.values()].find(e => e.model === model);
    if (!entry) return Promise.resolve(null);
    const version = model.getVersionId();
    // Coalesce simultaneous providers, but reread references on every subsequent request.
    if (entry.languageVersion === version && entry.languagePromise) return entry.languagePromise;
    const request = ++sequence; entry.languageVersion = version;
    const promise = new Promise(resolve => {
      const timeout = setTimeout(() => { pending.delete(request); resolve(null); }, 10000);
      pending.set(request, data => { clearTimeout(timeout); resolve(!model.isDisposed() && model.getVersionId() === version ? data : null); });
      send({ type: 'language', id: entry.id, version, request });
    });
    entry.languagePromise = promise;
    promise.finally(() => { if (entry.languagePromise === promise) entry.languagePromise = null; });
    return promise;
  }
  function visible(data, position, prefix, model) {
    if (!data) return [];
    const code = prefix.replace(/"(?:[^"]|"")*"/g, match => ' '.repeat(match.length));
    if (/["']/.test(code) || /(?:^|:|\bThen\b)\s*Rem\b/i.test(code)) return [];
    const implicitReceiver = /(^\s*|[=(,+*/&-]\s*)\./;
    if (implicitReceiver.test(prefix) && model) {
      const stack = [];
      for (let line = 1; line < position.lineNumber; line++) {
        const text = model.getLineContent(line).replace(/"(?:[^"]|"")*"/g, '""').replace(/'[^\n]*$/, '').trim();
        if (/^End\s+With\b/i.test(text)) stack.pop();
        else { const start = /^With\s+(.+)$/i.exec(text); if (start) stack.push(start[1].startsWith('.') ? (stack[stack.length - 1] || '') + start[1] : start[1]); }
      }
      if (stack.length) prefix = prefix.replace(implicitReceiver, (_, start) => start + stack[stack.length - 1] + '.');
    }
    const aliases = data.symbols.map(s => {
      const native = s.External && s.Library === 'VBA' && /^_B_(str|var)_(.+)$/.exec(s.Name);
      if (!native) return s;
      const name = native[2] + (native[1] === 'str' ? '$' : '');
      return { ...s, Name: name, Declaration: s.Declaration?.replace(s.Name, name) };
    });
    const all = [...aliases, ...intrinsic], local = all.filter(s => s.Module === data.module);
    const procedure = local.find(s => ['Procedure', 'Property'].includes(s.Kind) && s.Line <= position.lineNumber && s.EndLine >= position.lineNumber);
    let symbols = all.filter(s => !s.Conditional && (s.Module === data.module || !s.Private) &&
      (s.Scope === 'Module' || (s.Module === data.module && s.Scope === procedure?.Name)));
    // Resolve each receiver, including property/method calls, without evaluating VBA.
    const identifier = '[\\p{L}_][\\p{L}\\p{N}_]*\\$?';
    const expression = prefix.replace(/"(?:[^"]|"")*"/g, '""');
    let simplified = expression;
    for (let i = 0; i < 32 && /\([^()]*\)/.test(simplified); i++) simplified = simplified.replace(/\([^()]*\)/g, '§');
    const match = new RegExp('(' + identifier + '§?(?:\\s*\\.\\s*' + identifier + '§?)*)\\s*\\.\\s*[\\p{L}\\p{N}_]*\\$?$', 'u').exec(simplified);
    if (match) {
      const names = match[1].split('.').map(n => ({ name: n.trim().replace(/§/g, '').toLowerCase(), called: n.includes('§') }));
      const memberType = n => n?.split('.').pop()?.replace(/^_/, '').toLowerCase();
      const members = owner => all.filter(s => (!s.External || !owner.includes('.') || s.Library?.toLowerCase() === owner.split('.')[0].toLowerCase()) && !s.Conditional && (!s.Private || s.Module === data.module) &&
        ((memberType(s.Module) === memberType(owner) && s.Scope === 'Module' && !['Module', 'Library'].includes(s.Kind) && (s.Kind !== 'Class' || s.External)) ||
         (memberType(s.Scope) === memberType(owner) && ['Field', 'EnumMember'].includes(s.Kind))));
      const firstPart = names.shift(), first = firstPart.name;
      const variable = symbols.find(s => s.Name.toLowerCase() === first && s.Scope === procedure?.Name) || symbols.find(s => s.Name.toLowerCase() === first)
        || all.find(s => memberType(s.Module) === 'global' && s.Name.toLowerCase() === first);
      let type = first === 'me' ? data.module : variable?.TypeName || first;
      if (model && variable && /^(Object|Variant)$/i.test(type)) {
        const assignment = new RegExp('^\\s*Set\\s+' + first + '\\s*=\\s*(.*)', 'iu');
        for (let line = position.lineNumber - 1; line >= (procedure?.Line || 1); line--) {
          const declaration = assignment.exec(model.getLineContent(line));
          if (declaration) { const created = /^New\s+([\p{L}\p{N}_.]+)/iu.exec(declaration[1]); if (created) type = created[1]; break; }
        }
      }
      const qualify = owner => owner.includes('.') || all.some(s => !s.External && s.Kind === 'Class' && s.Name.toLowerCase() === owner.toLowerCase()) ? owner :
        all.find(s => s.External && ['Class', 'Type'].includes(s.Kind) && memberType(s.Name) === memberType(owner))?.TypeName || owner;
      type = qualify(type);
      const indexedType = owner => members(owner).find(s => s.DefaultMember || s.Name.toLowerCase() === 'item')?.TypeName || owner;
      if (firstPart.called && !['Procedure', 'Property'].includes(variable?.Kind)) type = indexedType(type);
      for (const part of names) {
        const member = members(type).sort((a, b) => Number(a.TypeName === 'Void') - Number(b.TypeName === 'Void')).find(s => s.Name.toLowerCase() === part.name);
        if (!member?.TypeName) return [];
        type = qualify(member.TypeName);
        if (part.called && member.Kind === 'Property' && !member.Parameters?.length) type = indexedType(type);
      }
      const seen = new Set();
      return members(type).sort((a, b) => Number(a.TypeName === 'Void') - Number(b.TypeName === 'Void')).filter(s => { const key = s.Name.toLowerCase(); if (seen.has(key)) return false; seen.add(key); return true; });
    }
    // Local declarations shadow module and project declarations.
    symbols = symbols.filter(s => s.External ? s.Global || ['Library', 'Class', 'Type', 'Enum'].includes(s.Kind) || memberTypeName(s.Module) === 'global' :
      (s.Module === data.module || ['Module', 'Class'].includes(s.Kind) || !all.some(t => t.Name === s.Module && t.Kind === 'Class')));
    symbols.sort((a, b) => Number(b.Scope === procedure?.Name) - Number(a.Scope === procedure?.Name) || Number(b.Module === data.module) - Number(a.Module === data.module));
    const seen = new Set();
    return symbols.filter(s => { const key = s.Name.toLowerCase(); if (seen.has(key)) return false; seen.add(key); return true; });
  }
  const memberTypeName = value => value?.split('.').pop()?.replace(/^_/, '').toLowerCase();
  const markdown = value => (value || '').replace(/[\\`*_{}\[\]<>]/g, '\\$&');
  function documentation(symbol) {
    const parts = ['```vba\n' + (symbol.Declaration || symbol.Name).replace(/`/g, '') + '\n```'];
    if (symbol.Documentation) parts.push(markdown(symbol.Documentation));
    if (symbol.External) {
      parts.push('Library: **' + markdown(symbol.Library) + '**');
      if (symbol.LibraryDescription) parts.push(markdown(symbol.LibraryDescription));
      if (symbol.LibraryPath) parts.push('`' + symbol.LibraryPath.replace(/`/g, '') + '`');
      if (symbol.HelpFile) parts.push('Help: ' + markdown(symbol.HelpFile) + (symbol.HelpContext ? ' (' + symbol.HelpContext + ')' : ''));
    }
    return parts.map(value => ({ value, isTrusted: false }));
  }
  const range = s => new monaco.Range(s.Line, s.Column, s.Line, s.Column + s.Name.length);
  const kind = s => ({ Procedure: 1, Property: 9, Variable: 4, Parameter: 4, Constant: 14, Module: 8, Class: 5, Type: 5, Enum: 15, EnumMember: 16, Field: 3 }[s.Kind] ?? 4);
  async function resolve(model, position) {
    const data = await request(model), word = model.getWordAtPosition(position);
    if (!word) return null;
    const prefix = model.getLineContent(position.lineNumber).slice(0, word.endColumn - 1);
    const symbol = visible(data, position, prefix, model).find(s => s.Name.toLowerCase() === word.word.toLowerCase());
    return symbol ? { data, symbol } : null;
  }
  monaco.languages.registerCompletionItemProvider('vba', {
    triggerCharacters: ['.'],
    async provideCompletionItems(model, position) {
      const data = await request(model), word = model.getWordUntilPosition(position);
      const prefix = model.getLineContent(position.lineNumber).slice(0, position.column - 1);
      return { suggestions: visible(data, position, prefix, model).map(s => ({ label: s.Name, kind: kind(s), detail: s.Declaration, insertText: s.Name,
        documentation: { value: documentation(s).map(part => part.value).join('\n\n'), isTrusted: false },
        range: new monaco.Range(position.lineNumber, word.startColumn, position.lineNumber, word.endColumn) })) };
    }
  });
  async function provideHover(model, position) {
    const found = await resolve(model, position); if (!found) return null;
    const word = model.getWordAtPosition(position);
    return { range: new monaco.Range(position.lineNumber, word.startColumn, position.lineNumber, word.endColumn), contents: documentation(found.symbol) };
  }
  monaco.languages.registerHoverProvider('vba', { provideHover });
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
    const name = /([\p{L}_][\p{L}\p{N}_]*\$?)\s*$/u.exec(prefix.slice(0, opening)); if (!name) return null;
    const data = await request(model), s = visible(data, position, prefix.slice(0, opening), model).find(s => s.Name.toLowerCase() === name[1].toLowerCase() && s.Parameters.length);
    if (!s) return null;
    return { value: { signatures: [{ label: s.Declaration || s.Name + '(' + s.Parameters.join(', ') + ')', documentation: s.Documentation, parameters: s.Parameters.map(label => ({ label })) }], activeSignature: 0, activeParameter: Math.min(count, s.Parameters.length - 1) }, dispose() {} };
  } });
  return { hover: provideHover, close(id) { for (const [uri, model] of definitions) if (model.uri.path.split('/')[1] === encodeURIComponent(id)) { model.dispose(); definitions.delete(uri); } }, reply(request, data) { const resolve = pending.get(request); if (resolve) { pending.delete(request); resolve(data); } }, async inspect(model, position) { return visible(await request(model), position, model.getLineContent(position.lineNumber).slice(0, position.column - 1), model); } };
}
