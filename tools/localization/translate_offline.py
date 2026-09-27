"""Build UI catalogues using already-downloaded Argos models; no network code.

Requires ctranslate2 and sentencepiece in a development-only Python environment.
Models and generated audit reports stay under ignored artifacts/, never in the DLL.
"""
import argparse
import hashlib
import json
import os
import pathlib
import re
import xml.etree.ElementTree as ET

import ctranslate2
import sentencepiece

LANGUAGES = dict(es='Spanish', de='German', pt='Portuguese', it='Italian',
                 ja='Japanese', ko='Korean', zh='ChineseSimplified', zt='ChineseTraditional',
                 ru='Russian', ar='Arabic', hi='Hindi')
TECHNICAL = re.compile(
    r'https?://[^\s]+|\{\d+[^}]*\}|(?:CODEXVBE|AZURE|AWS|OPENAI)_[A-Z_]+|'
    r'\b[a-z]+_[a-z_]+\b|/chat/completions|\.vba|\.frx|\.frm|\.bas|\.cls|'
    r'\b(?:VBAi|VBA|VBE|GitHub Copilot|GitHub|Git Credential Manager|Git|CodexVBA|Codex|'
    r'OpenAI|ChatGPT|Copilot|Claude|Bedrock|Microsoft Entra|Microsoft|Windows|Ollama|'
    r'SQLite|Markdown|HTTPS|HTTP|API|UTF-8|COM|IAM|Bearer|Sub|Function|Property|'
    r'Fetch|Push|Pull|ours|theirs|text|HEAD)\b|\d+(?:[,.]\d+)*|[↑↓■×]|[\r\n]+')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('language', choices=LANGUAGES)
    parser.add_argument('--root', type=pathlib.Path, default=pathlib.Path(__file__).resolve().parents[2])
    args = parser.parse_args()
    root = args.root.resolve()
    os.chdir(root)
    folder = root / 'src/CodexVBE/Localization'
    tree = ET.parse(folder / 'UiStrings.resx')
    entries = [(node.attrib['name'], node.findtext('value') or '') for node in tree.getroot().findall('data')]
    output = folder / ('UiStrings' + LANGUAGES[args.language] + '.resx')
    overrides = json.loads((root / 'tools/localization/overrides.json').read_text(encoding='utf-8'))[args.language]
    if args.language == 'zt':
        # The direct model produces repetition; use local script/terminology conversion instead.
        from opencc import OpenCC
        converter = OpenCC('s2twp')
        source = {n.attrib['name']: n.findtext('value') for n in ET.parse(folder/'UiStringsChineseSimplified.resx').getroot().findall('data')}
        for node in tree.getroot().findall('data'):
            key = node.attrib['name']
            node.find('value').text = overrides.get(key, converter.convert(source[key]))
        ET.indent(tree, space='  ')
        tree.write(output, encoding='utf-8', xml_declaration=True)
        print('DONE zt', len(entries), 'entries; local OpenCC s2twp conversion', flush=True)
        return
    existing = {n.attrib['name']: n.findtext('value') for n in ET.parse(output).getroot().findall('data')} if output.exists() else {}
    model = next((root / 'artifacts/localization/models' / args.language).glob('*/sentencepiece.model'))
    tokenizer = sentencepiece.SentencePieceProcessor(model_proto=model.read_bytes())
    translator = ctranslate2.Translator(str((model.parent / 'model').relative_to(root)), device='cpu', compute_type='int8', intra_threads=2)
    cache = {}

    def translate_many(texts):
        pending = list(dict.fromkeys(s.strip() for s in texts if s.strip() and s.strip() not in cache))
        for offset in range(0, len(pending), 24):
            batch = pending[offset:offset+24]
            results = translator.translate_batch([tokenizer.encode(s, out_type=str) for s in batch],
                                                beam_size=4, max_decoding_length=256, replace_unknowns=True, no_repeat_ngram_size=3)
            for source, result in zip(batch, results):
                cache[source] = tokenizer.decode(result.hypotheses[0]).replace("▁", " ").strip()
            print(args.language, min(offset+24, len(pending)), '/', len(pending), flush=True)

    def spaced(text):
        if not text.strip(): return text
        left = text[:len(text)-len(text.lstrip())]
        right = text[len(text.rstrip()):]
        return left + cache[text.strip()] + right

    translate_many(value for key, value in entries if key not in overrides and key not in existing)
    retry = []
    results = {}
    for key, source in entries:
        if key in overrides:
            results[key] = overrides[key]
            continue
        if key in existing:
            results[key] = existing[key].replace("▁", " ")
            continue
        result = spaced(source)
        # Keep commands, placeholders, versions, brands and line breaks byte-for-byte.
        if any(result.count(token) < source.count(token) for token in TECHNICAL.findall(source)) or not result.strip():
            parts = TECHNICAL.split(source)
            translate_many(parts)
            result = ''
            start = 0
            for match in TECHNICAL.finditer(source):
                result += spaced(source[start:match.start()]) + match.group()
                start = match.end()
            result += spaced(source[start:])
            retry.append(key)
        if not result.strip(): raise ValueError('Empty translation: ' + key)
        results[key] = result

    for node in tree.getroot().findall('data'):
        node.find('value').text = results[node.attrib['name']]
    ET.indent(tree, space='  ')
    output = folder / ('UiStrings' + LANGUAGES[args.language] + '.resx')
    tree.write(output, encoding='utf-8', xml_declaration=True)
    audit = dict(language=args.language, model=json.loads((model.parent/'metadata.json').read_text()),
                 model_sha256=hashlib.sha256((model.parent/'model/model.bin').read_bytes()).hexdigest(),
                 entries=len(entries), overrides=len(overrides), protected_term_repairs=retry)
    (root/'artifacts/localization'/('audit-'+args.language+'.json')).write_text(json.dumps(audit,ensure_ascii=False,indent=2),encoding='utf-8')
    print('DONE', args.language, len(entries), 'entries;', len(retry), 'protected-term repairs', flush=True)


if __name__ == '__main__':
    main()
