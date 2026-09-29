#!/usr/bin/env python3
"""Check maintained Markdown structure and local links without dependencies.

Supports the inline links, ATX headings and HTML image sources used in this repo;
it is not a complete CommonMark parser. External URLs are never requested.
Test-input Markdown is deliberately excluded from editorial checks.
"""
from __future__ import annotations

import argparse
import html
import json
import re
import sys
from pathlib import Path
from urllib.parse import unquote, urlsplit

IGNORED_DIRS = {'.git', 'artifacts', 'bin', 'obj', 'node_modules', '__pycache__'}
FIXTURE_PREFIXES = (
    'tests/Infrastructure/.agents/',
    'tests/VBAi.Tests/Infrastructure/Fixtures/',
)
HEADING = re.compile(r'^ {0,3}(#{1,6})\s+(.+?)\s*#*\s*$')
LINK = re.compile(r'!?\[[^\]\n]*\]\((<[^>\n]+>|[^\s)]+)(?:\s+"[^"\n]*")?\)')
HTML_SOURCE = re.compile(r'<(?:img|a)\b[^>]*?\b(?:src|href)=["\']([^"\']+)["\']', re.I)


def prose(text: str) -> str:
    """Remove YAML front matter, fenced examples and HTML comments."""
    lines = text.splitlines()
    out: list[str] = []
    front = bool(lines and lines[0] == '---')
    fence = ''
    fence_size = 0
    for index, line in enumerate(lines):
        if front:
            if index and line in ('---', '...'):
                front = False
            out.append('')
            continue
        marker = re.match(r'^ {0,3}(`{3,}|~{3,})', line)
        if not fence and marker:
            fence, fence_size = marker[1][0], len(marker[1])
            out.append('')
            continue
        if fence:
            if re.match(r'^ {0,3}' + re.escape(fence) + r'{' + str(fence_size) + r',}\s*$', line):
                fence = ''
            out.append('')
            continue
        out.append(line)
    return re.sub(r'<!--[\s\S]*?-->', '', '\n'.join(out))


def anchors(text: str) -> set[str]:
    """Generate heading anchors, including repeated-heading suffixes."""
    used: set[str] = set()
    for line in prose(text).splitlines():
        match = HEADING.match(line)
        if not match:
            continue
        title = re.sub(r'<[^>]*>', '', html.unescape(match[2]))
        title = re.sub(r'\[([^\]]+)\]\([^)]*\)', r'\1', title)
        slug = re.sub(r'[^\w\- ]', '', title.lower()).replace(' ', '-')
        candidate, counter = slug, 0
        while candidate in used:
            counter += 1
            candidate = f'{slug}-{counter}'
        used.add(candidate)
    return used


def maintained_files(root: Path) -> list[Path]:
    return sorted(p for p in root.rglob('*.md')
                  if not any(part in IGNORED_DIRS for part in p.relative_to(root).parts)
                  and not p.relative_to(root).as_posix().startswith(FIXTURE_PREFIXES))


def check(root: Path, known_paths: set[str] | None = None) -> tuple[int, int, list[str]]:
    """Return the number of pages, local links and validation errors.

    known_paths is useful only for an explicitly supplied sparse snapshot. It
    establishes path existence, not the contents of missing Markdown targets.
    """
    root = root.resolve()
    known_paths = known_paths or set()
    paths = maintained_files(root)
    texts: dict[Path, str] = {}
    errors: list[str] = []
    link_count = 0
    for path in paths:
        name = path.relative_to(root).as_posix()
        try:
            text = path.read_text(encoding='utf-8')
        except (OSError, UnicodeError) as exc:
            errors.append(f'{name}: cannot read UTF-8: {exc}')
            continue
        texts[path] = text
        clean = prose(text)
        if sum(bool(re.match(r'^#\s+', line)) for line in clean.splitlines()) != 1:
            errors.append(f'{name}: expected exactly one level-one heading')
        if not text.endswith('\n'):
            errors.append(f'{name}: missing final newline')
        targets = [m[1].strip('<>') for m in LINK.finditer(clean)]
        targets.extend(m[1] for m in HTML_SOURCE.finditer(clean))
        for target in targets:
            parts = urlsplit(html.unescape(target))
            if parts.scheme or parts.netloc:
                continue
            link_count += 1
            local_path = unquote(parts.path)
            destination = ((root / local_path.lstrip('/')) if local_path.startswith('/')
                           else (path.parent / local_path if local_path else path)).resolve()
            try:
                relative = destination.relative_to(root).as_posix()
            except ValueError:
                errors.append(f'{name}: link escapes the repository: {target}')
                continue
            if not destination.exists() and relative not in known_paths:
                errors.append(f'{name}: missing local destination: {target}')
                continue
            if parts.fragment and destination.suffix.lower() == '.md':
                try:
                    content = texts.get(destination)
                    if content is None:
                        content = destination.read_text(encoding='utf-8')
                    fragment = unquote(parts.fragment)
                    if fragment not in anchors(content):
                        errors.append(f'{name}: missing heading anchor: {target}')
                except (OSError, UnicodeError):
                    errors.append(f'{name}: cannot inspect Markdown anchor: {target}')
    return len(paths), link_count, errors


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path, default=Path(__file__).resolve().parents[2])
    parser.add_argument('--known-paths', type=Path,
                        help='Optional JSON list of verified paths for a sparse snapshot.')
    args = parser.parse_args()
    try:
        known = json.loads(args.known_paths.read_text(encoding='utf-8')) if args.known_paths else []
        if not isinstance(known, list) or not all(isinstance(p, str) for p in known):
            raise ValueError('--known-paths must contain a JSON array of repository-relative strings')
        if not args.root.is_dir():
            raise ValueError('--root must be an existing directory')
        pages, links, errors = check(args.root, set(known))
    except (OSError, ValueError) as exc:
        print(f'Documentation check failed: {exc}', file=sys.stderr)
        return 2
    for error in errors:
        print(error, file=sys.stderr)
    print(f'{pages} maintained Markdown files; {links} local links; {len(errors)} errors.')
    return 1 if errors else 0


if __name__ == '__main__':
    raise SystemExit(main())
