#!/usr/bin/env python3
"""Build and validate the dependency-free VBAi public website.

Only allowlisted, hash-verified live screenshots and the original product icon
are published. No help staging, local evidence or application binary is copied.
"""
from __future__ import annotations

import argparse
import hashlib
import html
import json
import shutil
from html.parser import HTMLParser
from pathlib import Path
from string import Template
from urllib.parse import unquote, urlsplit

ROOT = Path(__file__).resolve().parents[2]
SITE = ROOT / 'site'
SCREENSHOTS = ROOT / 'docs/help/fr-FR/screenshots'
PAGES = ('home', 'installation', 'workflows', 'providers')
ICONS = {
    'spark': '<path d="M12 1C10 8 8 10 1 12c7 2 9 4 11 11 2-7 4-9 11-11-7-2-9-4-11-11Z"/>',
    'arrow': '<path d="M5 12h14m-6-6 6 6-6 6"/>',
    'chat': '<path d="M4 4h16v12H9l-5 4z"/><path d="M8 8h8m-8 4h5"/>',
    'code': '<path d="m8 7-5 5 5 5m8-10 5 5-5 5m-3-13-2 16"/>',
    'form': '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 9h18m-13 5h3m4 0h3m-10 3h10"/>',
    'git': '<circle cx="7" cy="5" r="2"/><circle cx="7" cy="19" r="2"/><circle cx="17" cy="7" r="2"/><path d="M7 7v10m10-8a8 8 0 0 1-10 8"/>',
    'check': '<rect x="4" y="3" width="16" height="18" rx="2"/><path d="m8 12 3 3 5-6"/>',
    'shield': '<path d="m12 3 8 3v6c0 4-8 9-8 9s-8-5-8-9V6z"/><path d="m8 12 3 3 5-6"/>',
}


def escape(value: str) -> str:
    return html.escape(str(value), quote=True)


def heading(value: str) -> str:
    """Allow the two editorial heading tags; escape all other input."""
    return escape(value).replace('&lt;br&gt;', '<br>').replace('&lt;em&gt;', '<em>').replace('&lt;/em&gt;', '</em>')


def icon(name: str) -> str:
    return f'<svg class="icon" viewBox="0 0 24 24" aria-hidden="true">{ICONS[name]}</svg>'


class Website:
    def __init__(self) -> None:
        self.config = json.loads((SITE / 'config.json').read_text(encoding='utf-8'))
        self.languages = {lang: json.loads((SITE / 'content' / f'{lang}.json').read_text(encoding='utf-8')) for lang in ('en', 'fr')}
        self.base = urlsplit(self.config['url']).path
        if not self.base.startswith('/') or not self.base.endswith('/'):
            raise ValueError('The configured website URL must have an absolute, trailing-slash path.')
        self.template = Template((SITE / 'template.html').read_text(encoding='utf-8'))
        self.captures = {item['file']: item for item in json.loads((SCREENSHOTS / 'manifest.json').read_text(encoding='utf-8'))['captures']}

    def relative(self, lang: str, page: str) -> str:
        return ('' if lang == 'en' else 'fr/') + ('' if page == 'home' else page + '/')

    def local(self, lang: str, page: str = 'home') -> str:
        return self.base + self.relative(lang, page)

    def url(self, lang: str, page: str) -> str:
        return self.config['url'] + self.relative(lang, page)

    def target(self, value: str, lang: str) -> str:
        return value.replace('{repo}', self.config['repository']).replace('{docs}', self.config['repository'] + '/blob/main/docs/').replace('{download}', self.config['download']).replace('{local}', self.local(lang))

    def link(self, label: str, target: str, lang: str, css: str = 'text-link') -> str:
        return f'<a class="{css}" href="{escape(self.target(target, lang))}">{escape(label)} {icon("arrow")}</a>'

    def figure(self, file: str, alt: str, caption: str, lang: str, css: str = '') -> str:
        if file not in self.config['screenshots']:
            raise ValueError(f'Screenshot is not allowlisted: {file}')
        hint = self.languages[lang]['common']['image_hint']
        return (f'<figure class="screenshot-frame {css}"><a class="screenshot-link" href="{self.base}assets/{file}" '
                f'data-screenshot="{escape(caption)}" data-framing="{self.framing_style(file)}" aria-label="{escape(alt + ". " + hint)}">'
                f'{self.framed_image(file, alt)}</a><figcaption>{escape(caption)} {escape(hint)}</figcaption></figure>')

    def framing_style(self, file: str) -> str:
        """Frame out captured Windows margins while preserving original image bytes."""
        image = self.captures[file]
        width, height = image['width'], image['height']
        left, top, right, bottom = self.config.get('framing', {}).get(file, [0, 0, width, height])
        if not (0 <= left < right <= width and 0 <= top < bottom <= height):
            raise ValueError(f'Invalid reviewed screenshot bounds: {file}')
        frame_width, frame_height = right - left, bottom - top
        return (f'--frame-aspect:{frame_width / frame_height:.8f};--image-width:{width / frame_width * 100:.8f}%;'
                f'--image-left:{-left / frame_width * 100:.8f}%;--image-top:{-top / frame_height * 100:.8f}%;'
                f'--native-width:{frame_width}px')

    def framed_image(self, file: str, alt: str, priority: bool = False) -> str:
        image = self.captures[file]
        loading = 'fetchpriority="high"' if priority else 'loading="lazy"'
        return (f'<span class="screenshot-viewport" style="{self.framing_style(file)}">'
                f'<img src="{self.base}assets/{file}" width="{image["width"]}" height="{image["height"]}" '
                f'alt="{escape(alt)}" {loading}></span>')

    def render_home(self, lang: str) -> str:
        data = self.languages[lang]['home']
        common = self.languages[lang]['common']
        doc = self.config['repository'] + '/blob/main/docs/'
        buttons = self.link(common['download'], '{download}', lang, 'button') + self.link(data['secondary'], '{local}workflows/', lang, 'button secondary')
        features = ''.join(
            f'<article class="feature-card"><div class="feature-marker"><span class="feature-index">0{index}</span><span class="feature-icon">{icon(item["icon"])}</span></div><h3>{escape(item["title"])}</h3>'
            f'<p>{escape(item["text"])}</p>{self.link(item["link"], doc + item["doc"], lang)}</article>' for index, item in enumerate(data['features'], 1))
        steps = ''.join(f'<li class="step"><span class="step-number">0{index}</span><div><h3>{escape(title)}</h3><p>{escape(text)}</p></div></li>'
                        for index, (title, text) in enumerate(data['steps'], 1))
        provider_chips = ''.join(f'<span class="provider-chip{" primary" if index == 0 else ""}">{escape(provider)}</span>'
                                 for index, provider in enumerate(('Codex', 'GitHub Copilot', 'OpenAI API', 'Claude', 'Gemini', 'Ollama', 'LM Studio')))
        chat_frame = self.framing_style('chat-first-implementation.png')
        form_frame = self.framing_style('form-runtime.png')
        git = self.figure('git-changes.png', data['git_alt'], data['git_caption'], lang)
        return f'''
<section class="hero-stage"><div class="container hero">
  <div class="hero-copy"><p class="eyebrow">{escape(data['eyebrow'])}</p><div class="hero-identity" aria-hidden="true"><img src="{self.base}assets/assistant.png" width="112" height="112" alt=""><span>VBA<span class="brand-ai">i</span></span>{icon('spark')}</div><h1>{heading(data['heading'])}</h1><p class="lead">{escape(data['lead'])}</p>
    <div class="actions">{buttons}</div><p class="hero-note"><a href="{self.local(lang, 'installation')}#download">{escape(data['note'])}</a></p></div>
  <figure class="hero-visual"><span class="hero-spark" aria-hidden="true">{icon('spark')}</span><div class="visual-label"><span>{escape(data['visual_label'])}</span><span><span class="visual-dot"></span>{escape(data['visual_status'])}</span></div>
    <a class="hero-chat screenshot-link" href="{self.base}assets/chat-first-implementation.png" data-screenshot="{escape(data['chat_caption'])}" data-framing="{chat_frame}" aria-label="{escape(data['chat_alt'] + '. ' + common['image_hint'])}">{self.framed_image('chat-first-implementation.png', data['chat_alt'], True)}</a>
    <a class="hero-form screenshot-link" href="{self.base}assets/form-runtime.png" data-screenshot="{escape(data['floating_label'])}" data-framing="{form_frame}" aria-label="{escape(data['form_alt'] + '. ' + common['image_hint'])}"><div class="floating-label">{escape(data['floating_label'])}</div>{self.framed_image('form-runtime.png', data['form_alt'])}</a>
    <figcaption>{escape(data['chat_caption'])}</figcaption></figure>
</div></section>
<div class="host-strip"><div class="container host-inner"><p class="host-label">{escape(data['host_label'])}</p><div class="host-list">{''.join('<span>' + escape(host) + '</span>' for host in data['hosts'])}</div><a class="host-label" href="{doc}compatibility.md">{escape(data['host_note'])}</a></div></div>
<section class="container section" id="features"><div class="section-heading"><div><p class="eyebrow">{escape(data['features_eyebrow'])}</p><h2>{heading(data['features_heading'])}</h2></div><p>{escape(data['features_text'])}</p></div><div class="feature-grid">{features}</div></section>
<section class="workflow-band"><div class="container section story-grid"><div class="story-text"><p class="eyebrow">{escape(data['story_eyebrow'])}</p><h2>{heading(data['story_heading'])}</h2><p>{escape(data['story_text'])}</p><ol class="steps">{steps}</ol>{self.link(data['story_link'], '{local}workflows/', lang)}<div class="macro-map" aria-hidden="true"><span>.xlsm / .swp</span>{icon('arrow')}<span>.bas · .cls<br>.frm + .frx</span>{icon('arrow')}<span>Git</span></div></div><div><p class="eyebrow">{escape(data['git_caption_title'])}</p>{git}<p class="provider-note" style="margin-top:16px">{escape(common['capture_note'])}</p></div></div></section>
<section class="container section"><div class="section-heading"><div><p class="eyebrow">{escape(data['providers_eyebrow'])}</p><h2>{heading(data['providers_heading'])}</h2></div><p>{escape(data['providers_text'])}</p></div><div class="provider-list">{provider_chips}</div><p class="provider-note">{escape(data['provider_note'])}</p>{self.link(data['provider_link'], '{local}providers/', lang)}</section>
<section class="download-stage"><div class="container getting-started"><img class="download-logo" src="{self.base}assets/assistant.png" width="180" height="180" alt="" loading="lazy"><div><h2>{escape(data['start_heading'])}</h2><p>{escape(data['start_text'])}</p></div><div class="actions">{self.link(common['download'], '{download}', lang, 'button')}{self.link(common['guide'], '{local}installation/', lang, 'button secondary')}</div></div></section>'''

    def render_article(self, lang: str, page: str) -> str:
        data = self.languages[lang][page]
        sections = []
        if 'table' in data:
            rows = ''.join('<tr>' + ''.join(f'<td>{escape(cell)}</td>' for cell in row) + '</tr>' for row in data['table'])
            headers = ''.join(f'<th scope="col">{escape(cell)}</th>' for cell in data['table_head'])
            sections.append(f'<table class="provider-table"><caption class="eyebrow">{escape(data["eyebrow"])}</caption><thead><tr>{headers}</tr></thead><tbody>{rows}</tbody></table>')
        for section in data['sections']:
            body = ''.join(f'<p>{escape(paragraph)}</p>' for paragraph in section['paragraphs'])
            if 'quote' in section:
                label = 'Example request' if lang == 'en' else 'Exemple de demande'
                body += f'<p class="eyebrow" style="margin-top:24px">{label}</p><blockquote>{escape(section["quote"])}</blockquote>'
            if 'notice' in section:
                title, text = section['notice']
                body += f'<div class="notice"><h3>{escape(title)}</h3><p>{escape(text)}</p></div>'
            if 'image' in section:
                css = 'form-example' if section.get('form') else 'compact'
                if section.get('narrow'):
                    css = 'portrait-example'
                body += self.figure(section['image'], section['image_alt'], section['caption'], lang, css)
            body += ''.join(self.link(label, target, lang) + ' ' for label, target in section.get('links', []))
            sections.append(f'<section id="{escape(section["id"])}"><h2>{escape(section["title"])}</h2>{body}</section>')
        links = ''.join(self.link(label, target, lang, '') for label, target in data['links'])
        return (f'<div class="container"><header class="page-intro"><p class="eyebrow">{escape(data["eyebrow"])}</p>'
                f'<h1>{heading(data["heading"])}</h1><p class="lead">{escape(data["lead"])}</p></header>'
                f'<div class="article-layout"><article class="article">{"".join(sections)}</article>'
                f'<aside class="aside-card"><h2>{escape(data["aside_title"])}</h2>{links}<p>{escape(data["aside_note"])}</p></aside></div></div>')

    def render(self, lang: str, page: str) -> str:
        language = self.languages[lang]
        data = language[page]
        other_lang = language['common']['other_lang']
        navigation = []
        for item in ('home', 'workflows', 'providers', 'installation'):
            anchor = '#features' if item == 'home' else ''
            current = 'aria-current="page"' if item == page else ''
            navigation.append(f'<a href="{self.local(lang, item)}{anchor}" {current}>{escape(language["nav"][item])}</a>')
        nav = ''.join(navigation)
        schema = ''
        if page == 'home':
            schema = '<script type="application/ld+json">' + json.dumps({
                '@context': 'https://schema.org', '@type': 'SoftwareApplication',
                'name': 'VBAi', 'url': self.config['url'], 'description': data['description'],
                'applicationCategory': 'DeveloperApplication', 'operatingSystem': 'Windows x64',
                'softwareVersion': self.config['version'], 'downloadUrl': self.config['download'],
                'license': 'https://github.com/blackcancer/VBAi/blob/main/LICENSING.md',
                'offers': {'@type': 'Offer', 'price': '0', 'priceCurrency': 'EUR'},
            }, ensure_ascii=False).replace('<', '\\u003c') + '</script>'
        if page == 'not_found':
            content = f'<div class="not-found"><h1>{escape(data["heading"])}</h1><p>{escape(data["text"])}</p>{self.link(data["link"], "{local}", lang, "button")}</div>'
        else:
            content = self.render_home(lang) if page == 'home' else self.render_article(lang, page)
        return self.template.substitute(
            {**{key: escape(value) for key, value in language['common'].items()},
             'lang': lang, 'base': self.base, 'title': escape(data['title']),
             'description': escape(data.get('description', data.get('text', ''))),
             'canonical': self.url(lang, page if page in PAGES else 'home'),
             'english': self.url('en', page if page in PAGES else 'home'),
             'french': self.url('fr', page if page in PAGES else 'home'),
             'social_image': self.config['url'] + 'assets/git-changes.png',
             'social_alt': escape(language['home']['git_alt']),
             'home': self.local(lang), 'other': self.local(other_lang, page if page in PAGES else 'home'),
             'repository': self.config['repository'], 'nav': nav, 'schema': schema,
             'content': content, 'arrow': icon('arrow')})


class PageParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.ids: set[str] = set()
        self.references: list[str] = []
        self.h1_count = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        data = dict(attrs)
        if tag == 'h1':
            self.h1_count += 1
        if 'id' in data:
            if data['id'] in self.ids:
                raise ValueError(f'Duplicate HTML ID: {data["id"]}')
            self.ids.add(data['id'])
        if tag == 'img' and 'alt' not in data:
            raise ValueError('Image without alt attribute')
        if tag == 'a' and data.get('href'):
            self.references.append(data['href'])
        if tag in ('img', 'script') and data.get('src'):
            self.references.append(data['src'])
        if tag == 'link' and data.get('href'):
            self.references.append(data['href'])


def validate(output: Path, website: Website) -> None:
    parsers = {}
    for path in output.rglob('*.html'):
        parser = PageParser()
        text = path.read_text(encoding='utf-8')
        if '\ufffd' in text or 'Ã' in text or 'Â' in text:
            raise ValueError(f'Unexpected encoding artifact: {path}')
        parser.feed(text)
        if parser.h1_count != 1:
            raise ValueError(f'Expected one main heading: {path}')
        parsers[path] = parser
    for path, parser in parsers.items():
        for target in parser.references:
            parts = urlsplit(target)
            if parts.scheme or parts.netloc:
                continue
            if parts.path and not parts.path.startswith(website.base):
                raise ValueError(f'Local link escapes project Pages path: {target}')
            destination = output / unquote(parts.path.removeprefix(website.base)) if parts.path else path
            if destination.is_dir():
                destination /= 'index.html'
            if not destination.is_file():
                raise ValueError(f'{path}: missing local target {target}')
            if parts.fragment and parts.fragment not in parsers[destination].ids:
                raise ValueError(f'{path}: missing anchor {target}')
    print(f'Website links: {len(parsers)} HTML pages; local files, anchors and image alternatives verified.')


def build(output: Path) -> None:
    output = output.resolve()
    if not any(output.is_relative_to(ROOT / folder) and output != ROOT / folder for folder in ('artifacts', 'dist')):
        raise ValueError('Use a dedicated output directory within artifacts/ or dist/.')
    for parent in (output, *output.parents):
        if parent.is_symlink() or (hasattr(parent, 'is_junction') and parent.is_junction()):
            raise ValueError('Output ancestors must not be reparse points.')
        if parent == ROOT:
            break
    website = Website()
    assets = output / 'assets'
    assets.mkdir(parents=True, exist_ok=True)
    for source in ('styles.css', 'main.js'):
        shutil.copyfile(SITE / source, assets / source)
    shutil.copyfile(ROOT / 'assets/icons/assistant.png', assets / 'assistant.png')
    for file in website.config['screenshots']:
        source = SCREENSHOTS / file
        item = website.captures[file]
        if item.get('synthetic') is not False or item.get('provenance') not in ('live-workflow', 'live-interface'):
            raise ValueError(f'Screenshot lacks original live provenance: {file}')
        if hashlib.sha256(source.read_bytes()).hexdigest() != item['sha256']:
            raise ValueError(f'Screenshot no longer matches its reviewed manifest: {file}')
        shutil.copyfile(source, assets / file)
    expected = {'assets/styles.css', 'assets/main.js', 'assets/assistant.png', '.nojekyll', 'sitemap.xml', '404.html'}
    expected.update('assets/' + file for file in website.config['screenshots'])
    for lang in website.languages:
        for page in PAGES:
            relative = website.relative(lang, page) + 'index.html'
            target = output / relative
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text(website.render(lang, page), encoding='utf-8', newline='\n')
            expected.add(relative)
    (output / '404.html').write_text(website.render('en', 'not_found'), encoding='utf-8', newline='\n')
    urls = [website.url(lang, page) for lang in website.languages for page in PAGES]
    (output / 'sitemap.xml').write_text('<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n' +
                                      ''.join(f'  <url><loc>{escape(url)}</loc></url>\n' for url in urls) + '</urlset>\n', encoding='utf-8')
    (output / '.nojekyll').write_text('', encoding='utf-8')
    unexpected = {path.relative_to(output).as_posix() for path in output.rglob('*') if path.is_file()} - expected
    if unexpected:
        raise ValueError(f'Unreviewed files in publication output: {sorted(unexpected)}')
    validate(output, website)
    print(f'Built {len(urls)} localized pages, 4 original screenshots and a 404 page in {output}.')


if __name__ == '__main__':
    arguments = argparse.ArgumentParser(description=__doc__)
    arguments.add_argument('--output', type=Path, default=ROOT / 'artifacts/site-preview/VBAi')
    build(arguments.parse_args().output)
