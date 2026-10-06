#!/usr/bin/env python3
"""Build the French HTML Help manual from reviewed topics and real UI captures.

The compiler is an explicit local dependency, never downloaded or installed here.
Generated HTML remains available for review even when no compiler is supplied.
"""
from __future__ import annotations

import argparse
import hashlib
import html
import json
import re
import shutil
import subprocess
from pathlib import Path


STYLE = """body{font:15px 'Segoe UI',Arial,sans-serif;color:#27364a;background:#fff;
margin:0;line-height:1.6}.masthead{background:#182a40;color:white;padding:16px 30px}
.masthead a{color:#d7e8fa}.content{max-width:1040px;margin:24px auto;padding:0 30px}
h1{font-size:30px;line-height:1.2;margin:12px 0 18px}h2{font-size:22px;
color:#174c7c;margin:30px 0 14px}h3{font-size:17px}a{color:#075ea8}
li{margin:8px 0}code{font-family:Consolas,monospace;background:#f2f5f8;padding:2px 5px}
table{border-collapse:collapse;width:100%;margin:16px 0}th,td{padding:10px 14px;
border:1px solid #dce4ed;text-align:left;vertical-align:top}th{background:#eef4fa}
.lead{font-size:18px;color:#41556c;max-width:850px}.section{clear:both;padding-top:1px}
.section:after{content:'';display:block;clear:both}.note{background:#edf5fc;
border-left:4px solid #2474ad;padding:12px 16px;margin:18px 0;overflow:hidden}
.figure{margin:18px 0 24px;max-width:100%}.figure.side{float:right;width:43%;margin:4px 0 20px 26px}
.figure.compact{width:420px}.picture{position:relative;display:block;border:1px solid #cad5e1;
background:#f7f9fc;max-width:100%;box-shadow:0 3px 12px #dbe2e9}
.picture img{display:block;width:100%;height:auto;border:0}.mark{position:absolute;
border:2px solid #a94311;box-sizing:border-box;pointer-events:none}.number{
position:absolute;left:-10px;top:-12px;background:#a94311;color:#fff;border:2px solid #fff;
font:bold 12px 'Segoe UI',Arial;border-radius:14px;width:22px;height:22px;text-align:center;
line-height:22px}.caption{font-size:13px;color:#536779;margin:8px 0}.callouts{font-size:14px;
padding-left:25px}.callouts li{padding-left:3px}.footer{clear:both;border-top:1px solid #dce4ed;
color:#536779;font-size:13px;margin-top:32px;padding:18px 0}.related{clear:both}
@media screen and (max-width:760px){.content{padding:0 20px}.figure.side{float:none;width:auto;
margin:18px 0}.figure{max-width:100%}}
"""


def read_json(path: Path):
    """Read UTF-8 source data, accepting a BOM emitted by Windows PowerShell."""
    return json.loads(path.read_text(encoding="utf-8-sig"))


def validate(manual: dict, source: Path, captures: dict) -> None:
    """Reject duplicate topic IDs, dangling links and missing or altered captures."""
    topics = manual["topics"]
    ids = [topic["id"] for topic in topics]
    if any(not re.fullmatch(r"[a-z][a-z0-9-]*", key) for key in ids):
        raise ValueError("Topic IDs must be plain local identifiers")
    if len(ids) != len(set(ids)) or "start" not in ids:
        raise ValueError("Topic IDs must be unique and include start")
    assigned = [name for topic in topics for name in topic.get("interfaces", [])]
    if len(assigned) != len(set(assigned)):
        raise ValueError("An interface may belong to only one authoritative topic")
    expected = set(manual["interfaces"])
    if set(assigned) != expected:
        raise ValueError(f"Interface mapping mismatch: {expected ^ set(assigned)}")
    for topic in topics:
        for link in topic.get("related", []):
            if link not in ids:
                raise ValueError(f"Unknown related topic: {link}")
        if topic.get("screenshots"):
            raise ValueError("Place figures in their explanatory section, not in a chapter gallery")
        for figure in figures(topic):
            shot = figure["capture"]
            evidence = captures.get(shot)
            if not evidence or evidence.get("provenance") not in ("live-workflow", "live-interface"):
                raise ValueError(f"Missing live interface capture record: {shot}")
            if evidence.get("synthetic") or not evidence.get("windowTitle"):
                raise ValueError(f"Invented fixtures cannot illustrate user workflows: {shot}")
            if figure.get("layout", "wide") not in ("side", "wide", "compact"):
                raise ValueError("Unknown figure layout")
            if not figure.get("caption"):
                raise ValueError("Every figure needs a task-specific caption")
            crop = figure.get("crop")
            if crop is not None:
                if (not isinstance(crop, list) or len(crop) != 4
                        or any(not isinstance(value, int) or isinstance(value, bool) for value in crop)):
                    raise ValueError("Figure crop must contain four integer pixel coordinates")
                x, y, width, height = crop
                if (x < 0 or y < 0 or width <= 0 or height <= 0
                        or x + width > evidence["width"] or y + height > evidence["height"]):
                    raise ValueError("Figure crop must stay inside the original capture")
            for callout in figure.get("callouts", []):
                rect = callout.get("rect", [])
                if len(rect) != 4 or any(not isinstance(v, (int, float)) for v in rect):
                    raise ValueError("Callout must have four percentage coordinates")
                x, y, w, h = rect
                if x < 0 or y < 0 or w <= 0 or h <= 0 or x + w > 100 or y + h > 100:
                    raise ValueError("Callout lies outside the actual capture")
            if not re.fullmatch(r"[A-Za-z0-9_-]+\.png", evidence["file"]):
                raise ValueError("Capture must name a local PNG file")
            path = source / "screenshots" / evidence["file"]
            data = path.read_bytes()
            if not data.startswith(b"\x89PNG\r\n\x1a\n"):
                raise ValueError(f"Not a PNG capture: {path}")
            if hashlib.sha256(data).hexdigest() != evidence["sha256"]:
                raise ValueError(f"Capture hash mismatch: {path}")


def figures(topic: dict):
    """Yield illustrations at their point of explanation, never as a trailing gallery."""
    return (section["figure"] for section in topic.get("sections", []) if section.get("figure"))


def render_figure(figure: dict, captures: dict) -> str:
    """Annotate an untouched live PNG with numbered HTML overlays and accessible explanations."""
    esc = html.escape
    shot = captures[figure["capture"]]
    crop = figure.get("crop")
    native_width = crop[2] if crop else int(shot.get("width", 900))
    width = min(int(figure.get("width", native_width)), native_width)
    source = "screenshots/" + esc(shot["file"])
    parts = [f"<div class='figure {esc(figure.get('layout', 'wide'))}' style='max-width:{width}px'>",
             f"<a class='picture' href='{source}' title='Afficher la capture à sa taille originale'>",
             ]
    if crop:
        x, y, crop_width, crop_height = crop
        parts.append(f"<span style='display:block;position:relative;overflow:hidden;height:0;padding-top:{100 * crop_height / crop_width:.6f}%'>"
                     f"<img src='{source}' alt='{esc(figure['caption'])}' style='position:absolute;max-width:none;"
                     f"width:{100 * shot['width'] / crop_width:.6f}%;left:{-100 * x / crop_width:.6f}%;top:{-100 * y / crop_height:.6f}%'>")
    else:
        parts.append(f"<img src='{source}' alt='{esc(figure['caption'])}'>")
    for i, callout in enumerate(figure.get("callouts", []), 1):
        x, y, w, h = callout["rect"]
        parts.append(f"<span class='mark' style='left:{x}%;top:{y}%;width:{w}%;height:{h}%'>"
                     f"<span class='number'>{i}</span></span>")
    if crop:
        parts.append("</span>")
    parts.append(f"</a><p class='caption'>{esc(figure['caption'])}</p>")
    if figure.get("callouts"):
        parts.append("<ol class='callouts'>" + "".join(
            f"<li>{esc(item['text'])}</li>" for item in figure["callouts"]) + "</ol>")
    parts.append("</div>")
    return "\n".join(parts)


def page(topic: dict, titles: dict, captures: dict) -> str:
    """Render one task-oriented Trident-compatible page without remote dependencies."""
    esc = html.escape
    parts = ["<!DOCTYPE html><html lang='fr'><head><meta charset='utf-8'>",
             "<meta http-equiv='X-UA-Compatible' content='IE=edge'>",
             "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>",
             f"<title>{esc(topic['title'])} — VBAi</title>",
             "<link rel='stylesheet' href='manual.css'></head><body>",
             "<div class='masthead'><strong>VBAi · Guide utilisateur</strong><br>",
             "<a href='start.html'>Prise en main</a> · <a href='interfaces.html'>Retrouver une commande</a></div><div class='content'>",
             f"<h1>{esc(topic['title'])}</h1><p class='lead'>{esc(topic['intro'])}</p>"]
    for section in topic.get("sections", []):
        parts.append(f"<div class='section'><h2>{esc(section['title'])}</h2>")
        figure = section.get("figure")
        if figure and figure.get("layout") == "side":
            parts.append(render_figure(figure, captures))
        for paragraph in section.get("paragraphs", []):
            parts.append(f"<p>{esc(paragraph)}</p>")
        for key, tag in (("steps", "ol"), ("items", "ul")):
            if section.get(key):
                parts.append(f"<{tag}>" + "".join(f"<li>{esc(item)}</li>" for item in section[key]) + f"</{tag}>")
        if section.get("code"):
            parts.append(f"<pre><code>{esc(section['code'])}</code></pre>")
        if section.get("table"):
            rows = section["table"]
            parts.append("<table><thead><tr>" + "".join(f"<th>{esc(v)}</th>" for v in rows[0]) + "</tr></thead><tbody>")
            parts.extend("<tr>" + "".join(f"<td>{esc(v)}</td>" for v in row) + "</tr>" for row in rows[1:])
            parts.append("</tbody></table>")
        if section.get("note"):
            parts.append(f"<div class='note'>{esc(section['note'])}</div>")
        if figure and figure.get("layout") != "side":
            parts.append(render_figure(figure, captures))
        parts.append("</div>")
    if topic.get("related"):
        parts.append("<div class='related'><h2>Pour poursuivre</h2><ul>" + "".join(
            f"<li><a href='{esc(key)}.html'>{esc(titles[key])}</a></li>" for key in topic["related"]) + "</ul></div>")
    parts.append("<p class='footer'>VBAi · Guide français · Cliquez sur une illustration pour la lire à sa taille originale.</p></div></body></html>")
    return "\n".join(parts)


def sitemap(entries: list[tuple[str, str]]) -> str:
    """Emit the legacy HTML Help sitemap used for both contents and index."""
    esc = html.escape
    return '<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN"><HTML><HEAD><meta charset="windows-1252"></HEAD><BODY><UL>\n' + "\n".join(
        f'<LI><OBJECT type="text/sitemap"><param name="Name" value="{esc(title)}"><param name="Local" value="{esc(target)}"></OBJECT>'
        for title, target in entries) + "\n</UL></BODY></HTML>"


def build(source: Path, output: Path, compiler: Path | None) -> Path:
    """Write reviewable HTML and optionally compile it, verifying the CHM header."""
    manual = read_json(source / "manual.json")
    manifest = read_json(source / "screenshots" / "manifest.json")
    captures = {item["id"]: item for item in manifest["captures"]}
    validate(manual, source, captures)
    output.mkdir(parents=True, exist_ok=True)
    titles = {item["id"]: item["title"] for item in manual["topics"]}
    for topic in manual["topics"]:
        (output / f"{topic['id']}.html").write_text(page(topic, titles, captures), encoding="utf-8")
    shots = output / "screenshots"
    shots.mkdir(exist_ok=True)
    for topic in manual["topics"]:
        for key in (figure["capture"] for figure in figures(topic)):
            filename = captures[key]["file"]
            shutil.copyfile(source / "screenshots" / filename, shots / filename)
    (output / "manual.css").write_text(STYLE, encoding="utf-8")
    contents = [(item["title"], f"{item['id']}.html") for item in manual["topics"]]
    index = sorted([(key, f"{item['id']}.html") for item in manual["topics"]
                    for key in item.get("keywords", [item["title"]])], key=lambda pair: pair[0].casefold())
    (output / "manual.hhc").write_text(sitemap(contents), encoding="cp1252", errors="xmlcharrefreplace")
    (output / "manual.hhk").write_text(sitemap(index), encoding="cp1252", errors="xmlcharrefreplace")
    files = sorted({"manual.css", *(f"{topic['id']}.html" for topic in manual["topics"]),
                    *("screenshots/" + captures[key]["file"] for topic in manual["topics"]
                      for key in (figure["capture"] for figure in figures(topic)))})
    project = "\n".join(["[OPTIONS]", "Compatibility=1.1 or later", "Compiled file=VBAi.fr-FR.chm",
        "Contents file=manual.hhc", "Index file=manual.hhk", "Default topic=start.html",
        "Default Window=main", "Display compile progress=No", "Full-text search=Yes",
        "Language=0x40c French (France)", "Title=VBAi - Guide utilisateur", "Error log file=compiler.log",
        "[WINDOWS]", 'main="VBAi - Guide utilisateur","manual.hhc","manual.hhk","start.html","start.html",,,,,0x63520,,0x304e,[60,40,1200,860],,,,,,,0',
        "[FILES]", *files, ""])
    (output / "manual.hhp").write_text(project, encoding="cp1252")
    if compiler:
        chm = output / "VBAi.fr-FR.chm"
        # An old archive cannot serve as evidence for a failed new compilation.
        chm.unlink(missing_ok=True)
        result = subprocess.run([str(compiler.resolve()), "manual.hhp"], cwd=output,
                                stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=120)
        (output / "compiler-console.txt").write_bytes(result.stdout)
        # hhc.exe normally returns 1 on success. The actual archive is the oracle.
        # It can also return 1 and emit a partial archive after an indexing error.
        compiler_error = re.search(rb"HHC\d+:\s*(?:Error|Erreur):", result.stdout, re.IGNORECASE)
        if compiler_error or result.returncode not in (0, 1) or not chm.exists() or chm.read_bytes()[:4] != b"ITSF":
            chm.unlink(missing_ok=True)
            raise RuntimeError(f"HTML Help compilation failed (exit {result.returncode}); inspect compiler logs")
        return chm
    return output / "start.html"


def main() -> None:
    """Keep source/output/compiler choices explicit and avoid installation side effects."""
    parser = argparse.ArgumentParser(description=__doc__)
    repository = Path(__file__).resolve().parents[2]
    parser.add_argument("--source", type=Path, default=repository / "docs/help/fr-FR")
    parser.add_argument("--output", type=Path, default=repository / "artifacts/help/fr-FR")
    parser.add_argument("--compiler", type=Path)
    args = parser.parse_args()
    print(build(args.source.resolve(), args.output.resolve(), args.compiler))


if __name__ == "__main__":
    main()
