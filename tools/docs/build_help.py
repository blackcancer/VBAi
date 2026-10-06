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


STYLE = """body{font:15px 'Segoe UI',Arial,sans-serif;color:#243247;background:#fff;
margin:0;line-height:1.65}.masthead{background:#14283f;color:white;padding:22px 34px}
.masthead a{color:#cce5ff}.content{max-width:1080px;margin:28px auto;padding:0 34px}
h1{font-size:30px;line-height:1.25;margin:12px 0 18px}h2{font-size:22px;
color:#174c7c;margin-top:32px;border-bottom:1px solid #dce4ed;padding-bottom:8px}
h3{font-size:17px;margin-top:24px}a{color:#075ea8}li{margin:8px 0}code{
font-family:Consolas,monospace;background:#f2f5f8;padding:2px 5px}table{
border-collapse:collapse;width:100%;margin:16px 0}th,td{padding:10px 14px;
border:1px solid #dce4ed;text-align:left;vertical-align:top}th{background:#eef4fa}
.lead{font-size:18px;color:#41556c}.note{background:#edf5fc;border-left:4px solid
#2474ad;padding:14px 18px;margin:22px 0}.shot{border:1px solid #dce4ed;
background:#f8fafc;padding:16px;margin:24px 0}.shot img{max-width:100%;height:auto}
.caption{font-size:13px;color:#536779;margin-top:10px}.footer{border-top:1px
solid #dce4ed;color:#536779;font-size:13px;margin-top:36px;padding:18px 0}
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
        for shot in topic.get("screenshots", []):
            evidence = captures.get(shot)
            if not evidence or evidence.get("synthetic") is not True:
                raise ValueError(f"Missing genuine synthetic UI capture record: {shot}")
            if not re.fullmatch(r"[A-Za-z0-9_-]+\.png", evidence["file"]):
                raise ValueError("Capture must name a local PNG file")
            path = source / "screenshots" / evidence["file"]
            data = path.read_bytes()
            if not data.startswith(b"\x89PNG\r\n\x1a\n"):
                raise ValueError(f"Not a PNG capture: {path}")
            if hashlib.sha256(data).hexdigest() != evidence["sha256"]:
                raise ValueError(f"Capture hash mismatch: {path}")


def page(topic: dict, titles: dict, captures: dict) -> str:
    """Render one Trident-compatible UTF-8 page without remote dependencies."""
    esc = html.escape
    parts = ["<!DOCTYPE html><html lang='fr'><head><meta charset='utf-8'>",
             "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>",
             f"<title>{esc(topic['title'])} — VBAi</title>",
             "<link rel='stylesheet' href='manual.css'></head><body>",
             "<div class='masthead'><strong>VBAi · Guide utilisateur</strong><br>",
             "<a href='start.html'>Accueil</a> · <a href='interfaces.html'>Interfaces</a></div><div class='content'>",
             f"<h1>{esc(topic['title'])}</h1><p class='lead'>{esc(topic['intro'])}</p>"]
    for section in topic.get("sections", []):
        parts.append(f"<h2>{esc(section['title'])}</h2>")
        for paragraph in section.get("paragraphs", []):
            parts.append(f"<p>{esc(paragraph)}</p>")
        if section.get("steps"):
            parts.append("<ol>" + "".join(f"<li>{esc(step)}</li>" for step in section["steps"]) + "</ol>")
        if section.get("items"):
            parts.append("<ul>" + "".join(f"<li>{esc(item)}</li>" for item in section["items"]) + "</ul>")
        if section.get("table"):
            rows = section["table"]
            parts.append("<table><thead><tr>" + "".join(f"<th>{esc(v)}</th>" for v in rows[0]) + "</tr></thead><tbody>")
            parts.extend("<tr>" + "".join(f"<td>{esc(v)}</td>" for v in row) + "</tr>" for row in rows[1:])
            parts.append("</tbody></table>")
        if section.get("note"):
            parts.append(f"<div class='note'>{esc(section['note'])}</div>")
    for key in topic.get("screenshots", []):
        shot = captures[key]
        parts.append(f"<div class='shot'><img src='screenshots/{esc(shot['file'])}' alt='{esc(shot['caption'])}'>"
                     f"<p class='caption'>{esc(shot['caption'])} · Données fictives, capture du véritable contrôle VBAi.</p></div>")
    if topic.get("related"):
        parts.append("<h2>Voir aussi</h2><ul>" + "".join(
            f"<li><a href='{esc(key)}.html'>{esc(titles[key])}</a></li>" for key in topic["related"]) + "</ul>")
    parts.append("<p class='footer'>VBAi · Première édition française. Les captures pédagogiques ne constituent pas une qualification des hôtes ou des providers.</p></div></body></html>")
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
        for key in topic.get("screenshots", []):
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
                      for key in topic.get("screenshots", []))})
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
