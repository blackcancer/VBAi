"""Require complete managed coverage of the VBA test subsystem from one Coverlet run.

Cobertura is reconciled with its companion JSON, which retains branches with no
emitted sequence-point line. No source, generated class, or IL branch is excluded.
"""
import argparse
from collections import Counter, defaultdict
import hashlib
import json
from pathlib import Path
import re
import sys
import xml.etree.ElementTree as ET


INTERFACE_ONLY = "src/VBAi/Testing/VbaTestExplorerService.cs"
LLM_SOURCE = "src/VBAi/Llm/Chat/LlmVbeTools.Testing.cs"
CONDITION = re.compile(r"^(\d+(?:\.\d+)?)% \((\d+)/(\d+)\)$")


def unique_object(pairs):
    result = {}
    for name, value in pairs:
        if name in result:
            raise ValueError("Duplicate JSON object key: " + name)
        result[name] = value
    return result


def integer(value, label, minimum=0):
    if isinstance(value, bool) or not isinstance(value, (int, str)):
        raise ValueError(label + " must be an integer.")
    result = int(value)
    if str(result) != str(value) or result < minimum:
        raise ValueError(label + " is outside its valid integer range.")
    return result


def source_path(value):
    """Normalize Coverlet absolute, relative, deterministic and Windows paths."""
    normalized = value.replace("\\", "/")
    if ".." in normalized.split("/"):
        raise ValueError("Parent traversal in a collector source path: " + value)
    marker = "src/VBAi/"
    position = normalized.rfind(marker)
    if position >= 0:
        return normalized[position:]
    return marker + normalized.lstrip("/")


def interface_only(path):
    # The sole exception remains valid only while every body is a namespace or
    # interface declaration; a future concrete implementation must be measured.
    source = re.sub(r"/\*.*?\*/|//[^\n]*", "", path.read_text(encoding="utf-8-sig"), flags=re.S)
    if not re.search(r"\binterface\s+\w+", source):
        return False
    boundary = 0
    for match in re.finditer(r"[{};]", source):
        if match.group() == "{":
            header = source[boundary:match.start()].strip()
            if not re.fullmatch(r"(?:namespace\s+[\w.]+|(?:(?:public|internal)\s+)?interface\s+\w+(?:\s*:\s*[\w.,\s]+)?)", header):
                return False
        boundary = match.end()
    return True


def counts(covered, total):
    return {"covered": covered, "total": total,
            "percent": round(100 * covered / total, 6) if total else None}


def branch_pair(line):
    flag = line.get("branch", "false").lower()
    if flag not in ("true", "false"):
        raise ValueError("Invalid Cobertura branch flag.")
    if flag == "false":
        if "condition-coverage" in line.attrib:
            raise ValueError("A non-branch line has a condition counter.")
        return (0, 0)
    match = CONDITION.fullmatch(line.get("condition-coverage", ""))
    if not match:
        raise ValueError("Missing exact Cobertura condition numerator/denominator.")
    covered, total = int(match[2]), int(match[3])
    if total == 0 or covered > total or abs(float(match[1]) - 100 * covered / total) > 0.011:
        raise ValueError("Inconsistent Cobertura condition counter.")
    return covered, total


def xml_rows(lines):
    return Counter((integer(line.get("number"), "Source line", 1),
                    integer(line.get("hits"), "Line hits"), *branch_pair(line))
                   for line in lines)


def compare_rate(element, metric, covered, total, context, errors):
    if not total:
        return
    try:
        actual = float(element.attrib[metric + "-rate"])
        if not 0 <= actual <= 1 or abs(actual - covered / total) > 0.00011:
            errors.append(context + ": " + metric + " rate disagrees with raw JSON counters.")
    except (KeyError, ValueError):
        errors.append(context + ": invalid " + metric + " rate.")


def method_name(signature):
    tail = signature.rsplit(":", 1)[-1]
    name, separator, arguments = tail.partition("(")
    if not separator or not arguments.endswith(")"):
        raise ValueError("Invalid Coverlet method signature: " + signature)
    return name, "(" + arguments


def check(cobertura, repository, companion=None):
    repository, cobertura = Path(repository).resolve(), Path(cobertura).resolve()
    companion = Path(companion).resolve() if companion else cobertura.with_name("coverage.json")
    result = {"v": 1, "passed": False, "scope": "Managed VBA test subsystem",
              "reports": {"cobertura": str(cobertura), "coverletJson": str(companion)},
              "files": [], "total": {}, "errors": []}
    errors = result["errors"]
    try:
        testing = repository / "src/VBAi/Testing"
        if not testing.is_dir() or not (repository / LLM_SOURCE).is_file():
            raise ValueError("The complete expected production scope is absent from the repository.")
        expected = sorted(path.relative_to(repository).as_posix() for path in testing.glob("*.cs")) + [LLM_SOURCE]
        root = ET.parse(cobertura).getroot()
        if root.tag != "coverage":
            raise ValueError("Expected a Cobertura coverage document.")
        raw = json.loads(companion.read_text(encoding="utf-8-sig"), object_pairs_hook=unique_object)
        if not isinstance(raw, dict) or not raw:
            raise ValueError("Expected a nonempty Coverlet JSON document from the same collection.")
        result["reports"]["coberturaSha256"] = hashlib.sha256(cobertura.read_bytes()).hexdigest()
        result["reports"]["coverletJsonSha256"] = hashlib.sha256(companion.read_bytes()).hexdigest()
        files, classes = {}, {}
        global_lines, global_branches = [], []
        production_found = False
        for assembly, documents in raw.items():
            production = Path(assembly.replace("\\", "/")).name == "VBAi.dll"
            production_found |= production
            for filename, types in documents.items():
                normalized = source_path(filename) if production else filename
                if production and normalized in files:
                    raise ValueError("Duplicate canonical production document: " + normalized)
                details = {"path": normalized, "classes": 0, "methods": 0, "lineRecords": [], "branchRecords": []}
                if production:
                    files[normalized] = details
                for name, methods in types.items():
                    details["classes"] += 1
                    class_lines, class_branches, visible_rows, method_rows = [], [], Counter(), {}
                    for signature, method in methods.items():
                        details["methods"] += 1
                        lines = {integer(number, "Source line", 1): integer(hits, "Line hits") for number, hits in method["Lines"].items()}
                        branches = []
                        for record, branch in enumerate(method["Branches"]):
                            point = {key: integer(branch[key], "Branch " + key) for key in ("Offset", "EndOffset", "Path", "Ordinal", "Hits")}
                            # Unmapped IL branches may have a nonpositive source line.
                            point["Line"] = int(branch["Line"])
                            point.update({"class": name, "method": signature, "record": record, "mapped": point["Line"] in lines})
                            branches.append(point)
                        records = [{"class": name, "method": signature, "line": number, "hits": hits} for number, hits in lines.items()]
                        rows = Counter()
                        for number, hits in lines.items():
                            associated = [branch for branch in branches if branch["Line"] == number]
                            rows[(number, hits, sum(branch["Hits"] > 0 for branch in associated), len(associated))] += 1
                        visible_rows.update(rows)
                        class_lines.extend(records); class_branches.extend(branches)
                        key = method_name(signature)
                        if key in method_rows:
                            raise ValueError("Ambiguous method identity in " + name)
                        method_rows[key] = (rows, records, branches)
                    details["lineRecords"].extend(class_lines); details["branchRecords"].extend(class_branches)
                    global_lines.extend(class_lines); global_branches.extend(class_branches)
                    if production:
                        classes[(normalized, name)] = (visible_rows, method_rows, class_lines, class_branches)
        if not production_found:
            raise ValueError("The VBAi.dll assembly is absent from the raw collector report.")
        for attribute, actual in (("lines-valid", len(global_lines)), ("lines-covered", sum(line["hits"] > 0 for line in global_lines)),
                                  ("branches-valid", len(global_branches)), ("branches-covered", sum(branch["Hits"] > 0 for branch in global_branches))):
            if integer(root.get(attribute), attribute) != actual:
                errors.append("Cobertura root " + attribute + " disagrees with the companion JSON; use files from one collection.")
        xml_classes, xml_packages = set(), set()
        for package in root.findall("./packages/package"):
            package_name = package.get("name")
            if package_name in xml_packages:
                raise ValueError("Duplicate Cobertura assembly package: " + str(package_name))
            xml_packages.add(package_name)
            if package_name != "VBAi":
                continue
            for element in package.findall("./classes/class"):
                filename, name = source_path(element.attrib["filename"]), element.attrib["name"]
                key = (filename, name)
                if key in xml_classes:
                    raise ValueError("Duplicate Cobertura class/document: " + str(key))
                xml_classes.add(key)
                if key not in classes:
                    errors.append("Cobertura class is absent from companion JSON: " + str(key)); continue
                expected_rows, methods, lines, branches = classes[key]
                if xml_rows(element.findall("./lines/line")) != expected_rows:
                    errors.append("Cobertura class line/condition rows disagree with JSON: " + str(key))
                compare_rate(element, "line", sum(line["hits"] > 0 for line in lines), len(lines), str(key), errors)
                compare_rate(element, "branch", sum(branch["Hits"] > 0 for branch in branches), len(branches), str(key), errors)
                seen_methods = set()
                for xml_method in element.findall("./methods/method"):
                    method_key = (xml_method.attrib["name"], xml_method.attrib["signature"])
                    if method_key in seen_methods or method_key not in methods:
                        raise ValueError("Duplicate or unknown Cobertura method: " + str(method_key))
                    seen_methods.add(method_key)
                    rows, method_lines, method_branches = methods[method_key]
                    if xml_rows(xml_method.findall("./lines/line")) != rows:
                        errors.append("Cobertura method rows disagree with JSON: " + str(method_key))
                    compare_rate(xml_method, "line", sum(line["hits"] > 0 for line in method_lines), len(method_lines), str(method_key), errors)
                    compare_rate(xml_method, "branch", sum(branch["Hits"] > 0 for branch in method_branches), len(method_branches), str(method_key), errors)
                if seen_methods != {key for key, (_, lines, _) in methods.items() if lines}:
                    errors.append("Cobertura method inventory disagrees with JSON: " + str(key))
        if "VBAi" not in xml_packages or xml_classes != set(classes):
            errors.append("Cobertura production class inventory disagrees with companion JSON.")
        collected_scope = {name for name in files if name.startswith("src/VBAi/Testing/") or name == LLM_SOURCE}
        for obsolete in sorted(collected_scope - set(expected)):
            errors.append("Collected subsystem source is absent from the expected repository scope: " + obsolete)
        for filename in expected:
            path = repository / filename
            details = files.get(filename)
            exemption = filename == INTERFACE_ONLY and interface_only(path)
            if details is None or not details["lineRecords"]:
                if not exemption:
                    errors.append("Expected production file has no collector sequence points: " + filename)
            lines = details["lineRecords"] if details else []
            branches = details["branchRecords"] if details else []
            uncovered_lines = [line for line in lines if line["hits"] == 0]
            uncovered_branches = [branch for branch in branches if branch["Hits"] == 0]
            if uncovered_lines or uncovered_branches:
                errors.append("Coverage below 100%: " + filename)
            unmapped = [branch for branch in branches if not branch["mapped"]]
            result["files"].append({"path": filename, "sourceSha256": hashlib.sha256(path.read_bytes()).hexdigest(),
                                    "classes": details["classes"] if details else 0, "methods": details["methods"] if details else 0,
                                    "interfaceOnly": exemption, "lines": counts(len(lines) - len(uncovered_lines), len(lines)),
                                    "branches": counts(len(branches) - len(uncovered_branches), len(branches)),
                                    "unmappedBranches": counts(sum(branch["Hits"] > 0 for branch in unmapped), len(unmapped)),
                                    "uncoveredLines": uncovered_lines, "uncoveredBranches": uncovered_branches})
        result["total"] = {metric: counts(sum(item[metric]["covered"] for item in result["files"]), sum(item[metric]["total"] for item in result["files"])) for metric in ("lines", "branches", "unmappedBranches")}
        result["passed"] = not errors and result["total"]["lines"]["total"] > 0
    except (OSError, ValueError, TypeError, KeyError, AttributeError, ET.ParseError) as error:
        errors.append(str(error))
    return result


def human(result):
    output = ["Managed VBA testing coverage: " + ("PASS" if result["passed"] else "FAIL")]
    for item in result["files"]:
        line, branch = item["lines"], item["branches"]
        output.append(f"{item['path']}: lines {line['covered']}/{line['total']}; branches {branch['covered']}/{branch['total']}" + (" (interface declarations only)" if item["interfaceOnly"] else ""))
    if result["total"]:
        line, branch, unmapped = (result["total"][name] for name in ("lines", "branches", "unmappedBranches"))
        output.append(f"Total: lines {line['covered']}/{line['total']}; all IL branches {branch['covered']}/{branch['total']}; unmapped branches {unmapped['covered']}/{unmapped['total']}")
    output.extend("ERROR: " + error for error in result["errors"])
    return "\n".join(output)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("cobertura", type=Path)
    parser.add_argument("--coverlet-json", type=Path, help="Companion report; defaults to coverage.json beside Cobertura.")
    parser.add_argument("--repo-root", type=Path, default=Path(__file__).resolve().parents[2])
    parser.add_argument("--json", type=Path, help="Save the complete machine-readable gate result.")
    arguments = parser.parse_args(argv)
    result = check(arguments.cobertura, arguments.repo_root, arguments.coverlet_json)
    if arguments.json:
        arguments.json.parent.mkdir(parents=True, exist_ok=True)
        arguments.json.write_text(json.dumps(result, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(human(result))
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    sys.exit(main())