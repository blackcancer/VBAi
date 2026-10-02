"""Compare complete Q026 option receipts offline; never connect to a host."""

import argparse
import json
from pathlib import Path
import re


class IncompleteEvidence(ValueError):
    """A receipt cannot establish a complete preference observation."""


def snapshot(receipt):
    """Find one unambiguous complete snapshot, retaining every hashed tab field."""
    found = []

    def visit(value, depth=0):
        if depth > 64:
            raise IncompleteEvidence("Receipt nesting exceeds the bound.")
        if isinstance(value, dict):
            if value.get("Truncated") is True:
                raise IncompleteEvidence("Truncated receipts cannot establish complete state.")
            if "Tabs" in value and "OptionsVersion" in value:
                found.append(value)
                return
            # The fixture's bounded command envelope stores the actual reply as JSON.
            if isinstance(value.get("Json"), str):
                try:
                    visit(json.loads(value["Json"]), depth + 1)
                except json.JSONDecodeError as error:
                    raise IncompleteEvidence("Embedded reply JSON is incomplete.") from error
            for key, child in value.items():
                if key != "Json":
                    visit(child, depth + 1)
        elif isinstance(value, list):
            for child in value:
                visit(child, depth + 1)

    visit(receipt)
    if len(found) != 1:
        raise IncompleteEvidence("Select a receipt containing exactly one options snapshot.")
    state = found[0]
    if not isinstance(state["OptionsVersion"], str) or not re.fullmatch(r"[a-fA-F0-9]{64}", state["OptionsVersion"]):
        raise IncompleteEvidence("A complete options SHA-256 revision is required.")
    if state.get("DialogClosed") is not True:
        raise IncompleteEvidence("The receipt does not verify native dialog closure.")
    tabs = state["Tabs"]
    if not isinstance(tabs, list) or not tabs or state.get("Count") != len(tabs):
        raise IncompleteEvidence("The complete tab count is absent or inconsistent.")
    fields = {"Name", "Type", "Value", "Error", "Choices", "NativeChoices", "SelectedIndex"}

    def controls(items):
        if not isinstance(items, list):
            raise IncompleteEvidence("A control collection is missing.")
        for control in items:
            if not isinstance(control, dict) or not fields.issubset(control):
                raise IncompleteEvidence("Hashed control fields are missing.")
            if control["Error"] is not None or not isinstance(control["Choices"], list) or not isinstance(control["NativeChoices"], list):
                raise IncompleteEvidence("A control was unreadable or its catalogue is incomplete.")

    for tab in tabs:
        if not isinstance(tab, dict) or not isinstance(tab.get("Tab"), str):
            raise IncompleteEvidence("A tab identity is missing.")
        controls(tab.get("Controls"))
        if tab.get("Count") != len(tab["Controls"]):
            raise IncompleteEvidence("The tab control count is inconsistent.")
        if not isinstance(tab.get("FormatCategories"), list):
            raise IncompleteEvidence("The complete category collection is missing.")
        for category in tab["FormatCategories"]:
            if not isinstance(category, dict) or not isinstance(category.get("Category"), str):
                raise IncompleteEvidence("A category identity is missing.")
            controls(category.get("Palettes"))
    return state


def changes(before, after, pointer="/Tabs"):
    """Compare types and array order as well as values; do not normalize evidence."""
    if type(before) is not type(after):
        return [{"Path": pointer, "Before": before, "After": after}]
    if isinstance(before, dict):
        result = []
        for key in sorted(before.keys() | after.keys()):
            path = pointer + "/" + key.replace("~", "~0").replace("/", "~1")
            if key not in before or key not in after:
                result.append({"Path": path, "BeforePresent": key in before, "AfterPresent": key in after,
                               "Before": before.get(key), "After": after.get(key)})
            else:
                result.extend(changes(before[key], after[key], path))
        return result
    if isinstance(before, list):
        result = []
        if len(before) != len(after):
            result.append({"Path": pointer + "/length", "Before": len(before), "After": len(after)})
        for index, (left, right) in enumerate(zip(before, after)):
            result.extend(changes(left, right, pointer + "/" + str(index)))
        return result
    return [] if before == after else [{"Path": pointer, "Before": before, "After": after}]


def compare(before_receipt, after_receipt):
    before, after = snapshot(before_receipt), snapshot(after_receipt)
    difference = changes(before["Tabs"], after["Tabs"])
    revision_equal = before["OptionsVersion"].lower() == after["OptionsVersion"].lower()
    return {
        "Scope": "Offline recorded snapshots only; no host execution, restoration or crash causality proof",
        "BeforeVersion": before["OptionsVersion"], "AfterVersion": after["OptionsVersion"],
        "RevisionEqual": revision_equal, "TabsEqual": not difference,
        "Classification": ("RECORDED_STATE_EQUAL" if revision_equal and not difference else
                           "REVISION_CHANGED_WITHOUT_STRUCTURAL_CHANGE" if not difference else
                           "STRUCTURE_CHANGED_WITH_EQUAL_REVISION" if revision_equal else "RECORDED_STATE_CHANGED"),
        "Changes": difference,
    }


def load(path):
    path = Path(path)
    if path.stat().st_size > 16 * 1024 * 1024:
        raise IncompleteEvidence("Receipt exceeds the file size bound.")
    return json.loads(path.read_text(encoding="utf-8-sig"))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("before", type=Path)
    parser.add_argument("after", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    if args.output.exists():
        parser.error("Use a fresh output path; original evidence must not be overwritten.")
    report = compare(load(args.before), load(args.after))
    report.update(BeforeReceipt=str(args.before.resolve()), AfterReceipt=str(args.after.resolve()))
    with args.output.open("x", encoding="utf-8") as stream:
        json.dump(report, stream, indent=2, ensure_ascii=False)
        stream.write("\n")
    print(report["Classification"])


if __name__ == "__main__":
    main()
