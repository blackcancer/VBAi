"""Contract tests for the managed coverage gate, using disposable synthetic reports."""
from copy import deepcopy
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET

from check_managed_coverage import check, human


class ManagedCoverageChecks(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.source = self.root / "src/VBAi/Testing/Core.cs"
        self.source.parent.mkdir(parents=True)
        self.source.write_text("namespace VBAi { class Core { void Execute() {} } }", encoding="utf-8")
        interface = self.source.with_name("VbaTestExplorerService.cs")
        interface.write_text("namespace VBAi { internal interface ITests { void Run(); } }", encoding="utf-8")
        self.llm = self.root / "src/VBAi/Llm/Chat/LlmVbeTools.Testing.cs"
        self.llm.parent.mkdir(parents=True)
        self.llm.write_text("namespace VBAi { class Tools { void Run() {} } }", encoding="utf-8")
        self.raw = {"VBAi.dll": {
            self.source.as_posix(): {"VBAi.Core": {"System.Void VBAi.Core::Execute()": {
                "Lines": {"1": 1}, "Branches": [self.branch(0), self.branch(1)]}}},
            self.llm.as_posix(): {"VBAi.Tools": {"System.Void VBAi.Tools::Run()": {"Lines": {"1": 1}, "Branches": []}}}}}
        self.xml = self.root / "coverage.cobertura.xml"
        self.companion = self.root / "coverage.json"

    @staticmethod
    def branch(path, line=1, hits=1, offset=5):
        return {"Line": line, "Offset": offset, "EndOffset": 10 + path, "Path": path, "Ordinal": path, "Hits": hits}

    @staticmethod
    def rate(covered, total):
        return str(covered / total if total else 1)

    def core(self):
        return self.raw["VBAi.dll"][self.source.as_posix()]["VBAi.Core"]["System.Void VBAi.Core::Execute()"]

    def write(self):
        # Independently emit the documented collector schema, retaining raw method
        # and class multiplicities rather than collapsing repeated source lines.
        all_methods = [method for documents in self.raw.values() for classes in documents.values()
                       for methods in classes.values() for method in methods.values()]
        line_total = sum(len(method["Lines"]) for method in all_methods)
        line_covered = sum(hit > 0 for method in all_methods for hit in method["Lines"].values())
        branches = [branch for method in all_methods for branch in method["Branches"]]
        root = ET.Element("coverage", {"lines-valid": str(line_total), "lines-covered": str(line_covered),
                                      "branches-valid": str(len(branches)), "branches-covered": str(sum(branch["Hits"] > 0 for branch in branches))})
        packages = ET.SubElement(root, "packages")
        for assembly, documents in self.raw.items():
            package = ET.SubElement(packages, "package", {"name": assembly.removesuffix(".dll")})
            classes_node = ET.SubElement(package, "classes")
            for filename, classes in documents.items():
                for name, methods in classes.items():
                    lines = [hit for method in methods.values() for hit in method["Lines"].values()]
                    branches = [branch for method in methods.values() for branch in method["Branches"]]
                    element = ET.SubElement(classes_node, "class", {"name": name, "filename": filename,
                        "line-rate": self.rate(sum(hit > 0 for hit in lines), len(lines)),
                        "branch-rate": self.rate(sum(branch["Hits"] > 0 for branch in branches), len(branches))})
                    method_nodes = ET.SubElement(element, "methods")
                    class_lines = ET.SubElement(element, "lines")
                    for signature, method in methods.items():
                        if not method["Lines"]:
                            continue
                        tail = signature.rsplit(":", 1)[-1]
                        name, _, parameters = tail.partition("(")
                        branch_list = method["Branches"]
                        method_node = ET.SubElement(method_nodes, "method", {"name": name, "signature": "(" + parameters,
                            "line-rate": self.rate(sum(hit > 0 for hit in method["Lines"].values()), len(method["Lines"])),
                            "branch-rate": self.rate(sum(branch["Hits"] > 0 for branch in branch_list), len(branch_list))})
                        method_lines = ET.SubElement(method_node, "lines")
                        for number, hits in method["Lines"].items():
                            associated = [branch for branch in branch_list if branch["Line"] == int(number)]
                            attributes = {"number": str(number), "hits": str(hits), "branch": str(bool(associated))}
                            if associated:
                                covered = sum(branch["Hits"] > 0 for branch in associated)
                                attributes["condition-coverage"] = f"{100 * covered / len(associated):.2f}% ({covered}/{len(associated)})"
                            ET.SubElement(method_lines, "line", attributes)
                            ET.SubElement(class_lines, "line", attributes)
        ET.ElementTree(root).write(self.xml, encoding="utf-8", xml_declaration=True)
        self.companion.write_text(json.dumps(self.raw), encoding="utf-8")
        return check(self.xml, self.root)

    def test_valid_scope_includes_interface_exception_and_machine_and_human_totals(self):
        result = self.write()
        self.assertTrue(result["passed"], result["errors"])
        self.assertEqual(result["total"]["lines"], {"covered": 2, "total": 2, "percent": 100.0})
        self.assertEqual(result["total"]["branches"]["total"], 2)
        interface = next(item for item in result["files"] if item["interfaceOnly"])
        self.assertEqual(interface["lines"]["total"], 0)
        self.assertIn("Managed VBA testing coverage: PASS", human(result))
        self.assertIn("all IL branches 2/2", human(result))

    def test_zero_hit_line_fails_even_when_all_its_branches_are_covered(self):
        self.core()["Lines"]["1"] = 0
        result = self.write()
        self.assertFalse(result["passed"])
        core = next(item for item in result["files"] if item["path"].endswith("Core.cs"))
        self.assertEqual(core["lines"]["covered"], 0)
        self.assertEqual(core["branches"]["percent"], 100.0)
        self.assertEqual(core["uncoveredLines"][0]["line"], 1)

    def test_half_covered_branch_fails_and_retains_exact_offset_and_path(self):
        self.core()["Branches"][1]["Hits"] = 0
        result = self.write()
        self.assertFalse(result["passed"])
        core = next(item for item in result["files"] if item["path"].endswith("Core.cs"))
        self.assertEqual(core["branches"]["percent"], 50.0)
        self.assertEqual(core["uncoveredBranches"][0]["Offset"], 5)
        self.assertEqual(core["uncoveredBranches"][0]["Path"], 1)

    def test_expected_source_missing_from_collector_cannot_reduce_the_denominator(self):
        self.source.with_name("NewRunner.cs").write_text("class NewRunner {}", encoding="utf-8")
        result = self.write()
        self.assertFalse(result["passed"])
        self.assertTrue(any("NewRunner.cs" in error and "no collector" in error for error in result["errors"]))

    def test_stale_collector_document_for_removed_production_code_is_rejected(self):
        old = self.source.with_name("RemovedRunner.cs").as_posix()
        self.raw["VBAi.dll"][old] = {"VBAi.RemovedRunner": {"System.Void VBAi.RemovedRunner::Run()": {"Lines": {"1": 1}, "Branches": []}}}
        result = self.write()
        self.assertFalse(result["passed"])
        self.assertTrue(any("RemovedRunner.cs" in error and "absent" in error for error in result["errors"]))

    def test_missing_llm_commands_are_part_of_the_required_scope(self):
        del self.raw["VBAi.dll"][self.llm.as_posix()]
        result = self.write()
        self.assertFalse(result["passed"])
        self.assertTrue(any("LlmVbeTools.Testing.cs" in error for error in result["errors"]))

    def test_interface_exception_stops_when_executable_code_is_added(self):
        self.source.with_name("VbaTestExplorerService.cs").write_text("namespace VBAi { class Service { void Run() {} } }", encoding="utf-8")
        result = self.write()
        self.assertFalse(result["passed"])
        self.assertTrue(any("VbaTestExplorerService.cs" in error for error in result["errors"]))

    def test_default_interface_method_body_is_not_exempt(self):
        self.source.with_name("VbaTestExplorerService.cs").write_text("namespace VBAi { interface IService { void Run() {} } }", encoding="utf-8")
        self.assertFalse(self.write()["passed"])

    def test_generated_classes_on_the_same_source_line_remain_separate(self):
        classes = self.raw["VBAi.dll"][self.source.as_posix()]
        classes["VBAi.Core/<ExecuteAsync>d__1"] = {"System.Void VBAi.Core/<ExecuteAsync>d__1::MoveNext()": {
            "Lines": {"1": 0}, "Branches": [self.branch(0), self.branch(1)]}}
        result = self.write()
        self.assertFalse(result["passed"])
        core = next(item for item in result["files"] if item["path"].endswith("Core.cs"))
        self.assertEqual(core["lines"]["total"], 2)
        self.assertEqual(core["lines"]["covered"], 1)
        self.assertEqual(core["branches"]["total"], 4)
        self.assertIn("MoveNext", core["uncoveredLines"][0]["method"])
        classes["VBAi.Core/<ExecuteAsync>d__1"]["System.Void VBAi.Core/<ExecuteAsync>d__1::MoveNext()"]["Lines"]["1"] = 1
        self.assertTrue(self.write()["passed"])

    def test_unmapped_il_branch_is_counted_even_when_every_cobertura_line_branch_passes(self):
        self.core()["Branches"].extend([self.branch(0, line=0, offset=20), self.branch(1, line=0, hits=0, offset=20)])
        result = self.write()
        self.assertFalse(result["passed"])
        self.assertEqual(result["total"]["branches"], {"covered": 3, "total": 4, "percent": 75.0})
        self.assertEqual(result["total"]["unmappedBranches"], {"covered": 1, "total": 2, "percent": 50.0})

    def test_method_without_sequence_points_still_contributes_all_its_il_branches(self):
        self.raw["VBAi.dll"][self.source.as_posix()]["VBAi.Core"]["System.Void VBAi.Core::Hidden()"] = {
            "Lines": {}, "Branches": [self.branch(0, line=0, hits=0), self.branch(1, line=0)]}
        result = self.write()
        self.assertFalse(result["passed"])
        self.assertEqual(result["total"]["branches"]["total"], 4)

    def test_repeated_branch_offsets_are_retained_including_identical_raw_records(self):
        self.core()["Branches"].append(deepcopy(self.core()["Branches"][0]))
        result = self.write()
        self.assertTrue(result["passed"], result["errors"])
        self.assertEqual(result["total"]["branches"]["total"], 3)

    def test_multiple_methods_on_one_physical_line_are_not_deduplicated(self):
        self.raw["VBAi.dll"][self.source.as_posix()]["VBAi.Core"]["System.Void VBAi.Core::Second()"] = {
            "Lines": {"1": 1}, "Branches": [self.branch(0), self.branch(1)]}
        result = self.write()
        self.assertTrue(result["passed"], result["errors"])
        self.assertEqual(result["total"]["lines"]["total"], 3)
        self.assertEqual(result["total"]["branches"]["total"], 4)

    def test_companion_json_is_required_to_verify_branches_without_sequence_points(self):
        self.write(); self.companion.unlink()
        result = check(self.xml, self.root)
        self.assertFalse(result["passed"])
        self.assertTrue(any("coverage.json" in error for error in result["errors"]))

    def test_reports_from_different_collections_are_rejected(self):
        self.write()
        root = ET.parse(self.xml)
        root.getroot().set("branches-covered", "1")
        root.write(self.xml)
        result = check(self.xml, self.root)
        self.assertFalse(result["passed"])
        self.assertTrue(any("one collection" in error for error in result["errors"]))

    def test_duplicate_json_object_keys_cannot_silently_replace_a_collector_document(self):
        self.write()
        self.companion.write_text('{"VBAi.dll": {}, "VBAi.dll": {}}', encoding="utf-8")
        result = check(self.xml, self.root)
        self.assertFalse(result["passed"])
        self.assertIn("Duplicate JSON object key", result["errors"][0])

    def test_duplicate_cobertura_classes_are_rejected_instead_of_double_counted(self):
        self.write()
        tree = ET.parse(self.xml)
        classes = tree.getroot().find("./packages/package/classes")
        classes.append(deepcopy(classes[0]))
        tree.write(self.xml)
        result = check(self.xml, self.root)
        self.assertFalse(result["passed"])
        self.assertTrue(any("Duplicate Cobertura class" in error for error in result["errors"]))

    def test_wrong_class_rate_or_missing_method_cannot_hide_raw_counters(self):
        self.write()
        tree = ET.parse(self.xml)
        tree.getroot().find("./packages/package/classes/class").set("branch-rate", "0.5")
        tree.write(self.xml)
        self.assertFalse(check(self.xml, self.root)["passed"])
        self.write()
        tree = ET.parse(self.xml)
        methods = tree.getroot().find("./packages/package/classes/class/methods")
        methods.remove(methods[0]); tree.write(self.xml)
        self.assertFalse(check(self.xml, self.root)["passed"])

    def test_absolute_windows_json_paths_match_relative_windows_cobertura_paths(self):
        documents = self.raw["VBAi.dll"]
        classes = documents.pop(self.source.as_posix())
        documents[r"C:\other\checkout\src\VBAi\Testing\Core.cs"] = classes
        self.write()
        tree = ET.parse(self.xml)
        element = next(element for element in tree.getroot().findall(".//class") if element.get("name") == "VBAi.Core")
        element.set("filename", r"Testing\Core.cs"); tree.write(self.xml)
        result = check(self.xml, self.root)
        self.assertTrue(result["passed"], result["errors"])

    def test_invalid_source_line_zero_is_refused_instead_of_removed(self):
        self.core()["Lines"] = {"0": 1}
        self.assertFalse(self.write()["passed"])

    def test_invalid_condition_counter_is_refused(self):
        self.write()
        tree = ET.parse(self.xml)
        tree.getroot().find(".//class/lines/line").set("condition-coverage", "100% (1/2)")
        tree.write(self.xml)
        self.assertFalse(check(self.xml, self.root)["passed"])

    def test_cli_writes_complete_json_and_returns_nonzero_on_failure(self):
        self.core()["Branches"][1]["Hits"] = 0; self.write()
        output = self.root / "result/gate.json"
        process = subprocess.run([sys.executable, str(Path(__file__).with_name("check_managed_coverage.py")),
                                  str(self.xml), "--repo-root", str(self.root), "--json", str(output)],
                                 text=True, capture_output=True, check=False)
        self.assertEqual(process.returncode, 1, process.stderr)
        self.assertIn("FAIL", process.stdout)
        self.assertEqual(json.loads(output.read_text(encoding="utf-8"))["total"]["branches"]["percent"], 50.0)


if __name__ == "__main__":
    unittest.main()