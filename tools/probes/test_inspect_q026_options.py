import copy
import json
import unittest

from inspect_q026_options import compare, IncompleteEvidence, snapshot


def receipt():
    control = {"Name": "Size", "Type": "ControlType.ComboBox", "Value": "12", "Error": None,
               "Choices": [], "NativeChoices": [], "SelectedIndex": -1}
    return {"OptionsVersion": "a" * 64, "DialogClosed": True, "Count": 1,
            "Tabs": [{"Tab": "Editor Format", "Count": 1, "Controls": [control], "FormatCategories": []}]}


class RecordedOptionsTests(unittest.TestCase):
    def test_complete_wrapped_and_encoded_receipts_compare_without_host_access(self):
        before = {"Evidence": {"Response": {"Data": receipt()}}}
        after = {"Truncated": False, "Json": json.dumps({"Ok": True, "Data": receipt()})}
        self.assertEqual("RECORDED_STATE_EQUAL", compare(before, after)["Classification"])

    def test_value_change_is_located_even_with_inconsistent_equal_revision(self):
        after = receipt()
        after["Tabs"][0]["Controls"][0]["Value"] = "14"
        result = compare(receipt(), after)
        self.assertEqual("STRUCTURE_CHANGED_WITH_EQUAL_REVISION", result["Classification"])
        self.assertEqual("/Tabs/0/Controls/0/Value", result["Changes"][0]["Path"])

    def test_only_revision_change_is_reported_without_claiming_causality(self):
        after = receipt()
        after["OptionsVersion"] = "b" * 64
        self.assertEqual("REVISION_CHANGED_WITHOUT_STRUCTURAL_CHANGE", compare(receipt(), after)["Classification"])

    def test_missing_malformed_truncated_open_and_partial_receipts_are_rejected(self):
        partial = receipt()
        del partial["Tabs"][0]["Controls"][0]["NativeChoices"]
        for value in [{"Json": "{", "Truncated": True}, {"OptionsVersion": "a" * 64},
                      dict(receipt(), OptionsVersion="prefix"), dict(receipt(), DialogClosed=False),
                      dict(receipt(), Count=2), partial, {"Json": "{"}]:
            with self.subTest(value=value), self.assertRaises(IncompleteEvidence):
                snapshot(value)

    def test_multiple_snapshots_cannot_be_silently_selected(self):
        with self.assertRaises(IncompleteEvidence):
            snapshot({"Before": receipt(), "After": receipt()})

    def test_palette_fields_and_numeric_types_are_not_ignored(self):
        before = receipt()
        palette = copy.deepcopy(before["Tabs"][0]["Controls"][0])
        before["Tabs"][0]["FormatCategories"] = [{"Category": "Keyword", "Palettes": [palette]}]
        after = copy.deepcopy(before)
        after["OptionsVersion"] = "b" * 64
        after["Tabs"][0]["FormatCategories"][0]["Palettes"][0]["SelectedIndex"] = False
        self.assertEqual("RECORDED_STATE_CHANGED", compare(before, after)["Classification"])

    def test_reordered_catalogue_remains_a_structural_change(self):
        before = receipt()
        before["Tabs"][0]["Controls"][0]["Choices"] = ["12", "14"]
        after = copy.deepcopy(before)
        after["Tabs"][0]["Controls"][0]["Choices"].reverse()
        self.assertFalse(compare(before, after)["TabsEqual"])


if __name__ == "__main__":
    unittest.main()
