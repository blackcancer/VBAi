import importlib.util
import json
from pathlib import Path
import unittest

spec = importlib.util.spec_from_file_location("guard", Path(__file__).parents[1] / "read_q026_clr_guard.py")
guard = importlib.util.module_from_spec(spec)
spec.loader.exec_module(guard)


class GuardFramesTests(unittest.TestCase):
    def frames(self, label="Font"):
        state = {"Capture": 1, "Tabs": [{"Tab": label}], "Request": {"ExpectedOptionsVersion": "0" * 64}}
        text = json.dumps(state, ensure_ascii=False)
        frame = {"Capture": 1, "Chunks": 1, "Length": len(text.encode("utf-16-le")) // 2}
        return ["Q026_GUARD_BEGIN " + json.dumps(frame),
                "Q026_GUARD_CHUNK " + json.dumps({"Capture": 1, "Index": 0, "Text": text}),
                "Q026_GUARD_END " + json.dumps(frame)]

    def test_complete_unicode_capture_preserves_values(self):
        actual = guard.captures("\n".join(self.frames("Éditeur 🛠")))
        self.assertEqual("Éditeur 🛠", actual[0]["Tabs"][0]["Tab"])

    def test_missing_end_refuses_partial_json(self):
        with self.assertRaises(ValueError): guard.captures("\n".join(self.frames()[:-1]))

    def test_duplicate_chunk_refuses(self):
        lines = self.frames()
        with self.assertRaises(ValueError): guard.captures("\n".join(lines[:2] + lines[1:]))

    def test_length_mismatch_refuses(self):
        lines = self.frames()
        for index in (0, 2): lines[index] = lines[index].replace('"Length": ', '"Length": 9')
        with self.assertRaises(ValueError): guard.captures("\n".join(lines))

    def test_collector_error_refuses_prior_success(self):
        with self.assertRaises(ValueError): guard.captures("\n".join(self.frames() + ["Q026_GUARD_CAPTURE_ERROR incomplete"]))

    def test_no_guard_is_not_invented(self):
        self.assertEqual([], guard.captures("Q026_OPTIONS_TRACE_READY pid=1\n"))

    def test_armed_exact_breakpoint_without_hit_has_no_capture(self):
        self.assertEqual([], guard.captures("Q026_GUARD_IL_ARMED offset=233 breakpoint=0\n"))

    def test_armed_receipt_preserves_complete_capture(self):
        actual = guard.captures("\n".join(["Q026_GUARD_IL_ARMED offset=233 breakpoint=0"] + self.frames()))
        self.assertEqual(1, len(actual))

    def test_wrong_offset_and_interleaved_receipt_fail(self):
        for text in ["Q026_GUARD_IL_ARMED offset=234 breakpoint=0",
                     "\n".join(self.frames()[:1] + ["Q026_GUARD_IL_ARMED offset=233 breakpoint=0"] + self.frames()[1:])]:
            with self.assertRaises(ValueError): guard.captures(text)


if __name__ == "__main__": unittest.main()
