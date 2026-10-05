"""Read-only review of Q028 evidence; never launches a host or sends a request."""
import argparse
import codecs
from datetime import datetime
import hashlib
import json
import re
from pathlib import Path
import unittest
import xml.etree.ElementTree as ET

HOSTS = ("Excel", "Word", "PowerPoint", "Access", "Publisher", "Outlook")
SCENARIOS = ("managed", "tool-roundtrip", "cancel-recovery", "detached-ui") + tuple(h.lower() + "-embedded" for h in HOSTS)
STREAM_PROMPT = ("This is a synthetic interface test unrelated to VBA. Do not use tools. Output 1000 lines with exactly this format: "
                 "'1. Object 1', then '2. Object 2', then '3. Object 3', continuing the same pattern. "
                 "Begin immediately with the first line; no introductory remarks.")


def read_json(path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def require(condition, reason):
    if not condition:
        raise ValueError(reason)


def utc(value):
    value = datetime.fromisoformat(value.replace("Z", "+00:00"))
    require(value.tzinfo is not None, "Evidence timestamp lacks timezone")
    return value


def sha_file(path):
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def same_path(a, b):
    return str(Path(a).resolve()).casefold() == str(Path(b).resolve()).casefold()


def parse_response(raw, model, partial=False):
    # Cancellation may cut the final UTF-8 code point or SSE frame. Only that
    # unfinished tail is allowed; a malformed complete frame is always refused.
    text = codecs.getincrementaldecoder("utf-8-sig")().decode(raw, final=not partial)
    content, calls, done, finish, identity = "", {}, False, None, None
    lines = text.splitlines(keepends=True)
    for index, line in enumerate(lines):
        line = line.rstrip("\r\n")
        if not line or line.startswith(":"):
            continue
        require(not done, "SSE data follows DONE")
        require(line.startswith("data:"), "Unexpected SSE field or error event")
        payload = line[5:].strip()
        if payload == "[DONE]":
            require(finish is not None, "DONE without terminal finish reason")
            done = True
            continue
        try:
            chunk = json.loads(payload)
        except json.JSONDecodeError:
            if partial and index == len(lines) - 1 and not text.endswith(("\n", "\r")):
                break
            raise
        require(isinstance(chunk, dict) and "error" not in chunk, "SSE backend error")
        require(chunk.get("model") == model, "SSE model differs from request")
        require(chunk.get("object") == "chat.completion.chunk", "Unexpected SSE object")
        require(isinstance(chunk.get("id"), str) and chunk["id"], "Missing SSE response identity")
        if identity is None:
            identity = chunk["id"]
        require(chunk["id"] == identity, "Mixed SSE response identities")
        choices = chunk.get("choices")
        require(isinstance(choices, list) and len(choices) == 1 and choices[0].get("index") == 0,
                "Unexpected or ambiguous SSE choices")
        require(finish is None, "SSE choices follow terminal finish")
        choice = choices[0]
        delta = choice.get("delta")
        require(isinstance(delta, dict), "Invalid SSE delta")
        require(delta.get("role", "assistant") == "assistant", "Non-assistant SSE delta")
        piece = delta.get("content")
        require(piece is None or isinstance(piece, str), "Non-text SSE content")
        content += piece or ""
        for call in delta.get("tool_calls") or []:
            require(isinstance(call.get("index"), int) and call["index"] >= 0, "Invalid tool index")
            require(call.get("type", "function") == "function", "Unexpected tool type")
            merged = calls.setdefault(call["index"], {"id": "", "name": "", "arguments": ""})
            function = call.get("function") or {}
            for key, value in (("id", call.get("id")), ("name", function.get("name")), ("arguments", function.get("arguments"))):
                require(value is None or isinstance(value, str), "Invalid tool fragment")
                merged[key] += value or ""
        if choice.get("finish_reason") is not None:
            finish = choice["finish_reason"]
            require(finish in ("stop", "tool_calls"), "Unexpected or truncated SSE finish reason")
    require(partial or done, "Complete response lacks DONE")
    require(not done or finish is not None, "Complete response lacks finish reason")
    return {"text": content, "calls": list(calls.values()), "done": done, "finish": finish}


def response(root, row, model, partial=False):
    return parse_response((root / "embedded-wire" / (row["id"] + "-response.bin")).read_bytes(), model, partial)


def complete(row):
    receipt = row["receipt"]
    require(receipt.get("State") == "COMPLETE" and receipt.get("BackendEofObserved") is True and receipt.get("HttpStatus") == 200,
            "Required response is not complete/HTTP200")


def require_ready(parsed):
    require(parsed["done"] and parsed["finish"] == "stop" and not parsed["calls"] and parsed["text"].strip() == "UI_READY_42",
            "Next assistant response is empty, incorrect, tool-only or incomplete")


def frozen_digest(plan, path):
    rows = [f for f in plan["FrozenFiles"] if same_path(f["Path"], path)]
    require(len(rows) == 1, "Referenced evidence is not uniquely frozen")
    return rows[0]["Sha256"].lower()


def require_managed_reuse(plan, row):
    reuse = plan.get("ReuseManaged")
    require(isinstance(reuse, dict) and reuse.get("InvocationCount") == 1 and row["InvocationCount"] == 0,
            "Managed reuse lacks an original single invocation")
    require(same_path(row["Trx"], reuse["Trx"]) and same_path(row["ReusedFrom"], reuse["Campaign"]), "Managed reused source differs")
    for path in (reuse["Plan"], reuse["Campaign"], reuse["Trx"]):
        require(sha_file(path) == frozen_digest(plan, path), "Reused managed evidence hash differs")
    prior_plan, prior_campaign = read_json(Path(reuse["Plan"])), read_json(Path(reuse["Campaign"]))
    prior_rows = [r for r in prior_campaign["Scenarios"] if r["Id"] == "managed"]
    require(len(prior_rows) == 1 and prior_rows[0]["State"] == "PASS" and prior_rows[0]["InvocationCount"] == 1 and
            same_path(prior_rows[0]["Trx"], reuse["Trx"]), "Original managed invocation is not PASS once")
    prior_filter = [r["Filter"] for r in prior_plan["Scenarios"] if r["Id"] == "managed"]
    require(prior_filter == [plan["Scenarios"][0]["Filter"]] == [reuse["Filter"]], "Managed filter changed")
    expected = frozen_digest(plan, plan["TestAssembly"])
    require(expected == frozen_digest(prior_plan, prior_plan["TestAssembly"]) == reuse["TestAssemblySha256"].lower() and
            sha_file(plan["TestAssembly"]) == expected, "Managed test assembly changed")
    require(prior_plan["ProductSha256"].lower() == plan["ProductSha256"].lower() == reuse["ProductSha256"].lower() and
            prior_plan["ProductMvid"] == plan["ProductMvid"] and prior_campaign["ProductMvid"] == plan["ProductMvid"] and
            prior_campaign["ProductSha256"].lower() == plan["ProductSha256"].lower(), "Managed reused product changed")
    counters = ET.parse(reuse["Trx"]).getroot().findall(".//{*}Counters")
    require(len(counters) == 1 and int(counters[0].get("total", "0")) > 0 and counters[0].get("total") == counters[0].get("passed"),
            "Reused managed TRX is not all-pass")
    require(all(int(counters[0].get(key, "0")) == 0 for key in ("failed", "error", "timeout", "aborted", "notExecuted")),
            "Reused managed TRX contains failure or skipped tests")


def require_matrix(plan, campaign):
    require([h["Name"] for h in plan["Hosts"]] == list(HOSTS), "Unexpected host inventory")
    require([row["Id"] for row in plan["Scenarios"]] == list(SCENARIOS), "Frozen scenario matrix differs")
    require([row["Id"] for row in campaign["Scenarios"]] == list(SCENARIOS), "Executed scenario matrix differs")
    for row in campaign["Scenarios"]:
        if row["Id"] == "managed" and row["State"] == "PASS_REUSED":
            require_managed_reuse(plan, row)
            continue
        require(row["State"] == "PASS" and type(row["InvocationCount"]) is int and row["InvocationCount"] == 1,
                "A bank failed, was skipped or replayed")
    require(campaign["ProductMvid"] == plan["ProductMvid"] and campaign["ProductSha256"].lower() == plan["ProductSha256"].lower(),
            "Campaign product identity differs from frozen plan")


def require_registration(campaign):
    restored, noop = campaign.get("RegistrationRestore"), campaign.get("RegistrationNoop")
    require(bool(restored) != bool(noop), "Exactly one explicit registration terminal receipt is required")
    if restored:
        require(restored.get("Restored") is True and restored.get("Verified") is True and restored.get("Hive") == "HKCU" and restored.get("View") == "Registry64",
                "Registration was not actually restored and verified")
        return "RESTORED_AND_VERIFIED"
    require(noop.get("AlreadyIdentical") is True and noop.get("Verified") is True and noop.get("BeforeFingerprint") and noop["BeforeFingerprint"] == noop.get("AfterFingerprint"),
            "Registration no-op lacks exact unchanged baseline proof")
    return "ALREADY_IDENTICAL_VERIFIED_NOOP"


def require_lifecycle(root, plan, campaign):
    folder = root / "launcher"
    terminal = read_json(folder / "desktop" / "terminal.json")
    limited = read_json(folder / "limited-terminal.json")
    intent = read_json(folder / "intent.json")
    worker = read_json(folder / "desktop" / "worker-desktop.json")
    desktop = read_json(folder / "desktop" / "desktop-plan.json")
    started = read_json(folder / "desktop" / "campaign-started.json")
    exited = read_json(folder / "desktop" / "campaign-exit.json")
    require(limited["State"] == "CHILD_TERMINAL" and limited["ExitCode"] == 0 and limited["User"] == intent["User"] and
            limited.get("DirectGuiLauncher") is True and limited.get("ConsoleAttached") is False, "GUI launcher terminal/actor differs")
    require(terminal["State"] == "ORIGINAL_CHILD_EXIT_OBSERVED" and terminal["ExitCode"] == 0 and terminal["DesktopCloseSucceeded"] is True and
            terminal["DesktopCloseAttempted"] is True and terminal["DesktopCloseError"] == 0 and terminal["DesktopSwitches"] == 0 and terminal["OwnedForegroundObservations"] == 0,
            "Original child/desktop lifecycle is incomplete")
    require(terminal["Desktop"] == campaign["Desktop"] == desktop["Desktop"] == worker["ExpectedDesktop"] == worker["ActualDesktop"] and
            terminal["InputDesktop"] == desktop["InputDesktop"] == worker["InputDesktop"] and terminal["Desktop"] != terminal["InputDesktop"] and
            worker["DesktopSwitches"] == 0 and desktop["SwitchDesktopCalled"] is False, "Private worker desktop identity differs")
    require(started["ProcessId"] > 0 and started["ProcessId"] == exited["ProcessId"] == worker["WorkerPid"] and started["ThreadId"] == worker["ThreadId"] and
            started["OriginalHandle"] > 0 and started["OriginalHandle"] == exited["OriginalHandle"] and exited["ExitCode"] == 0, "Original worker generation/handle differs")
    require(same_path(intent["Script"], desktop["Script"]) and same_path(intent["Script"], root / "Invoke-FrozenQ028.ps1") and same_path(intent["Helper"], plan["Helper"]),
            "Launcher script/helper does not belong to this candidate")
    frozen = {str(Path(f["Path"]).resolve()).casefold(): f["Sha256"].lower() for f in plan["FrozenFiles"]}
    for path, digest in ((intent["Script"], intent["ScriptSha256"]), (intent["Helper"], intent["HelperSha256"])):
        require(frozen.get(str(Path(path).resolve()).casefold()) == digest.lower(), "Launcher bytes differ from frozen plan")


def require_embedded_owner(observed, bank, campaign):
    require(observed["ProcessId"] == bank["ProcessId"] and observed["Desktop"] == campaign["Desktop"],
            "Installed assistant process/desktop owner differs")
    require(observed.get("HostedToolWindow") is True and type(observed.get("ToolContainer")) is int and observed["ToolContainer"] > 0 and
            type(observed.get("NativeSite")) is int and observed["NativeSite"] > 0,
            "Assistant lacks the verified native ChatToolWindow container/site")
    require(observed.get("ContainerIdentity") in ("ChatToolWindow", "ControlAxSourcingSite") and
            observed.get("NativeSiteClass") == "GenericPane" and observed.get("NativeToolCaptionMatches") is True,
            "Native ActiveX container identity, GenericPane or VBAi tool caption differs")
    require(observed.get("ChildOfVbe") is True or observed.get("NativeSiteOwnedByVbe") is True,
            "Native hosted assistant lacks the original VBE relationship")



def is_stream_request(prompt):
    if prompt == STREAM_PROMPT or "Write a long numbered list of 1000 everyday objects" in prompt:
        return True
    # ChatWindow adds encoding and selected-project context before the user's
    # text. Accept the exact prompt at that boundary, never an appended suffix.
    boundary = "\n\n" + STREAM_PROMPT
    if not prompt.endswith(boundary):
        return False
    prefix = prompt[:-len(boundary)]
    closing = "\n</vbe-encoding-context>\n\n"
    return (prefix.startswith("<vbe-encoding-context>\n") and
            prefix.count("<vbe-encoding-context>") == 1 and
            prefix.count("</vbe-encoding-context>") == 1 and
            closing in prefix and bool(prefix.split(closing, 1)[1].strip()))


def is_numbered_response(text):
    return len(text) > 20 and re.search(r"(?m)^\s*1[.)]\s+\S", text) is not None



def tool_project(bank):
    value = bank.get("ToolProject")
    require(isinstance(value, str) and value.strip() and value == value.strip(), "Missing or invalid canonical ToolProject")
    return value


def require_tool_project_context(user, bank):
    # The final context line is emitted from the actual selected conversation
    # scope, independently of the fixture's native project name and prompt.
    prefix, boundary, _ = user.partition("\n\nThis is a synthetic qualification.")
    require(boundary and prefix.startswith("<vbe-encoding-context>\n") and
            "\n</vbe-encoding-context>\n\n" in prefix, "Native request lacks actual selected-project context")
    line = prefix.splitlines()[-1]
    label, separator, identifier = line.partition(": ")
    require(separator and "project" in label.casefold() and identifier == tool_project(bank),
            "Canonical ToolProject differs from actual conversation context")


def require_native_tool_arguments(arguments, bank):
    require(arguments.get("Project") == tool_project(bank) and arguments.get("Module") == "Q028Marker",
            "Native tool canonical target differs")



def require_native_read_identity(data, bank):
    # VbeSession.ReadModule echoes the resolved request selector, which is a
    # canonical path for saved documents; it does not return the native name.
    require(data["Project"] == tool_project(bank) and data["Module"] == "Q028Marker" and
            data["Sha256"] == bank["SourceBeforeSha256"] and
            hashlib.sha256(data["Code"].encode()).hexdigest() == data["Sha256"],
            "Native read canonical identity/source differs from independent baseline")


def require_native_transcript(observations, embedded, bank):
    bounds = [(index, row) for index, row in enumerate(observations) if row.get("Phase") == "NativeTranscriptBound"]
    require(len(bounds) == 1, "Native transcript binding is missing or ambiguous")
    index, bound = bounds[0]
    require(type(bound.get("ProcessId")) is int and bound["ProcessId"] == bank["ProcessId"] and
            type(bound.get("NativeThread")) is int and bound["NativeThread"] > 0 and
            type(embedded.get("NativeThread")) is int and bound["NativeThread"] == embedded["NativeThread"] and
            type(bound.get("Panel")) is int and bound["Panel"] > 0,
            "Native transcript process/thread/panel differs from the installed assistant")
    composer = [i for i, row in enumerate(observations) if row.get("Phase") == "ComposerSetIntentOnce"]
    send = [i for i, row in enumerate(observations) if row.get("Phase") == "ButtonIntentOnce" and row.get("Id") == "send"]
    require(composer and send and index < composer[0] < send[0], "Native transcript was not bound before first composer/send")


def review(root):
    plan, campaign = read_json(root / "q028-plan.json"), read_json(root / "campaign.json")
    require(campaign["State"] == "FUNCTIONAL_MATRIX_PASS_PENDING_OFFLINE_WIRE_REVIEW", "Functional matrix is not complete")
    require_matrix(plan, campaign)
    require(sha_file(plan["Product"]) == plan["ProductSha256"].lower(), "Retained product bytes differ")
    require(campaign.get("SettingsRestored") is True and campaign.get("ProxyClosed") is True and campaign.get("BackendStopped") is True,
            "Campaign resources are not restored")
    registration = require_registration(campaign)
    require_lifecycle(root, plan, campaign)
    records = []
    for request_path in sorted((root / "embedded-wire").glob("*-request.json")):
        prefix = request_path.name.removesuffix("-request.json")
        receipt = read_json(request_path.with_name(prefix + "-receipt.json"))
        request = read_json(request_path)
        require(prefix.isdecimal() and receipt["Sequence"] == int(prefix), "Wire sequence/receipt differs")
        require(request["model"] == plan["Model"] and request["stream"] is True, "Wrong model or non-streaming request")
        require(request.get("temperature") == plan["Temperature"] and request.get("top_p") == plan["TopP"], "Wire sampling profile differs")
        require(receipt.get("CaptureTruncated") is not True and receipt.get("Path") == "/v1/chat/completions", "Truncated or wrong-route wire")
        users = [m["content"] for m in request["messages"] if m.get("role") == "user"]
        require(users and isinstance(users[-1], str), "Missing last user request")
        records.append({"id": prefix, "request": request, "receipt": receipt, "user": users[-1], "start": utc(receipt["StartedUtc"]), "end": utc(receipt["CompletedUtc"])})
    require(records, "Installed assistant wire is absent")
    outcomes, used = [], set()
    previous_end = utc(campaign["StartedUtc"])
    for host in plan["Hosts"]:
        bank = read_json(root / "host-results" / (host["Name"] + "-assistant.json"))
        require(bank["State"] == "PASS" and bank["Host"] == host["Name"] and bank["SourceMvid"] == plan["ProductMvid"] and bank["ProcessId"] > 0,
                "Native bank owner or MVID differs")
        start, end = utc(bank["StartedUtc"]), utc(bank["CompletedUtc"])
        require(previous_end <= start < end <= utc(campaign["CompletedUtc"]), "Native bank chronology overlaps or differs")
        previous_end = end
        rows = [row for row in records if start <= row["start"] < end and row["start"] <= row["end"] <= utc(campaign["CompletedUtc"])]
        require(rows and not any(row["id"] in used for row in rows), "Absent or shared host wire")
        used.update(row["id"] for row in rows)
        for field in ("StreamedWhileBusy", "CancellationAcknowledged", "NextSendCompleted", "NativeMarkerVisible", "SourceAndReferencesUnchanged", "FixtureShutdownVerified"):
            require(bank.get(field) is True, host["Name"] + ": missing " + field)
        embedded = [r for r in bank["Observations"] if r.get("Phase") == "ActualEmbeddedAssistant"]
        require(len(embedded) == 1, "Installed assistant owner observation is missing or ambiguous")
        require_embedded_owner(embedded[0], bank, campaign)
        require_native_transcript(bank["Observations"], embedded[0], bank)
        project = tool_project(bank)
        candidates = [r for r in rows if "ObservedMarker" in r["user"] and ("Project=" + project + ",") in r["user"]]
        require(len(candidates) == 2, host["Name"] + ": native read requires exactly two requests")
        first, final = candidates
        for candidate in candidates:
            require_tool_project_context(candidate["user"], bank)
        ready = [r for r in rows if "Reply with exactly UI_READY_42 and nothing else." in r["user"]]
        require(len(ready) == 1 and int(ready[0]["id"]) < int(first["id"]) < int(final["id"]), "Next-send/native wire order differs")
        complete(ready[0]); require_ready(response(root, ready[0], plan["Model"]))
        stream = [r for r in rows if is_stream_request(r["user"]) and int(r["id"]) < int(ready[0]["id"])]
        numbered = []
        for row in stream:
            require(row["receipt"].get("HttpStatus") == 200, "Streaming request did not receive HTTP200")
            parsed = response(root, row, plan["Model"], partial=row["receipt"].get("State") != "COMPLETE")
            if is_numbered_response(parsed["text"]):
                numbered.append(row["id"])
        require(numbered, "No independent streamed assistant text; echoed prompt cannot qualify")
        complete(first); complete(final)
        initial = response(root, first, plan["Model"])
        require(initial["done"] and initial["finish"] in ("tool_calls", "stop") and len(initial["calls"]) == 1 and initial["calls"][0]["name"] == "read_module", "Expected exactly one direct native read_module")
        call = initial["calls"][0]
        require(call["id"], "Missing native tool call identity")
        arguments = json.loads(call["arguments"])
        require_native_tool_arguments(arguments, bank)
        results = [m for m in final["request"]["messages"] if m.get("role") == "tool" and m.get("tool_call_id") == call["id"]]
        require(len(results) == 1, "Exact native tool result is absent")
        native = json.loads(results[0]["content"])
        require(native["Ok"] is True, "Native tool was refused")
        data = native["Data"]
        require_native_read_identity(data, bank)
        markers = set(re.findall(r"NATIVE_[0-9a-f]{32}", data["Code"]))
        require(len(markers) == 1, "Native module has no unique marker")
        marker = markers.pop()
        require(hashlib.sha256(marker.encode()).hexdigest() == bank["MarkerSha256"], "Native marker hash differs")
        require(marker not in json.dumps(first["request"]) and marker not in initial["text"], "Marker was provided before the native read")
        answer = response(root, final, plan["Model"])
        require(answer["done"] and answer["finish"] == "stop" and marker in answer["text"] and not answer["calls"], "Final native answer is incomplete")
        outcomes.append({"Host": host["Name"], "ProcessId": bank["ProcessId"], "State": "PASS", "StreamRequests": numbered,
                         "NextSendRequest": ready[0]["id"], "ReadRequest": first["id"], "AnswerRequest": final["id"], "MarkerSha256": bank["MarkerSha256"]})
    require(len(used) == len(records), "Unassigned installed-assistant request exists outside qualified bank chronology")
    return {"State": "OFFLINE_NATIVE_WIRE_PASS", "ProductMvid": plan["ProductMvid"], "ProductSha256": plan["ProductSha256"],
            "Registration": registration, "Hosts": outcomes,
            "Scope": "Installed-assistant wire/marker/lifecycle review for this frozen profile only; historical empty-stream cause remains unresolved"}


class OracleTests(unittest.TestCase):
    def chunk(self, content="UI_READY_42", finish="stop", **override):
        chunk = {"model": "m", "object": "chat.completion.chunk", "id": "one", "choices": [{"index": 0, "delta": {"content": content}, "finish_reason": finish}]}
        chunk.update(override)
        return ("data: " + json.dumps(chunk) + "\n\ndata: [DONE]\n\n").encode()

    def test_completed_assistant_ready(self):
        require_ready(parse_response(self.chunk(), "m"))

    def test_echo_or_fallback_is_not_ready(self):
        for text in ("", "No text response.", "Reply with exactly UI_READY_42 and nothing else."):
            with self.assertRaises(ValueError):
                require_ready(parse_response(self.chunk(text), "m"))

    def test_missing_finish_wrong_model_and_backend_error_refused(self):
        for raw in (self.chunk(finish=None), self.chunk(model="foreign"), self.chunk(error="backend failed"), self.chunk(finish="length")):
            with self.assertRaises(ValueError):
                parse_response(raw, "m")

    def test_multiple_choices_refused(self):
        with self.assertRaises(ValueError):
            parse_response(self.chunk(choices=[{"index": 0, "delta": {}, "finish_reason": "stop"}, {"index": 1, "delta": {}, "finish_reason": "stop"}]), "m")

    def test_cancellation_allows_only_unfinished_tail(self):
        prefix = self.chunk("1. An everyday object", finish=None).split(b"\n\ndata: [DONE]")[0] + b"\n\n"
        parsed = parse_response(prefix + b'data: {"partial":', "m", partial=True)
        self.assertEqual("1. An everyday object", parsed["text"])
        with self.assertRaises(json.JSONDecodeError):
            parse_response(prefix + b'data: {"partial":\n\n', "m", partial=True)

    def test_new_and_legacy_stream_prompts_keep_strict_text_oracle(self):
        self.assertTrue(is_stream_request(STREAM_PROMPT))
        self.assertTrue(is_stream_request("Write a long numbered list of 1000 everyday objects, starting immediately with item 1."))
        self.assertFalse(is_stream_request(STREAM_PROMPT + " Another instruction."))
        self.assertTrue(is_numbered_response("1. Object 1\n2. Object 2\n3. Object 3"))
        self.assertTrue(is_numbered_response("1) Object 1\n2) Object 2\n3) Object 3"))
        self.assertFalse(is_numbered_response("1\n2\n3\n" * 10))
        self.assertFalse(is_numbered_response("1. Object 1"))

    def test_real_product_prefix_accepts_only_exact_final_stream_prompt(self):
        prefix = ("<vbe-encoding-context>\n"
                  "Before inserting a user-provided code file, call inspect_code_file on its exact path.\n"
                  "</vbe-encoding-context>\n\n"
                  "Mode de cette demande : Discussion. Analyse uniquement ; aucune modification ni exécution de macro.\r\n"
                  "Projet VBA de cette conversation : VBAProject · document non enregistré\r\n"
                  "Identifiant Project à utiliser dans les outils : VBAProject\n\n")
        self.assertTrue(is_stream_request(prefix + STREAM_PROMPT))
        for prompt in (prefix + STREAM_PROMPT + " Another instruction.", prefix + STREAM_PROMPT + "\n",
                       "Unrelated context\n\n" + STREAM_PROMPT,
                       prefix.replace("</vbe-encoding-context>", "</unknown>") + STREAM_PROMPT,
                       prefix.rstrip() + STREAM_PROMPT):
            self.assertFalse(is_stream_request(prompt))

    def test_canonical_tool_project_name_and_document_path(self):
        for project in ("VBAProject", r"E:\Private\Word\Disposable.docm"):
            bank = {"Project": "Project", "ToolProject": project}
            user = ("<vbe-encoding-context>\nEncoding guidance.\n</vbe-encoding-context>\n\n"
                    "Mode de cette demande : Discussion.\r\nProjet VBA de cette conversation : Project · Disposable.docm\r\n"
                    "Identifiant Project à utiliser dans les outils : " + project +
                    "\n\nThis is a synthetic qualification. Call read_module exactly once for Project=" + project + ", Module=Q028Marker.")
            require_tool_project_context(user, bank)
            require_native_tool_arguments({"Project": project, "Module": "Q028Marker"}, bank)
            with self.assertRaises(ValueError):
                require_native_tool_arguments({"Project": "Project", "Module": "Q028Marker"}, bank)
            with self.assertRaises(ValueError):
                require_tool_project_context(user, dict(bank, ToolProject="Project"))
        for invalid in (None, "", " ", " Project "):
            with self.assertRaises(ValueError):
                tool_project({"ToolProject": invalid})

    def test_word_read_result_echoes_canonical_selector_not_native_name(self):
        path = r"E:\Private\Word\Disposable.docm"
        code = 'Option Explicit\r\nPublic Const ObservedMarker As String = "NATIVE_0123456789abcdef0123456789abcdef"\r\n'
        digest = hashlib.sha256(code.encode()).hexdigest()
        bank = {"Project": "Project", "ToolProject": path, "SourceBeforeSha256": digest}
        data = {"Project": path, "Module": "Q028Marker", "Code": code, "Sha256": digest}
        require_native_read_identity(data, bank)
        for change in ({"Project": "Project"}, {"Project": r"E:\Private\Other\Foreign.docm"},
                       {"Module": "OtherModule"}, {"Code": code + "' changed"}):
            with self.assertRaises(ValueError):
                require_native_read_identity(dict(data, **change), bank)

    def test_native_transcript_requires_exact_owner_before_composer_and_send(self):
        bound = {"Phase": "NativeTranscriptBound", "ProcessId": 10, "NativeThread": 11, "Panel": 12}
        composer = {"Phase": "ComposerSetIntentOnce"}
        send = {"Phase": "ButtonIntentOnce", "Id": "send"}
        embedded, bank = {"NativeThread": 11}, {"ProcessId": 10}
        require_native_transcript([bound, composer, send], embedded, bank)
        for change in ({"ProcessId": 99}, {"NativeThread": 99}, {"NativeThread": 0}, {"Panel": 0}, {"Panel": True}):
            with self.assertRaises(ValueError):
                require_native_transcript([dict(bound, **change), composer, send], embedded, bank)
        for rows in ([composer, send], [bound, bound, composer, send], [composer, bound, send],
                     [bound, send, composer], [bound, composer]):
            with self.assertRaises(ValueError):
                require_native_transcript(rows, embedded, bank)

    def test_empty_matrix_refused(self):
        with self.assertRaises(ValueError):
            require_matrix({"Hosts": [], "Scenarios": []}, {"Scenarios": []})

    def test_docked_and_floating_native_sites_are_supported(self):
        observed = {"ProcessId": 10, "Desktop": "private", "HostedToolWindow": True,
                    "ToolContainer": 20, "NativeSite": 30, "ChildOfVbe": True, "NativeSiteOwnedByVbe": False,
                    "ContainerIdentity": "ChatToolWindow", "NativeSiteClass": "GenericPane", "NativeToolCaptionMatches": True}
        require_embedded_owner(observed, {"ProcessId": 10}, {"Desktop": "private"})
        observed.update(ChildOfVbe=False, NativeSiteOwnedByVbe=True, ContainerIdentity="ControlAxSourcingSite")
        require_embedded_owner(observed, {"ProcessId": 10}, {"Desktop": "private"})

    def test_detached_foreign_or_unattached_sites_are_refused(self):
        valid = {"ProcessId": 10, "Desktop": "private", "HostedToolWindow": True,
                 "ToolContainer": 20, "NativeSite": 30, "ChildOfVbe": True, "NativeSiteOwnedByVbe": False,
                    "ContainerIdentity": "ChatToolWindow", "NativeSiteClass": "GenericPane", "NativeToolCaptionMatches": True}
        for change in ({"HostedToolWindow": False}, {"ToolContainer": 0}, {"NativeSite": 0},
                       {"ChildOfVbe": False, "NativeSiteOwnedByVbe": False}, {"ProcessId": 99}, {"Desktop": "Default"},
                       {"ContainerIdentity": "UnrelatedPane"}, {"NativeSiteClass": "WindowsForms"}, {"NativeToolCaptionMatches": False}):
            with self.assertRaises(ValueError):
                require_embedded_owner(dict(valid, **change), {"ProcessId": 10}, {"Desktop": "private"})

    def test_reuse_without_original_once_receipt_refused(self):
        with self.assertRaises(ValueError):
            require_managed_reuse({"ReuseManaged": {}}, {"InvocationCount": 0})

    def test_failed_restore_is_not_successful_noop(self):
        with self.assertRaises(ValueError):
            require_registration({"RegistrationRestore": {"Restored": False, "Verified": True}})
        self.assertEqual("ALREADY_IDENTICAL_VERIFIED_NOOP", require_registration({"RegistrationNoop": {
            "AlreadyIdentical": True, "Verified": True, "BeforeFingerprint": "same", "AfterFingerprint": "same"}}))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("evidence_root", type=Path, nargs="?")
    parser.add_argument("--self-test", action="store_true", help="Run pure synthetic refusal-oracle tests; no host or requests")
    args = parser.parse_args()
    if args.self_test:
        unittest.main(argv=[__file__], exit=True)
    if args.evidence_root is None:
        parser.error("evidence_root is required unless --self-test is selected")
    try:
        result = review(args.evidence_root.resolve())
    except (ValueError, KeyError, TypeError, AttributeError, OSError, ET.ParseError) as error:
        print(json.dumps({"State": "FAILED_OR_BLOCKED", "Qualified": False, "Error": str(error)}))
        raise SystemExit(1)
    print(json.dumps(result, indent=2))
