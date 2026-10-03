"""Decode complete framed Q026 CLR guard captures. Offline; never attaches to a host."""
import argparse
import json
from pathlib import Path


def captures(text):
    result, active = [], None
    for line in text.splitlines():
        if line.startswith("Q026_GUARD_CAPTURE_ERROR "):
            raise ValueError("CLR collector reported an incomplete capture.")
        if not line.startswith("Q026_GUARD_"):
            continue
        tag, raw = line.split(" ", 1)
        frame = json.loads(raw)
        if tag == "Q026_GUARD_BEGIN":
            if active is not None or frame.get("Capture") != len(result) + 1:
                raise ValueError("Duplicate, nested or unordered guard capture.")
            if not 1 <= frame.get("Chunks", 0) <= 525 or not 1 <= frame.get("Length", 0) <= 1024 * 1024:
                raise ValueError("Guard frame exceeds its bound.")
            active = (frame, [])
        elif tag == "Q026_GUARD_CHUNK":
            if active is None or frame.get("Capture") != active[0]["Capture"] or frame.get("Index") != len(active[1]):
                raise ValueError("Missing, duplicate or unordered guard chunk.")
            piece = frame.get("Text")
            if not isinstance(piece, str) or len(piece.encode("utf-16-le", "surrogatepass")) // 2 > 2000:
                raise ValueError("Guard chunk is incomplete or oversized.")
            active[1].append(piece)
        elif tag == "Q026_GUARD_END":
            if active is None or frame != active[0] or len(active[1]) != frame["Chunks"]:
                raise ValueError("Guard end/count does not match its beginning.")
            encoded = "".join(active[1]).encode("utf-16-le", "surrogatepass")
            if len(encoded) // 2 != frame["Length"]:
                raise ValueError("Guard serialization is truncated.")
            state = json.loads(encoded.decode("utf-16-le"))
            if state.get("Capture") != frame["Capture"] or not isinstance(state.get("Tabs"), list) or not state["Tabs"]:
                raise ValueError("Captured Options tabs are absent or misidentified.")
            result.append(state)
            active = None
        else:
            raise ValueError("Unknown guard frame.")
    if active is not None:
        raise ValueError("Guard capture has no terminal frame.")
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("log", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    if args.log.stat().st_size > 16 * 1024 * 1024 or args.output.exists():
        parser.error("Use a bounded retained log and a fresh output path.")
    states = captures(args.log.read_text(encoding="utf-16"))
    with args.output.open("x", encoding="utf-8") as file:
        json.dump({"Scope": "Read-only CLR snapshots from retained log; not historical causality by itself", "Captures": states}, file, ensure_ascii=False, indent=2)
        file.write("\n")
    print("Complete guard captures:", len(states))


if __name__ == "__main__":
    main()
