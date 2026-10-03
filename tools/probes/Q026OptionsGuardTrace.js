"use strict";

// Read the already-captured CLR objects only when the historical stale guard throws.
// No target function evaluation, heap search, memory writes or preference action.
let layouts = new Map();
let reads = 0;
let captures = 0;
let stage = "starting";
const zero = host.parseInt64("0");
let ownedPid = 0;
function configure(pid) {
    if (!Number.isInteger(Number(pid)) || Number(pid) < 1 || Number(host.currentProcess.Id) !== Number(pid))
        throw new Error("Exact owned process mismatch; no trace armed.");
    ownedPid = Number(pid);
    host.diagnostics.debugLog("Q026_OPTIONS_TRACE_READY pid=" + ownedPid + "\n");
}
function output(command) {
    return Array.from(host.namespace.Debugger.Utility.Control.ExecuteCommand(command));
}
function address(value) { return host.parseInt64(value.replace(/`/g, ""), 16); }
function initializeScript() { return [new host.apiVersionSupport(1, 3)]; }
function isZero(value) { return Number(value.compareTo(zero)) === 0; }
function pointer(value, offset) { return host.memory.readMemoryValues(value.add(offset), 1, 8)[0]; }
function layout(value) {
    stage = "layout " + value.toString(16);
    let mt = pointer(value, 0).toString(16);
    if (layouts.has(mt)) return layouts.get(mt);
    const lines = output("!do " + value.toString(16));
    const name = lines.find(line => /^Name:\s+/.test(line));
    if (!name) throw new Error("CLR object type was not available.");
    const fields = [];
    for (const line of lines) {
        const match = /^\s*\S+\s+\S+\s+([0-9a-f]+)\s+(.+?)\s+([01])\s+instance\s+\S+\s+(\S+)\s*$/i.exec(line);
        if (match) fields.push({ offset: parseInt(match[1], 16), type: match[2].trim(), valueType: match[3] === "1", name: match[4] });
    }
    const result = { name: name.replace(/^Name:\s+/, ""), fields: fields };
    layouts.set(mt, result);
    return result;
}
function fieldName(field) {
    const backing = /^<(.+)>(?:k__BackingField|i__Field)$/.exec(field.name);
    return backing ? backing[1] : field.name;
}
function primitive(value, field) {
    const position = value.add(field.offset);
    if (field.type === "System.Boolean") return Number(host.memory.readMemoryValues(position, 1, 1)[0]) !== 0;
    if (field.type === "System.Int32") return Number(host.memory.readMemoryValues(position, 1, 4)[0]) | 0;
    throw new Error("Unsupported CLR value type: " + field.type);
}
function fieldValue(value, field, depth) {
    stage = "field " + value.toString(16) + " " + field.name + " offset=" + field.offset;
    return field.valueType ? primitive(value, field) : decode(pointer(value, field.offset), depth + 1);
}
function decode(value, depth) {
    stage = "decode " + value.toString(16) + " depth=" + depth;
    if (isZero(value)) return null;
    if (++reads > 16000 || depth > 16) throw new Error("CLR snapshot bound exceeded.");
    const type = layout(value);
    if (type.name === "System.String") {
        const count = Number(host.memory.readMemoryValues(value.add(8), 1, 4)[0]);
        if (count > 4096) throw new Error("CLR string exceeds the native label bound.");
        if (count === 0) return "";
        const result = host.memory.readWideString(value.add(12), count);
        if (result.length !== count) throw new Error("CLR string is incomplete or contains NUL.");
        return result;
    }
    if (type.name === "System.Boolean" || type.name === "System.Int32") {
        if (type.fields.length !== 1) throw new Error("Unexpected boxed scalar layout.");
        return primitive(value, type.fields[0]);
    }
    if (type.name.endsWith("[]")) {
        const count = Number(pointer(value, 8));
        if (count < 0 || count > 2000) throw new Error("CLR reference array exceeds the catalogue bound.");
        const result = [];
        for (let index = 0; index < count; ++index) result.push(decode(pointer(value, 16 + index * 8), depth + 1));
        return result;
    }
    if (type.name.startsWith("System.Collections.Generic.List`1")) {
        const items = type.fields.find(field => field.name === "_items");
        const size = type.fields.find(field => field.name === "_size");
        if (!items || !size) throw new Error("CLR list layout is incomplete.");
        const count = primitive(value, size);
        const array = pointer(value, items.offset);
        if (count < 0 || count > 2000 || (isZero(array) && count)) throw new Error("CLR list exceeds the native bound.");
        const result = [];
        for (let index = 0; index < count; ++index) result.push(decode(pointer(array, 16 + index * 8), depth + 1));
        return result;
    }
    if (!type.name.startsWith("<>f__AnonymousType") && !type.name.startsWith("VBAi.VbeDebugWindows+Options"))
        throw new Error("Object is outside the Options snapshot graph: " + type.name);
    const result = {};
    for (const field of type.fields) result[fieldName(field)] = fieldValue(value, field, depth);
    return result;
}
function request(value) {
    const type = layout(value);
    if (type.name !== "VBAi.Request") throw new Error("Guard request type differs.");
    const result = {};
    for (const name of ["Pane", "Property", "Query", "Value", "ExpectedOptionsVersion"]) {
        const field = type.fields.find(item => fieldName(item) === name);
        if (!field) throw new Error("Missing captured request field: " + name);
        result[name] = fieldValue(value, field, 0);
    }
    return result;
}
function capture() {
    try {
        if (ownedPid && Number(host.currentProcess.Id) !== ownedPid) throw new Error("Owned trace process changed.");
        const exception = output("!pe");
        if (!exception.some(line => /^Message:\s+VBE options changed since inspection; read them again\.\s*$/.test(line))) return;
        if (++captures > 8) throw new Error("Guard capture count exceeds its bound.");
        const stack = output("!clrstack -a");
        let inside = false, capturedRequest = null;
        const candidates = [];
        for (const line of stack) {
            if (line.includes("VBAi.VbeDebugWindows.SetVbeOption(VBAi.Request, IWritableOptionsProbe)")) { inside = true; continue; }
            if (!inside) continue;
            if (/^\S+\s+\S+\s+\S/.test(line)) break;
            const parameter = /request .* = (0x[0-9a-f]+)/i.exec(line);
            if (parameter) capturedRequest = address(parameter[1]);
            const local = /^\s+0x[0-9a-f]+ = (0x[0-9a-f]+)\s*$/i.exec(line);
            if (local && local[1] !== "0x0000000000000000") candidates.push(address(local[1]));
        }
        if (!capturedRequest) throw new Error("Historical guard frame/request was not available.");
        reads = 0;
        const states = [];
        for (const value of candidates) {
            if (Number(value.compareTo(host.parseInt64("10000", 16))) < 0) continue;
            // SOS reports every JIT local slot, including uninitialized/scalar slots.
            // They are candidates only; the sole complete typed Options list is required below.
            let type;
            try { type = layout(value); } catch (ignored) { continue; }
            if (!type.name.startsWith("System.Collections.Generic.List`1[[System.Object,")) continue;
            const tabs = decode(value, 0);
            if (tabs.length && tabs.every(tab => tab && typeof tab.Tab === "string" && Array.isArray(tab.Controls) && Array.isArray(tab.FormatCategories)))
                states.push(tabs);
        }
        if (states.length !== 1) throw new Error("Guard snapshot is missing or ambiguous: " + states.length);
        const json = JSON.stringify({ Request: request(capturedRequest), Tabs: states[0],
            Source: "FirstChanceClrException.OwningGuardFrame.HeapReadOnly", Capture: captures });
        if (json.length > 1024 * 1024) throw new Error("Guard serialization exceeds its bound.");
        // WinDbg truncates one debugLog item at 16 KiB. Frame each bounded chunk;
        // consumers must verify sequence, count and complete length before parsing.
        const count = Math.ceil(json.length / 2000);
        host.diagnostics.debugLog("Q026_GUARD_BEGIN " + JSON.stringify({ Capture: captures, Chunks: count, Length: json.length }) + "\n");
        for (let index = 0; index < count; ++index)
            host.diagnostics.debugLog("Q026_GUARD_CHUNK " + JSON.stringify({ Capture: captures, Index: index, Text: json.slice(index * 2000, (index + 1) * 2000) }) + "\n");
        host.diagnostics.debugLog("Q026_GUARD_END " + JSON.stringify({ Capture: captures, Chunks: count, Length: json.length }) + "\n");
    } catch (error) {
        host.diagnostics.debugLog("Q026_GUARD_CAPTURE_ERROR " + stage + ": " + String(error) + "\n");
    }
}
