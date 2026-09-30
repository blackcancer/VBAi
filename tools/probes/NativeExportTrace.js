"use strict";

// CDB x64 script. Only matched OBJECT_ATTRIBUTES names and syscall metadata
// are logged. No file contents, token contents, handle enumeration or globals.
var trace = { pid: 0, root: "", calls: {}, next: 0, errors: 0, completed: 0, relativeUnresolved: 0 };

function initializeScript() { return [new host.apiVersionSupport(1, 3)]; }
function normalized(path) { return path.replace(/^\\\?\?\\/, "").replace(/^\\\\\?\\/, "").replace(/\/$/, "").toLowerCase(); }
function matchesPath(path, root) {
    var value = normalized(path), prefix = normalized(root);
    return value === prefix || value.indexOf(prefix + "\\") === 0;
}
function hex(value) { return "0x" + value.toString(16); }
function u32(value) { return parseInt(value.toString(16).slice(-8), 16) >>> 0; }
function execute(command) { for (var ignored of host.namespace.Debugger.Utility.Control.ExecuteCommand(command)) { } }
function emit(value) { host.diagnostics.debugLog("VBAI_FS " + JSON.stringify(value) + "\n"); }
function pointer(address) { return host.memory.readMemoryValues(address, 1, 8)[0]; }
function scalar(address, size) { return Number(host.memory.readMemoryValues(address, 1, size)[0]); }
function registers() { return host.currentThread.Registers.User; }

function configure(pid, root) {
    if (Number(host.currentProcess.Id) !== Number(pid)) throw new Error("Exact PID mismatch; trace not armed.");
    if (!/\\[0-9a-f]{32}$/i.test(root)) throw new Error("Expected fresh GUID export child; trace not armed.");
    trace.pid = Number(pid); trace.root = root;
    // Resolve first; a missing export is a preparation failure, never readiness.
    var create = host.getModuleSymbolAddress("ntdll", "NtCreateFile");
    var open = host.getModuleSymbolAddress("ntdll", "NtOpenFile");
    var query = host.getModuleSymbolAddress("ntdll", "NtQueryAttributesFile");
    var queryFull = host.getModuleSymbolAddress("ntdll", "NtQueryFullAttributesFile");
    execute('bp100 ' + hex(create) + ' "dx @$scriptContents.enter(0); gc"');
    execute('bp101 ' + hex(open) + ' "dx @$scriptContents.enter(1); gc"');
    execute('bp102 ' + hex(query) + ' "dx @$scriptContents.enter(2); gc"');
    execute('bp103 ' + hex(queryFull) + ' "dx @$scriptContents.enter(3); gc"');
    var listing = [];
    for (var line of host.namespace.Debugger.Utility.Control.ExecuteCommand("bl")) listing.push(String(line));
    if (![100, 101, 102, 103].every(function (id) {
        return listing.some(function (line) { return new RegExp("^\\s*" + id + "\\s+e\\b").test(line); });
    })) throw new Error("Entry breakpoints not verified; trace not armed.");
    host.diagnostics.debugLog("VBAI_TRACE_READY pid=" + trace.pid + "\n");
}

function enter(kind) {
    try {
        if (Number(host.currentProcess.Id) !== trace.pid) throw new Error("PID changed.");
        var reg = registers(), attributes = kind < 2 ? reg.r8 : reg.rcx;
        var objectName = pointer(attributes.add(16));
        if (objectName.compareTo(0) === 0) return;
        var length = scalar(objectName, 2);
        if (length > 2048 || (length & 1) !== 0) throw new Error("Object name exceeds bounded UTF16 read.");
        var buffer = pointer(objectName.add(8));
        var chars = host.memory.readMemoryValues(buffer, length / 2, 2), path = "";
        for (var character of chars) path += String.fromCharCode(Number(character));
        if (!matchesPath(path, trace.root)) {
            // Relative OBJECT_ATTRIBUTES cannot be attributed without enumerating
            // unrelated handle names. Report the limitation; never claim no denial.
            if (path.indexOf("\\") < 0 && /\.(frm|frx)$/i.test(path)) trace.relativeUnresolved++;
            return;
        }
        if (trace.next >= 64) throw new Error("Matched call bound exhausted.");
        var id = ++trace.next, call = {
            Id: id, Api: ["NtCreateFile", "NtOpenFile", "NtQueryAttributesFile", "NtQueryFullAttributesFile"][kind], Path: path,
            ProcessId: trace.pid, ThreadId: Number(host.currentThread.Id),
            ObjectAttributes: "0x" + scalar(attributes.add(24), 4).toString(16),
            RootDirectoryHandle: hex(pointer(attributes.add(8))),
            ReturnAddress: hex(pointer(reg.rsp)), ReturnStack: reg.rsp.add(8)
        };
        if (kind < 2) {
            call.DesiredAccess = "0x" + u32(reg.rdx).toString(16);
            call.ShareAccess = "0x" + scalar(reg.rsp.add(kind === 0 ? 56 : 40), 4).toString(16);
            call.CreateOptions = "0x" + scalar(reg.rsp.add(kind === 0 ? 72 : 48), 4).toString(16);
        }
        if (kind === 0) call.CreateDisposition = scalar(reg.rsp.add(64), 4);
        trace.calls[id] = call;
        execute('bp' + (200 + id) + ' /1 /w "@$scriptContents.isReturn(' + id + ')" ' + call.ReturnAddress +
            ' "dx @$scriptContents.recordReturn(' + id + '); gc"');
        var logged = Object.assign({}, call); delete logged.ReturnStack;
        logged.Stack = [];
        for (var frame of host.namespace.Debugger.Utility.Control.ExecuteCommand("kn 8")) logged.Stack.push(String(frame));
        emit({ Stage: "ENTRY", Call: logged });
    } catch (error) {
        // No error text/memory dump: it might disclose an unmatched argument.
        trace.errors++; emit({ Stage: "TRACE_READ_OR_BREAKPOINT_ERROR", Errors: trace.errors });
    }
}

function isReturn(id) {
    var call = trace.calls[id];
    return call !== undefined && Number(host.currentProcess.Id) === trace.pid &&
        Number(host.currentThread.Id) === call.ThreadId && registers().rsp.compareTo(call.ReturnStack) === 0;
}
function recordReturn(id) {
    // Snapshot EAX FIRST. Formatting, memory/string reads and debugger commands
    // cannot replace the syscall return with another API's LastError/HRESULT.
    var status = u32(registers().rax), call = trace.calls[id];
    if (!isReturn(id)) { trace.errors++; emit({ Stage: "UNPAIRED_RETURN", Id: id }); return; }
    trace.completed++;
    emit({ Stage: "RETURN", Id: id, Api: call.Api, Path: call.Path,
        ProcessId: trace.pid, ThreadId: call.ThreadId, NtStatus: "0x" + ("00000000" + status.toString(16)).slice(-8),
        NtSuccess: (status & 0x80000000) === 0, NtPending: status === 0x103 });
    delete trace.calls[id];
}
function summary() {
    emit({ Stage: "SUMMARY", Matched: trace.next, Completed: trace.completed,
        Pending: Object.keys(trace.calls).length, ReadOrBreakpointErrors: trace.errors,
        UnresolvedRelativeFormNames: trace.relativeUnresolved,
        Scope: "NtCreateFile/NtOpenFile/NtQueryAttributesFile/NtQueryFullAttributesFile matched absolute GUID child only; no file content. Zero calls/incomplete pairing are inconclusive, not proof of filesystem success." });
}
