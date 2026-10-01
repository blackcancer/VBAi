import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const source = fs.readFileSync(fileURLToPath(new URL('../../../../tools/probes/NativeExportTrace.js', import.meta.url)), 'utf8');
const root = 'C:\\Users\\test\\AppData\\Local\\0123456789abcdef0123456789abcdef';
const logs = [], commands = [], memory = new Map();
class Address {
    constructor(n) { this.n = n; }
    add(n) { return new Address(this.n + n); }
    // CDB's host comparison result is a boxed numeric value, unlike Node's
    // primitive number. Strict comparison with 0 silently rejected equal RSPs.
    compareTo(value) { return new Number(this.n - (value instanceof Address ? value.n : value)); }
    toString(base) { return this.n.toString(base); }
}
const registers = { r8: new Address(0x1000), rcx: new Address(0x1000), rdx: new Address(0x40000000), rsp: new Address(0x5000), rax: new Address(0xc0000022) };
const host = {
    apiVersionSupport: class {}, currentProcess: { Id: 42 }, currentThread: { Id: 7, Registers: { User: registers } },
    memory: { readMemoryValues: (address, count, size) => {
        const values = [];
        for (let i = 0; i < count; i++) {
            const item = memory.get(address.n + i * size);
            if (item === undefined) throw new Error('fixture missing memory');
            values.push(item);
        }
        return values;
    } },
    diagnostics: { debugLog: value => logs.push(value) },
    getModuleSymbolAddress: (_module, name) => new Address(name === 'NtCreateFile' ? 0x10000 : 0x20000),
    namespace: { Debugger: { Utility: { Control: { ExecuteCommand: command => {
        commands.push(command);
        if (command === 'bl') return ['100 e 00007fff`10000000 ntdll!NtCreateFile', '101 e 00007fff`10000001 ntdll!NtOpenFile',
            '102 e 00007fff`10000002 ntdll!NtQueryAttributesFile', '103 e 00007fff`10000003 ntdll!NtQueryFullAttributesFile'];
        if (command === 'kn 8') return ['0 1234 VBA!export+0x1'];
        return [];
    } } } } }
};
const context = vm.createContext({ host }); vm.runInContext(source, context);
assert.throws(() => context.configure(43, root), /PID mismatch/);
assert.throws(() => context.configure(42, root + '\\old-file'), /GUID/);
context.configure(42, root);
assert.equal(context.matchesPath('\\??\\' + root + '\\QualificationForm.frm', root), true);
assert.equal(context.matchesPath(root + 'f\\unrelated.txt', root), false);
assert.equal(context.matchesPath(root.slice(0, -32) + 'another-secret\\key.txt', root), false);
assert.equal(commands.some(c => c.includes('bc *')), false);

function setName(path) {
    memory.clear();
    memory.set(0x1010, new Address(0x2000)); memory.set(0x1008, new Address(0)); memory.set(0x1018, 0x40);
    memory.set(0x2000, path.length * 2); memory.set(0x2008, new Address(0x3000));
    [...path].forEach((c, i) => memory.set(0x3000 + i * 2, c.charCodeAt(0)));
    memory.set(0x5000, new Address(0x12345)); memory.set(0x5038, 3); memory.set(0x5040, 2); memory.set(0x5048, 0x60);
}
setName(root + '\\QualificationForm.frm'); context.enter(0);
assert.equal(context.trace.next, 1);
assert.equal(context.isReturn(1), false);
registers.rsp = new Address(0x5008);
host.currentThread.Id = 8; assert.equal(context.isReturn(1), false);
host.currentThread.Id = 7; assert.equal(context.isReturn(1), true);
context.recordReturn(1);
const returned = JSON.parse(logs.find(line => line.includes('"Stage":"RETURN"')).slice('VBAI_FS '.length));
assert.equal(returned.NtStatus, '0xc0000022'); assert.equal(returned.NtSuccess, false);
assert.equal(returned.Id, 1); assert.equal(returned.ThreadId, 7);
assert.equal(context.trace.completed, 1);
assert.equal(commands.includes('bc201'), true);
assert.equal(commands.some(c => c.includes('/1')), false);
assert.equal(commands.some(c => c.includes('/w')), false);
assert.equal(context.isReturn(1), false);
registers.rsp = new Address(0x5000); setName(root + '\\QualificationForm.frm'); context.enter(2);
assert.equal(context.trace.calls[2].Api, 'NtQueryAttributesFile');
assert.equal(context.trace.calls[2].DesiredAccess, undefined); // RDX is an output pointer for this API, never an access mask.
registers.rsp = new Address(0x5008); registers.rax = new Address(0xc0000034); context.recordReturn(2);
assert.equal(logs.some(line => line.includes('"NtStatus":"0xc0000034"')), true);
registers.rsp = new Address(0x5000); setName('C:\\unrelated\\credentials.txt');
const countBefore = logs.length; context.enter(0);
assert.equal(logs.length, countBefore); assert.equal(context.trace.next, 2);
setName('QualificationForm.frx'); context.enter(0);
assert.equal(context.trace.relativeUnresolved, 1); assert.equal(logs.length, countBefore);
memory.clear(); context.enter(0); assert.equal(context.trace.errors, 1);
assert.equal(logs.some(line => line.includes('credentials')), false);
context.summary();
assert.equal(logs.some(line => line.includes('"ReadOrBreakpointErrors":1')), true);
console.log('NativeExportTrace pure contracts PASS: PID/path privacy, paired thread/stack return, NTSTATUS capture, bounded error/inconclusive reporting. No debugger or host used.');
