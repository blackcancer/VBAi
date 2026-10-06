#pragma once
#include <windows.h>
#include <delayimp.h>
#include <cstdint>
#include <cstring>
#include <vector>

/// <summary>Borrowed symbol name and exact original/replacement addresses used to match named x64 imports.</summary>
struct ImportTarget { const char* name; void* original; void* replacement; };

// Used only from the owning UI thread, outside graphics callbacks. The native
// module containing replacements must remain pinned after these slots restore.
/// <summary>Retains x64 PE modules and bounded import slots for guarded replacement and restoration on the owning UI thread.</summary>
class ImportTable {
    /// <summary>Retained module handle and PE SizeOfImage used to bound relative addresses.</summary>
    struct Module { HMODULE handle; DWORD size; };
    /// <summary>Import address with its expected pointer pair and original page protection saved for restoration.</summary>
    struct Slot { void** address; void* original; void* replacement; DWORD protection = 0; bool saved = false; };
    /// <summary>At most eight owned module references; released only after every import slot restores.</summary>
    std::vector<Module> modules;
    /// <summary>At most 128 discovered import addresses, including partially installed or unrestored slots.</summary>
    std::vector<Slot> slots;
    /// <summary>Borrowed target array whose names and addresses must outlive this table.</summary>
    const ImportTarget* targets;
    /// <summary>Number of entries in the borrowed target array.</summary>
    size_t targetCount;
    /// <summary>Checks an RVA and byte extent against SizeOfImage without subtraction underflow.</summary>
    static bool Contains(const Module& m, DWORD rva, size_t size) { return rva < m.size && size <= m.size - rva; }
    /// <summary>Changes a slot only from its expected pointer and verifies original page-protection restoration.</summary>
    /// <param name="slot">Tracked import address; original protection is saved on first exchange.</param>
    /// <param name="restore">True restores the original pointer; false installs the replacement.</param>
    /// <returns>ERROR_SUCCESS after pointer/protection verification, otherwise a Win32 error.</returns>
    DWORD Exchange(Slot& slot, bool restore) {
        DWORD protection;
        if (!VirtualProtect(slot.address, sizeof(void*), PAGE_READWRITE, &protection)) return GetLastError();
        if (!slot.saved) { slot.protection = protection; slot.saved = true; }
        void* expected = restore ? slot.replacement : slot.original;
        void* desired = restore ? slot.original : slot.replacement;
        void* previous = InterlockedCompareExchangePointer(reinterpret_cast<void* volatile*>(slot.address), desired, expected);
        DWORD ignored;
        BOOL ok = VirtualProtect(slot.address, sizeof(void*), slot.protection, &ignored);
        DWORD error = ok ? ERROR_SUCCESS : GetLastError();
        MEMORY_BASIC_INFORMATION info{};
        if (!ok || VirtualQuery(slot.address, &info, sizeof(info)) != sizeof(info) || info.Protect != slot.protection)
            return error ? error : ERROR_INVALID_DATA;
        return (previous == expected || previous == desired) && *slot.address == desired ? ERROR_SUCCESS : ERROR_INVALID_DATA;
    }
    /// <summary>Scans bounded named x64 imports, skipping ordinal and unresolved delay imports.</summary>
    /// <returns>ERROR_SUCCESS after the terminator, or an error for invalid extents, conflicting pointers or exhausted capacity.</returns>
    DWORD ReadThunks(const Module& m, DWORD lookupBase, DWORD addressBase, bool delayed) {
        auto base = reinterpret_cast<BYTE*>(m.handle);
        for (DWORD index = 0; index < 65536; ++index) {
            uint64_t offset = static_cast<uint64_t>(index) * sizeof(IMAGE_THUNK_DATA64);
            if (lookupBase + offset > MAXDWORD || addressBase + offset > MAXDWORD) return ERROR_BAD_EXE_FORMAT;
            DWORD lookup = lookupBase + static_cast<DWORD>(offset), address = addressBase + static_cast<DWORD>(offset);
            if (!Contains(m, lookup, sizeof(IMAGE_THUNK_DATA64)) || !Contains(m, address, sizeof(IMAGE_THUNK_DATA64))) return ERROR_BAD_EXE_FORMAT;
            auto thunk = reinterpret_cast<IMAGE_THUNK_DATA64*>(base + lookup);
            if (!thunk->u1.AddressOfData) return ERROR_SUCCESS;
            if (IMAGE_SNAP_BY_ORDINAL64(thunk->u1.Ordinal)) continue;
            if (thunk->u1.AddressOfData > MAXDWORD - sizeof(WORD)) return ERROR_BAD_EXE_FORMAT;
            DWORD rva = static_cast<DWORD>(thunk->u1.AddressOfData) + sizeof(WORD);
            if (!Contains(m, rva, 1)) return ERROR_BAD_EXE_FORMAT;
            const char* name = reinterpret_cast<const char*>(base + rva);
            if (strnlen_s(name, m.size - rva) == m.size - rva) return ERROR_BAD_EXE_FORMAT;
            for (size_t i = 0; i < targetCount; ++i) if (!strcmp(name, targets[i].name)) {
                void** pointer = reinterpret_cast<void**>(base + address);
                Slot* known = nullptr;
                for (auto& slot : slots) if (slot.address == pointer) { known = &slot; break; }
                if (known) {
                    if (*pointer != known->replacement && *pointer != known->original) return ERROR_ALREADY_EXISTS;
                    break;
                }
                auto value = reinterpret_cast<uintptr_t>(*pointer), start = reinterpret_cast<uintptr_t>(base);
                if (delayed && value >= start && value < start + m.size) break;
                if (*pointer != targets[i].original) return ERROR_ALREADY_EXISTS;
                if (slots.size() >= 128) return ERROR_NOT_ENOUGH_MEMORY;
                slots.push_back({pointer, *pointer, targets[i].replacement});
                break;
            }
        }
        return ERROR_BAD_EXE_FORMAT;
    }
    /// <summary>Reads ordinary and RVA-form delay import directories from a retained x64 PE module.</summary>
    /// <returns>ERROR_SUCCESS or a discovery error; earlier discovered state remains on failure.</returns>
    DWORD Discover(const Module& m) {
        auto base = reinterpret_cast<BYTE*>(m.handle);
        auto dos = reinterpret_cast<IMAGE_DOS_HEADER*>(base);
        auto nt = reinterpret_cast<IMAGE_NT_HEADERS64*>(base + dos->e_lfanew);
        auto imports = nt->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_IMPORT];
        if (imports.Size) {
            if (!Contains(m, imports.VirtualAddress, imports.Size)) return ERROR_BAD_EXE_FORMAT;
            auto rows = reinterpret_cast<IMAGE_IMPORT_DESCRIPTOR*>(base + imports.VirtualAddress);
            for (DWORD i = 0; (i + 1) * sizeof(*rows) <= imports.Size; ++i) {
                if (!rows[i].Name) break;
                if (!rows[i].OriginalFirstThunk || !rows[i].FirstThunk) return ERROR_BAD_EXE_FORMAT;
                DWORD error = ReadThunks(m, rows[i].OriginalFirstThunk, rows[i].FirstThunk, false);
                if (error) return error;
            }
        }
        auto delay = nt->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT];
        if (delay.Size) {
            if (!Contains(m, delay.VirtualAddress, delay.Size)) return ERROR_BAD_EXE_FORMAT;
            auto rows = reinterpret_cast<ImgDelayDescr*>(base + delay.VirtualAddress);
            for (DWORD i = 0; (i + 1) * sizeof(*rows) <= delay.Size; ++i) {
                if (!rows[i].rvaDLLName) break;
                if (rows[i].grAttrs != dlattrRva || !rows[i].rvaINT || !rows[i].rvaIAT) return ERROR_BAD_EXE_FORMAT;
                DWORD error = ReadThunks(m, rows[i].rvaINT, rows[i].rvaIAT, true);
                if (error) return error;
            }
        }
        return ERROR_SUCCESS;
    }
public:
    /// <summary>Number of slots successfully verified during the most recent Stop attempt.</summary>
    DWORD restoredCount = 0;
    /// <summary>Borrows the target array without loading modules or installing hooks.</summary>
    ImportTable(const ImportTarget* list, size_t count) : targets(list), targetCount(count) {}
    /// <summary>Retains one validated AMD64 PE module; repeated handles need no extra reference.</summary>
    /// <returns>ERROR_SUCCESS or a loader/format/capacity error. Allocation exceptions propagate after releasing the new reference.</returns>
    DWORD AddModule(HMODULE handle) {
        if (!handle) return ERROR_MOD_NOT_FOUND;
        for (auto& item : modules) if (item.handle == handle) return ERROR_SUCCESS;
        if (modules.size() >= 8) return ERROR_NOT_ENOUGH_MEMORY;
        HMODULE retained;
        if (!GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, reinterpret_cast<LPCWSTR>(handle), &retained)) return GetLastError();
        auto base = reinterpret_cast<BYTE*>(retained);
        auto dos = reinterpret_cast<IMAGE_DOS_HEADER*>(base);
        if (dos->e_magic != IMAGE_DOS_SIGNATURE || dos->e_lfanew <= 0 || dos->e_lfanew > 1024 * 1024) { FreeLibrary(retained); return ERROR_BAD_EXE_FORMAT; }
        auto nt = reinterpret_cast<IMAGE_NT_HEADERS64*>(base + dos->e_lfanew);
        if (nt->Signature != IMAGE_NT_SIGNATURE || nt->FileHeader.Machine != IMAGE_FILE_MACHINE_AMD64 || nt->OptionalHeader.Magic != IMAGE_NT_OPTIONAL_HDR64_MAGIC) {
            FreeLibrary(retained); return ERROR_BAD_EXE_FORMAT;
        }
        try { modules.push_back({retained, nt->OptionalHeader.SizeOfImage}); }
        catch (...) { FreeLibrary(retained); throw; }
        return ERROR_SUCCESS;
    }
    /// <summary>Discovers retained modules, then installs replacements; stops at the first error without rolling back earlier slots.</summary>
    /// <returns>ERROR_SUCCESS or the first discovery/exchange error.</returns>
    DWORD Refresh() {
        for (auto& module : modules) { DWORD error = Discover(module); if (error) return error; }
        for (auto& slot : slots) if (*slot.address != slot.replacement) { DWORD error = Exchange(slot, false); if (error) return error; }
        return ERROR_SUCCESS;
    }
    /// <summary>Attempts all restorations; releases slot/module tracking only when every restoration succeeds.</summary>
    /// <returns>The first error or ERROR_SUCCESS; failed state is retained for diagnosis.</returns>
    DWORD Stop() {
        DWORD firstError = ERROR_SUCCESS; restoredCount = 0;
        for (auto& slot : slots) {
            DWORD error = slot.saved ? Exchange(slot, true) : (*slot.address == slot.original ? ERROR_SUCCESS : ERROR_INVALID_DATA);
            if (error && !firstError) firstError = error;
            if (!error) ++restoredCount;
        }
        if (!firstError) {
            slots.clear();
            for (auto& item : modules) FreeLibrary(item.handle);
            modules.clear();
        }
        return firstError;
    }
    /// <summary>Returns the tracked slot count, including discovered but unsuccessfully installed slots.</summary>
    DWORD Count() const { return static_cast<DWORD>(slots.size()); }
};
