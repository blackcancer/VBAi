#pragma once
#include <windows.h>
#include <delayimp.h>
#include <cstdint>
#include <cstring>
#include <vector>

struct ImportTarget { const char* name; void* original; void* replacement; };

// Used only from the owning UI thread, outside graphics callbacks. The native
// module containing replacements must remain pinned after these slots restore.
class ImportTable {
    struct Module { HMODULE handle; DWORD size; };
    struct Slot { void** address; void* original; void* replacement; DWORD protection = 0; bool saved = false; };
    std::vector<Module> modules;
    std::vector<Slot> slots;
    const ImportTarget* targets;
    size_t targetCount;
    static bool Contains(const Module& m, DWORD rva, size_t size) { return rva < m.size && size <= m.size - rva; }
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
    DWORD restoredCount = 0;
    ImportTable(const ImportTarget* list, size_t count) : targets(list), targetCount(count) {}
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
    DWORD Refresh() {
        for (auto& module : modules) { DWORD error = Discover(module); if (error) return error; }
        for (auto& slot : slots) if (*slot.address != slot.replacement) { DWORD error = Exchange(slot, false); if (error) return error; }
        return ERROR_SUCCESS;
    }
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
    DWORD Count() const { return static_cast<DWORD>(slots.size()); }
};
