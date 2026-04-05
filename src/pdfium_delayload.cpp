// Встраивание pdfium.dll в ресурс + отложенная загрузка (/DELAYLOAD).
// Один файл компоненты для развёртывания; извлечение в %LOCALAPPDATA%.

#if defined(_WIN32) && defined(VISUAL_ADDIN_HAVE_PDFIUM) && defined(VISUAL_ADDIN_EMBED_PDFIUM)

#include <mutex>
#include <string>

#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#include <delayimp.h>

namespace {

constexpr int kPdfiumRcDataId = 101;

std::wstring LocalAppDataDir() {
    wchar_t buf[MAX_PATH]{};
    const DWORD n = GetEnvironmentVariableW(L"LOCALAPPDATA", buf, MAX_PATH);
    if (n == 0 || n >= MAX_PATH) {
        return {};
    }
    return std::wstring(buf, buf + n);
}

HMODULE ThisModule() {
    HMODULE h = nullptr;
    GetModuleHandleExW(
        GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
        reinterpret_cast<LPCWSTR>(&LocalAppDataDir),
        &h);
    return h;
}

bool WriteResourceToFile(HMODULE module, int resId, const std::wstring& path) {
    HRSRC res = FindResourceW(module, MAKEINTRESOURCEW(resId), L"RCDATA");
    if (!res) {
        return false;
    }
    const DWORD size = SizeofResource(module, res);
    HGLOBAL mem = LoadResource(module, res);
    if (!mem || size == 0) {
        return false;
    }
    const void* data = LockResource(mem);
    if (!data) {
        return false;
    }

    const HANDLE file = CreateFileW(
        path.c_str(),
        GENERIC_WRITE,
        0,
        nullptr,
        CREATE_ALWAYS,
        FILE_ATTRIBUTE_NORMAL,
        nullptr);
    if (file == INVALID_HANDLE_VALUE) {
        return false;
    }
    DWORD written = 0;
    const BOOL ok = WriteFile(file, data, size, &written, nullptr);
    CloseHandle(file);
    return ok && written == size;
}

std::wstring EnsureEmbeddedPdfiumPath() {
    static std::wstring path;
    static std::once_flag once;

    std::call_once(once, [] {
        const std::wstring base = LocalAppDataDir();
        if (base.empty()) {
            return;
        }
        const std::wstring dir = base + L"\\VisualRecognitionAddIn";
        CreateDirectoryW(dir.c_str(), nullptr);

        const std::wstring dest = dir + L"\\pdfium.dll";
        HMODULE self = ThisModule();
        if (!self) {
            return;
        }

        HRSRC res = FindResourceW(self, MAKEINTRESOURCEW(kPdfiumRcDataId), L"RCDATA");
        const DWORD embedded = (res && self) ? SizeofResource(self, res) : 0;
        if (embedded == 0) {
            return;
        }

        WIN32_FILE_ATTRIBUTE_DATA info{};
        DWORD existing = 0;
        if (GetFileAttributesExW(dest.c_str(), GetFileExInfoStandard, &info)) {
            existing = info.nFileSizeLow;
        }
        if (existing != embedded) {
            if (!WriteResourceToFile(self, kPdfiumRcDataId, dest)) {
                return;
            }
        }
        path = dest;
    });

    return path;
}

} // namespace

extern "C" FARPROC WINAPI VisualPdfiumDelayLoadHook(unsigned dliNotify, PDelayLoadInfo pdli) {
    if (dliNotify != dliNotePreLoadLibrary || !pdli || !pdli->szDll) {
        return nullptr;
    }
    if (_stricmp(pdli->szDll, "pdfium.dll") != 0) {
        return nullptr;
    }
    const std::wstring p = EnsureEmbeddedPdfiumPath();
    if (p.empty()) {
        return nullptr;
    }
    HMODULE h = LoadLibraryW(p.c_str());
    return reinterpret_cast<FARPROC>(h);
}

extern "C" const PfnDliHook __pfnDliNotifyHook2 = VisualPdfiumDelayLoadHook;

#endif
