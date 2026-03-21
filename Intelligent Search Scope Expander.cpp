// Intelligent Search Scope Expander.cpp : Defines the entry point for the application.
//

#include "framework.h"
#include "Intelligent Search Scope Expander.h"

#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "Comctl32.lib")
#pragma comment(lib, "Propsys.lib")
#pragma comment(lib, "OleAut32.lib")

#pragma comment(linker,"\"/manifestdependency:type='win32' \
name='Microsoft.Windows.Common-Controls' version='6.0.0.0' \
processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#define MAX_LOADSTRING 100

// ---------------------------------------------------------------------------
// Global Variables
// ---------------------------------------------------------------------------
HINSTANCE hInst;
WCHAR szTitle[MAX_LOADSTRING];
WCHAR szWindowClass[MAX_LOADSTRING];

HWND g_hWnd = nullptr;
HWND g_hBtnShow = nullptr;
HWND g_hBtnAddScope = nullptr;
HWND g_hTreeView = nullptr;
HWND g_hChkOneDrive = nullptr;
NOTIFYICONDATA g_nid = {};
bool g_treeVisible = false;
bool g_firstShowDone = false;
bool g_exiting = false;
bool g_showOneDrive = true;

std::set<std::wstring> g_searchFolders;
std::mutex g_folderMutex;
HANDLE g_monitorThread = nullptr;
volatile bool g_stopMonitor = false;
std::vector<std::wstring*> g_treeItemPaths;

// UI resources
static HFONT  g_hFontUI = nullptr;       // Segoe UI font

// ---------------------------------------------------------------------------
// Forward declarations
// ---------------------------------------------------------------------------
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

void InitTrayIcon(HWND hWnd);
void RemoveTrayIcon();
void ShowTrayContextMenu(HWND hWnd);
void MinimizeToTray(HWND hWnd);
void RestoreFromTray(HWND hWnd);

std::wstring GetDatabasePath();
std::wstring GetLogPath();
void LogFolderSource(const std::wstring& folder, const wchar_t* source);
void SaveDatabase();
void LoadDatabase();

static bool IsExcludedPath(const std::wstring& normalizedPath);
static bool IsOneDrivePath(const std::wstring& normalizedPath);

struct TaggedPath {
    std::wstring path;
    const wchar_t* source;
};

void CollectSearchFolders();
void AddFolderRecursive(const std::wstring& folder);
std::vector<std::wstring> GetRecentMRUPaths();
std::vector<std::wstring> GetExplorerFolderJumplistPaths();
std::vector<TaggedPath> GetJumplistPaths();
std::vector<std::wstring> ParseJumplistFile(const std::wstring& filePath);
std::wstring ResolveShellLinkTarget(IShellLinkW* psl);
std::wstring ExpandEnvPath(const std::wstring& path);

void PopulateTreeView();
void AddFolderToTree(HTREEITEM hParent, const std::wstring& path);
DWORD WINAPI MonitorThreadProc(LPVOID lpParam);
static void ClearTreeItemPaths();
static void InitUIResources();
static void FreeUIResources();
static void AddToWindowsSearchScope(HWND hWndParent);

// ---------------------------------------------------------------------------
// Helper: case-insensitive wstring compare for set
// ---------------------------------------------------------------------------
static std::wstring ToLower(const std::wstring& s)
{
    std::wstring result = s;
    for (auto& c : result)
        c = towlower(c);
    return result;
}

// ---------------------------------------------------------------------------
// Helper: normalize path (remove trailing backslash, lowercase)
// ---------------------------------------------------------------------------
static std::wstring NormalizePath(const std::wstring& path)
{
    std::wstring p = path;
    while (!p.empty() && (p.back() == L'\\' || p.back() == L'/'))
        p.pop_back();
    return ToLower(p);
}

// ---------------------------------------------------------------------------
// Helper: check if directory exists
// ---------------------------------------------------------------------------
static bool DirectoryExists(const std::wstring& path)
{
    DWORD attr = GetFileAttributesW(path.c_str());
    return (attr != INVALID_FILE_ATTRIBUTES && (attr & FILE_ATTRIBUTE_DIRECTORY));
}

// ---------------------------------------------------------------------------
// Helper: check if a path is a drive root (e.g. "c:", "c:\", "d:")
// ---------------------------------------------------------------------------
static bool IsDriveRoot(const std::wstring& path)
{
    std::wstring p = path;
    while (!p.empty() && (p.back() == L'\\' || p.back() == L'/'))
        p.pop_back();
    // A drive root after normalization is exactly 2 chars: letter + colon
    return (p.size() == 2 && iswalpha(p[0]) && p[1] == L':');
}

// ---------------------------------------------------------------------------
// Helper: check if a normalized path falls under an excluded system tree
// ---------------------------------------------------------------------------
static bool IsExcludedPath(const std::wstring& normalizedPath)
{
    // Build excluded prefixes once (all lowercased to match NormalizePath output)
    static std::vector<std::wstring> s_prefixes;
    static bool s_init = false;
    if (!s_init)
    {
        s_init = true;
        // Fixed well-known roots
        s_prefixes.push_back(L"c:\\windows");
        s_prefixes.push_back(L"c:\\program files");
        s_prefixes.push_back(L"c:\\program files (x86)");
        // Per-user AppData variants (roaming + local + locallow)
        WCHAR buf[MAX_PATH] = {};
        if (SUCCEEDED(SHGetFolderPathW(nullptr, CSIDL_APPDATA, nullptr, 0, buf)))
            s_prefixes.push_back(NormalizePath(buf));
        if (SUCCEEDED(SHGetFolderPathW(nullptr, CSIDL_LOCAL_APPDATA, nullptr, 0, buf)))
            s_prefixes.push_back(NormalizePath(buf));
        // LocalLow
        PWSTR pLow = nullptr;
        // FOLDERID_LocalAppDataLow
        static const GUID fid = {0xA520A1A4,0x1780,0x4FF6,{0xBD,0x18,0x16,0x73,0x43,0xC5,0xAF,0x16}};
        if (SUCCEEDED(SHGetKnownFolderPath(fid, 0, nullptr, &pLow)))
        {
            s_prefixes.push_back(NormalizePath(pLow));
            CoTaskMemFree(pLow);
        }
    }
    for (auto& prefix : s_prefixes)
    {
        if (normalizedPath.size() >= prefix.size() &&
            normalizedPath.compare(0, prefix.size(), prefix) == 0 &&
            (normalizedPath.size() == prefix.size() || normalizedPath[prefix.size()] == L'\\'))
            return true;
    }
    return false;
}

// ---------------------------------------------------------------------------
// Helper: check if a normalized path is under a OneDrive folder
// ---------------------------------------------------------------------------
static bool IsOneDrivePath(const std::wstring& normalizedPath)
{
    // Match paths like ...\onedrive\... or ...\onedrive - companyname\...
    // The component must start with "onedrive" and be preceded by a backslash.
    size_t pos = 0;
    while ((pos = normalizedPath.find(L"\\onedrive", pos)) != std::wstring::npos)
    {
        size_t afterMatch = pos + 9; // length of "\\onedrive"
        // Must be end-of-string, followed by '\\', or followed by ' ' (OneDrive - xxx)
        if (afterMatch == normalizedPath.size() ||
            normalizedPath[afterMatch] == L'\\' ||
            normalizedPath[afterMatch] == L' ')
            return true;
        pos = afterMatch;
    }
    return false;
}

// ---------------------------------------------------------------------------
// Log file path: same directory as the database
// ---------------------------------------------------------------------------
std::wstring GetLogPath()
{
    PWSTR appData = nullptr;
    SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, nullptr, &appData);
    std::wstring dir = appData;
    CoTaskMemFree(appData);
    dir += L"\\IntelligentSearchScopeExpander";
    CreateDirectoryW(dir.c_str(), nullptr);
    return dir + L"\\source_log.txt";
}

// ---------------------------------------------------------------------------
// Append a line to the log: "[timestamp] <source>: <folder>"
// ---------------------------------------------------------------------------
void LogFolderSource(const std::wstring& folder, const wchar_t* source)
{
    static std::mutex s_logMutex;
    std::lock_guard<std::mutex> lock(s_logMutex);
    FILE* f = nullptr;
    _wfopen_s(&f, GetLogPath().c_str(), L"a, ccs=UTF-8");
    if (!f) return;
    SYSTEMTIME st;
    GetLocalTime(&st);
    fwprintf(f, L"[%04d-%02d-%02d %02d:%02d:%02d] %s: %s\n",
        st.wYear, st.wMonth, st.wDay,
        st.wHour, st.wMinute, st.wSecond,
        source, folder.c_str());
    fclose(f);
}

// ---------------------------------------------------------------------------
// Helper: expand environment variables in a path
// ---------------------------------------------------------------------------
std::wstring ExpandEnvPath(const std::wstring& path)
{
    if (path.empty()) return path;
    WCHAR expanded[MAX_PATH * 2] = {};
    DWORD len = ExpandEnvironmentStringsW(path.c_str(), expanded, MAX_PATH * 2);
    if (len > 0 && len < MAX_PATH * 2)
        return expanded;
    return path;
}

// ---------------------------------------------------------------------------
// Helper: resolve a shell link to a file/folder path using GetPath + PIDL fallback
// ---------------------------------------------------------------------------
std::wstring ResolveShellLinkTarget(IShellLinkW* psl)
{
    WCHAR targetPath[MAX_PATH] = {};

    // Try GetPath with different flags
    HRESULT hr = psl->GetPath(targetPath, MAX_PATH, nullptr, 0);
    if (SUCCEEDED(hr) && targetPath[0] != L'\0')
    {
        return ExpandEnvPath(targetPath);
    }

    // Try with SLGP_RAWPATH
    memset(targetPath, 0, sizeof(targetPath));
    hr = psl->GetPath(targetPath, MAX_PATH, nullptr, SLGP_RAWPATH);
    if (SUCCEEDED(hr) && targetPath[0] != L'\0')
    {
        return ExpandEnvPath(targetPath);
    }

    // Fallback: use GetIDList + SHGetPathFromIDList for folder shortcuts
    PIDLIST_ABSOLUTE pidl = nullptr;
    hr = psl->GetIDList(&pidl);
    if (SUCCEEDED(hr) && pidl)
    {
        WCHAR pidlPath[MAX_PATH] = {};
        if (SHGetPathFromIDListW(pidl, pidlPath) && pidlPath[0] != L'\0')
        {
            CoTaskMemFree(pidl);
            return ExpandEnvPath(pidlPath);
        }
        CoTaskMemFree(pidl);
    }

    return L"";
}

// ---------------------------------------------------------------------------
// Database path: %LOCALAPPDATA%\IntelligentSearchScopeExpander\folders.dat
// ---------------------------------------------------------------------------
std::wstring GetDatabasePath()
{
    PWSTR appData = nullptr;
    SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, nullptr, &appData);
    std::wstring dir = appData;
    CoTaskMemFree(appData);
    dir += L"\\IntelligentSearchScopeExpander";
    CreateDirectoryW(dir.c_str(), nullptr);
    return dir + L"\\folders.dat";
}

void SaveDatabase()
{
    std::lock_guard<std::mutex> lock(g_folderMutex);
    std::wstring path = GetDatabasePath();
    FILE* f = nullptr;
    _wfopen_s(&f, path.c_str(), L"wb");
    if (!f) return;
    for (auto& folder : g_searchFolders)
    {
        DWORD len = (DWORD)folder.size();
        fwrite(&len, sizeof(DWORD), 1, f);
        fwrite(folder.c_str(), sizeof(wchar_t), len, f);
    }
    fclose(f);
}

void LoadDatabase()
{
    std::wstring path = GetDatabasePath();
    FILE* f = nullptr;
    _wfopen_s(&f, path.c_str(), L"rb");
    if (!f) return;
    std::lock_guard<std::mutex> lock(g_folderMutex);
    while (true)
    {
        DWORD len = 0;
        if (fread(&len, sizeof(DWORD), 1, f) != 1) break;
        if (len > 32768) break;
        std::wstring folder(len, L'\0');
        if (fread(&folder[0], sizeof(wchar_t), len, f) != len) break;
        g_searchFolders.insert(folder);
    }
    fclose(f);
}

// ---------------------------------------------------------------------------
// Collect paths from File Explorer Recent / MRU
// Resolves .lnk shortcuts in the Recent folder to their targets.
// Uses PIDL fallback for folder shortcuts where GetPath returns empty.
// ---------------------------------------------------------------------------
std::vector<std::wstring> GetRecentMRUPaths()
{
    std::vector<std::wstring> results;

    PWSTR recentPathW = nullptr;
    if (!SUCCEEDED(SHGetKnownFolderPath(FOLDERID_Recent, 0, nullptr, &recentPathW)))
        return results;

    std::wstring recentDir = recentPathW;
    CoTaskMemFree(recentPathW);
    recentPathW = nullptr;

    std::wstring searchPattern = recentDir + L"\\*";
    WIN32_FIND_DATAW fd;
    HANDLE hFind = FindFirstFileW(searchPattern.c_str(), &fd);
    if (hFind == INVALID_HANDLE_VALUE)
        return results;

    do
    {
        if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) continue;

        std::wstring name = fd.cFileName;
        std::wstring nameLower = ToLower(name);

        // Process .lnk shortcut files
        if (nameLower.size() > 4 && nameLower.substr(nameLower.size() - 4) == L".lnk")
        {
            std::wstring fullLnk = recentDir + L"\\" + name;

            IShellLinkW* psl = nullptr;
            HRESULT hr = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER,
                IID_IShellLinkW, (void**)&psl);
            if (SUCCEEDED(hr))
            {
                IPersistFile* ppf = nullptr;
                hr = psl->QueryInterface(IID_IPersistFile, (void**)&ppf);
                if (SUCCEEDED(hr))
                {
                    hr = ppf->Load(fullLnk.c_str(), STGM_READ);
                    if (SUCCEEDED(hr))
                    {
                        std::wstring target = ResolveShellLinkTarget(psl);
                        if (!target.empty())
                            results.push_back(target);
                    }
                    ppf->Release();
                }
                psl->Release();
            }
        }
    } while (FindNextFileW(hFind, &fd));
    FindClose(hFind);

    return results;
}

// ---------------------------------------------------------------------------
// Collect folder paths from File Explorer's folder jumplist / Quick Access.
// Quick Access pins are .lnk files in %USERPROFILE%\Links.
// We resolve each shortcut to its target path.
// ---------------------------------------------------------------------------
std::vector<std::wstring> GetExplorerFolderJumplistPaths()
{
    std::vector<std::wstring> results;

    // Approach 1: Enumerate .lnk files in the user's Links folder (Quick Access pins)
    PWSTR linksPathW = nullptr;
    if (SUCCEEDED(SHGetKnownFolderPath(FOLDERID_Links, 0, nullptr, &linksPathW)))
    {
        std::wstring linksDir = linksPathW;
        CoTaskMemFree(linksPathW);

        std::wstring searchPattern = linksDir + L"\\*.lnk";
        WIN32_FIND_DATAW fd;
        HANDLE hFind = FindFirstFileW(searchPattern.c_str(), &fd);
        if (hFind != INVALID_HANDLE_VALUE)
        {
            do
            {
                if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) continue;
                std::wstring fullLnk = linksDir + L"\\" + fd.cFileName;

                IShellLinkW* psl = nullptr;
                HRESULT hr = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER,
                    IID_IShellLinkW, (void**)&psl);
                if (SUCCEEDED(hr))
                {
                    IPersistFile* ppf = nullptr;
                    hr = psl->QueryInterface(IID_IPersistFile, (void**)&ppf);
                    if (SUCCEEDED(hr))
                    {
                        hr = ppf->Load(fullLnk.c_str(), STGM_READ);
                        if (SUCCEEDED(hr))
                        {
                            std::wstring target = ResolveShellLinkTarget(psl);
                            if (!target.empty())
                                results.push_back(target);
                        }
                        ppf->Release();
                    }
                    psl->Release();
                }
            } while (FindNextFileW(hFind, &fd));
            FindClose(hFind);
        }
    }

    return results;
}


// ---------------------------------------------------------------------------
// Parse a single jumplist (.automaticDestinations-ms) file using COM.
// These are OLE Compound Documents. Each stream is a serialized shell link.
// Uses ResolveShellLinkTarget for proper path + PIDL resolution.
// ---------------------------------------------------------------------------
std::vector<std::wstring> ParseJumplistFile(const std::wstring& filePath)
{
    std::vector<std::wstring> results;

    IStorage* pStorage = nullptr;
    HRESULT hr = StgOpenStorage(filePath.c_str(), nullptr,
        STGM_READ | STGM_SHARE_DENY_WRITE, nullptr, 0, &pStorage);
    if (FAILED(hr))
    {
        // Retry with STGM_SHARE_EXCLUSIVE (some files are locked by the shell)
        hr = StgOpenStorage(filePath.c_str(), nullptr,
            STGM_READ | STGM_SHARE_EXCLUSIVE, nullptr, 0, &pStorage);
    }
    if (FAILED(hr))
    {
        // Last resort: copy the file to a temp location and open the copy
        WCHAR tempDir[MAX_PATH] = {};
        WCHAR tempFile[MAX_PATH] = {};
        GetTempPathW(MAX_PATH, tempDir);
        GetTempFileNameW(tempDir, L"jl", 0, tempFile);
        if (CopyFileW(filePath.c_str(), tempFile, FALSE))
        {
            hr = StgOpenStorage(tempFile, nullptr,
                STGM_READ | STGM_SHARE_EXCLUSIVE, nullptr, 0, &pStorage);
            if (FAILED(hr))
            {
                DeleteFileW(tempFile);
                return results;
            }
            // We'll delete after reading
            auto paths = [&]() -> std::vector<std::wstring> {
                std::vector<std::wstring> inner;
                IEnumSTATSTG* pEnumInner = nullptr;
                HRESULT hr2 = pStorage->EnumElements(0, nullptr, 0, &pEnumInner);
                if (SUCCEEDED(hr2))
                {
                    STATSTG st;
                    while (pEnumInner->Next(1, &st, nullptr) == S_OK)
                    {
                        if (st.type == STGTY_STREAM && st.pwcsName)
                        {
                            IStream* pStr = nullptr;
                            hr2 = pStorage->OpenStream(st.pwcsName, nullptr,
                                STGM_READ | STGM_SHARE_EXCLUSIVE, 0, &pStr);
                            if (SUCCEEDED(hr2))
                            {
                                IShellLinkW* pLink = nullptr;
                                hr2 = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER,
                                    IID_IShellLinkW, (void**)&pLink);
                                if (SUCCEEDED(hr2))
                                {
                                    IPersistStream* pPStr = nullptr;
                                    hr2 = pLink->QueryInterface(IID_IPersistStream, (void**)&pPStr);
                                    if (SUCCEEDED(hr2))
                                    {
                                        hr2 = pPStr->Load(pStr);
                                        if (SUCCEEDED(hr2))
                                        {
                                            std::wstring t = ResolveShellLinkTarget(pLink);
                                            if (!t.empty())
                                                inner.push_back(t);
                                        }
                                        pPStr->Release();
                                    }
                                    pLink->Release();
                                }
                                pStr->Release();
                            }
                        }
                        if (st.pwcsName) CoTaskMemFree(st.pwcsName);
                    }
                    pEnumInner->Release();
                }
                return inner;
            }();
            pStorage->Release();
            DeleteFileW(tempFile);
            return paths;
        }
        return results;
    }

    IEnumSTATSTG* pEnum = nullptr;
    hr = pStorage->EnumElements(0, nullptr, 0, &pEnum);
    if (SUCCEEDED(hr))
    {
        STATSTG stat;
        while (pEnum->Next(1, &stat, nullptr) == S_OK)
        {
            if (stat.type == STGTY_STREAM && stat.pwcsName)
            {
                IStream* pStream = nullptr;
                hr = pStorage->OpenStream(stat.pwcsName, nullptr,
                    STGM_READ | STGM_SHARE_EXCLUSIVE, 0, &pStream);
                if (SUCCEEDED(hr))
                {
                    IShellLinkW* psl = nullptr;
                    hr = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER,
                        IID_IShellLinkW, (void**)&psl);
                    if (SUCCEEDED(hr))
                    {
                        IPersistStream* pps = nullptr;
                        hr = psl->QueryInterface(IID_IPersistStream, (void**)&pps);
                        if (SUCCEEDED(hr))
                        {
                            hr = pps->Load(pStream);
                            if (SUCCEEDED(hr))
                            {
                                std::wstring target = ResolveShellLinkTarget(psl);
                                if (!target.empty())
                                    results.push_back(target);
                            }
                            pps->Release();
                        }
                        psl->Release();
                    }
                    pStream->Release();
                }
            }
            if (stat.pwcsName) CoTaskMemFree(stat.pwcsName);
        }
        pEnum->Release();
    }
    pStorage->Release();
    return results;
}

// ---------------------------------------------------------------------------
// Collect paths from Application Jumplists
// (AutomaticDestinations and CustomDestinations)
// Only processes jumplists from a curated set of applications:
//   Word, Excel, PowerPoint, Notepad, Copilot, and File Explorer.
// Returns nullptr for unrecognized AppIDs (those files are skipped).
// ---------------------------------------------------------------------------
static const wchar_t* GetJumplistAppTag(const std::wstring& filename)
{
    struct AppIdEntry {
        const wchar_t* prefix;
        const wchar_t* tag;
    };
    // Curated list of known AppIDs.
    // AppIDs vary by installation (MSI vs Store vs 365), so multiple
    // variants are listed for each application.
    static const AppIdEntry s_allowed[] = {
        // File Explorer
        { L"5f7b5f1e01b83767", L"FE Folder Jumplist" },
        { L"7e4dca80246863e3", L"FE Folder Jumplist" },
        { L"1b4dd67f29cb1962", L"FE Folder Jumplist" },
        // Microsoft Word
        { L"9d1f905ce5044aee", L"Application Jumplist" },  // Word (classic/MSI)
        { L"fb3b0dbfee58fac8", L"Application Jumplist" },  // Word (Microsoft 365)
        { L"a7bd71699cd38d1c", L"Application Jumplist" },  // Word (alternate)
        // Microsoft Excel
        { L"f01b4d95cf55d32a", L"Application Jumplist" },  // Excel (classic/MSI)
        { L"2bdf5912e495ca40", L"Application Jumplist" },  // Excel (Microsoft 365 alt)
        // Microsoft PowerPoint
        { L"d00655d2aa12ff6d", L"Application Jumplist" },  // PowerPoint
        { L"b8c29862d9f95832", L"Application Jumplist" },  // PowerPoint (alternate)
        // Notepad
        { L"9b9cdc69c1c24e2b", L"Application Jumplist" },  // Notepad (Win10 classic)
        { L"b8ab77100df80ab2", L"Application Jumplist" },  // Notepad (Win11 / Store)
        { L"bc0c37eb2bf32f03", L"Application Jumplist" },  // Notepad (alternate)
        // Microsoft Copilot (Edge-hosted PWA / standalone)
        { L"e8c65855a731e95f", L"Application Jumplist" },  // Copilot (Edge PWA)
        { L"ccba5a5986c77e43", L"Application Jumplist" },  // Edge (hosts Copilot)
    };
    std::wstring lower = ToLower(filename);
    for (auto& entry : s_allowed)
    {
        if (lower.find(entry.prefix) == 0)
            return entry.tag;
    }
    return nullptr; // Not in the allowed set — skip this jumplist file
}

std::vector<TaggedPath> GetJumplistPaths()
{
    std::vector<TaggedPath> results;

    WCHAR appDataPath[MAX_PATH] = {};
    if (!SUCCEEDED(SHGetFolderPathW(nullptr, CSIDL_APPDATA, nullptr, 0, appDataPath)))
        return results;

    std::wstring autoDestDir = std::wstring(appDataPath) +
        L"\\Microsoft\\Windows\\Recent\\AutomaticDestinations";
    std::wstring customDestDir = std::wstring(appDataPath) +
        L"\\Microsoft\\Windows\\Recent\\CustomDestinations";

    // Process AutomaticDestinations (OLE compound documents)
    {
        std::wstring searchPath = autoDestDir + L"\\*.automaticDestinations-ms";
        WIN32_FIND_DATAW fd;
        HANDLE hFind = FindFirstFileW(searchPath.c_str(), &fd);
        if (hFind != INVALID_HANDLE_VALUE)
        {
            do
            {
                if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) continue;
                const wchar_t* tag = GetJumplistAppTag(fd.cFileName);
                if (!tag) continue; // Not in the allowed application set
                std::wstring fullPath = autoDestDir + L"\\" + fd.cFileName;
                auto paths = ParseJumplistFile(fullPath);
                for (auto& p : paths)
                    results.push_back({p, tag});
            } while (FindNextFileW(hFind, &fd));
            FindClose(hFind);
        }
    }

    // Process CustomDestinations (.customDestinations-ms are shell link arrays)
    {
        std::wstring searchPath = customDestDir + L"\\*.customDestinations-ms";
        WIN32_FIND_DATAW fd;
        HANDLE hFind = FindFirstFileW(searchPath.c_str(), &fd);
        if (hFind != INVALID_HANDLE_VALUE)
        {
            do
            {
                if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) continue;
                const wchar_t* tag = GetJumplistAppTag(fd.cFileName);
                if (!tag) continue; // Not in the allowed application set
                std::wstring fullPath = customDestDir + L"\\" + fd.cFileName;

                // CustomDestinations files contain concatenated shell links
                // Try to parse by opening as binary and looking for shell link headers
                HANDLE hFile = CreateFileW(fullPath.c_str(), GENERIC_READ,
                    FILE_SHARE_READ | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, 0, nullptr);
                if (hFile != INVALID_HANDLE_VALUE)
                {
                    DWORD fileSize = GetFileSize(hFile, nullptr);
                    if (fileSize > 0 && fileSize < 10 * 1024 * 1024)
                    {
                        std::vector<BYTE> data(fileSize);
                        DWORD bytesRead = 0;
                        if (ReadFile(hFile, data.data(), fileSize, &bytesRead, nullptr) && bytesRead == fileSize)
                        {
                            // Shell link magic: 4C 00 00 00
                            const BYTE magic[] = { 0x4C, 0x00, 0x00, 0x00 };
                            for (DWORD i = 0; i + 4 <= fileSize; i++)
                            {
                                if (memcmp(data.data() + i, magic, 4) == 0)
                                {
                                    // Check CLSID at offset +4: 01 14 02 00 00 00 00 00 C0 00 00 00 00 00 00 46
                                    if (i + 20 <= fileSize)
                                    {
                                        const BYTE clsid[] = { 0x01, 0x14, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00,
                                                               0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 };
                                        if (memcmp(data.data() + i + 4, clsid, 16) == 0)
                                        {
                                            // Create IStream from this data
                                            DWORD remaining = fileSize - i;
                                            IStream* pStream = SHCreateMemStream(data.data() + i, remaining);
                                            if (pStream)
                                            {
                                                IShellLinkW* psl = nullptr;
                                                HRESULT hr = CoCreateInstance(CLSID_ShellLink, nullptr,
                                                    CLSCTX_INPROC_SERVER, IID_IShellLinkW, (void**)&psl);
                                                if (SUCCEEDED(hr))
                                                {
                                                    IPersistStream* pps = nullptr;
                                                    hr = psl->QueryInterface(IID_IPersistStream, (void**)&pps);
                                                    if (SUCCEEDED(hr))
                                                    {
                                                    hr = pps->Load(pStream);
                                                        if (SUCCEEDED(hr))
                                                        {
                                                            std::wstring target = ResolveShellLinkTarget(psl);
                                                            if (!target.empty())
                                                            {
                                                                results.push_back({target, tag});
                                                            }
                                                        }
                                                        pps->Release();
                                                    }
                                                    psl->Release();
                                                }
                                                pStream->Release();
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    CloseHandle(hFile);
                }
            } while (FindNextFileW(hFind, &fd));
            FindClose(hFind);
        }
    }

    return results;
}

// ---------------------------------------------------------------------------
// Add a folder and all its subfolders recursively
// ---------------------------------------------------------------------------
void AddFolderRecursive(const std::wstring& folder)
{
    std::wstring normalized = NormalizePath(folder);
    if (normalized.empty()) return;
    if (IsExcludedPath(normalized)) return;
    if (!DirectoryExists(normalized)) return;

    {
        std::lock_guard<std::mutex> lock(g_folderMutex);
        if (g_searchFolders.count(normalized))
            return; // Already added, skip recursion
        g_searchFolders.insert(normalized);
    }

    // Enumerate subfolders
    std::wstring searchPath = normalized + L"\\*";
    WIN32_FIND_DATAW fd;
    HANDLE hFind = FindFirstFileW(searchPath.c_str(), &fd);
    if (hFind != INVALID_HANDLE_VALUE)
    {
        do
        {
            if (!(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) continue;
            if (wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0) continue;
            // Skip reparse points (junctions, symlinks) to avoid infinite loops
            if (fd.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) continue;
            // Skip hidden or system folders to avoid scanning $Recycle.Bin,
            // System Volume Information, AppData internals, etc.
            if ((fd.dwFileAttributes & FILE_ATTRIBUTE_HIDDEN) ||
                (fd.dwFileAttributes & FILE_ATTRIBUTE_SYSTEM))
                continue;

            std::wstring subDir = normalized + L"\\" + fd.cFileName;
            AddFolderRecursive(subDir);
        } while (FindNextFileW(hFind, &fd));
        FindClose(hFind);
    }
}

// ---------------------------------------------------------------------------
// Collect all search folders from all sources
// ---------------------------------------------------------------------------
void CollectSearchFolders()
{
    auto getParentFolder = [](const std::wstring& p) -> std::wstring {
        if (p.empty()) return L"";
        std::wstring parent = p;
        DWORD attr = GetFileAttributesW(parent.c_str());
        if (attr == INVALID_FILE_ATTRIBUTES || !(attr & FILE_ATTRIBUTE_DIRECTORY))
        {
            size_t pos = parent.find_last_of(L"\\/");
            if (pos != std::wstring::npos)
                parent = parent.substr(0, pos);
        }
        if (IsDriveRoot(parent)) return L"";
        if (IsExcludedPath(NormalizePath(parent))) return L"";
        return parent;
    };

    auto addAndLog = [&](const std::wstring& rawPath, const wchar_t* source) {
        std::wstring folder = getParentFolder(rawPath);
        if (folder.empty()) return;
        std::wstring norm = NormalizePath(folder);
        {
            std::lock_guard<std::mutex> lock(g_folderMutex);
            if (g_searchFolders.count(norm)) return;
        }
        LogFolderSource(folder, source);
        AddFolderRecursive(folder);
    };

    // 1. File Explorer Recent / MRU
    auto mruPaths = GetRecentMRUPaths();
    for (auto& p : mruPaths)
        addAndLog(p, L"FE MRU");

    // 2. File Explorer folder jumplist / Quick Access
    auto folderPaths = GetExplorerFolderJumplistPaths();
    for (auto& p : folderPaths)
        addAndLog(p, L"FE Folder Jumplist");

    // 3. Application Jumplists (Word, Excel, PowerPoint, Notepad, etc.)
    //    File Explorer jumplist entries are tagged as "FE Folder Jumplist"
    auto jlPaths = GetJumplistPaths();
    for (auto& tp : jlPaths)
        addAndLog(tp.path, tp.source);
}

// ---------------------------------------------------------------------------
// Background monitor thread - polls for changes every 30 seconds
// ---------------------------------------------------------------------------
DWORD WINAPI MonitorThreadProc(LPVOID /*lpParam*/)
{
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    while (!g_stopMonitor)
    {
        Sleep(30000); // 30-second poll interval
        if (g_stopMonitor) break;
        CollectSearchFolders();
    }
    CoUninitialize();
    return 0;
}

// ---------------------------------------------------------------------------
// Free all heap strings stored in tree item LPARAMs
// ---------------------------------------------------------------------------
static void ClearTreeItemPaths()
{
    for (auto* p : g_treeItemPaths)
        delete p;
    g_treeItemPaths.clear();
}

// ---------------------------------------------------------------------------
// TreeView: populate with top-level search folders
// ---------------------------------------------------------------------------
void PopulateTreeView()
{
    if (!g_hTreeView) return;
    TreeView_DeleteAllItems(g_hTreeView);
    ClearTreeItemPaths();

    std::lock_guard<std::mutex> lock(g_folderMutex);

    // Find top-level folders: folders whose parent is NOT in the set.
    // A folder is top-level if:
    //   - Its parent directory is not in g_searchFolders, OR
    //   - It has no parent (drive root like "c:"), OR
    //   - Its parent is a drive root (like "c:") — we treat direct children
    //     of drive roots as top-level since drive roots themselves are not
    //     meaningful search scope entries.
    std::vector<std::wstring> topLevel;
    for (auto& f : g_searchFolders)
    {
        if (IsDriveRoot(f)) continue;
        if (IsExcludedPath(f)) continue;
        if (!g_showOneDrive && IsOneDrivePath(f)) continue;

        std::wstring parent;
        size_t pos = f.find_last_of(L'\\');
        if (pos != std::wstring::npos)
            parent = NormalizePath(f.substr(0, pos));

        // A folder is top-level if its parent is empty, a drive root,
        // excluded, filtered-out by OneDrive toggle, or not in the set.
        bool isTopLevel = true;
        if (!parent.empty() && !IsDriveRoot(parent) && !IsExcludedPath(parent))
        {
            bool parentHiddenByOneDrive = (!g_showOneDrive && IsOneDrivePath(parent));
            if (!parentHiddenByOneDrive &&
                g_searchFolders.find(parent) != g_searchFolders.end())
            {
                isTopLevel = false;
            }
        }
        if (isTopLevel)
            topLevel.push_back(f);
    }

    for (auto& folder : topLevel)
    {
        // Show full path for top-level items for clarity
        std::wstring displayName = folder;

        // Allocate a persistent copy of the full path for LPARAM storage
        std::wstring* pathCopy = new std::wstring(folder);
        g_treeItemPaths.push_back(pathCopy);

        TVINSERTSTRUCTW tvis = {};
        tvis.hParent = TVI_ROOT;
        tvis.hInsertAfter = TVI_SORT;
        tvis.item.mask = TVIF_TEXT | TVIF_CHILDREN | TVIF_PARAM;
        tvis.item.pszText = const_cast<LPWSTR>(displayName.c_str());
        tvis.item.cChildren = 1; // indicate expandable
        tvis.item.lParam = reinterpret_cast<LPARAM>(pathCopy);
        HTREEITEM hItem = TreeView_InsertItem(g_hTreeView, &tvis);

        AddFolderToTree(hItem, folder);
    }
}

// ---------------------------------------------------------------------------
// TreeView: add child folders for a given parent
// ---------------------------------------------------------------------------
void AddFolderToTree(HTREEITEM hParent, const std::wstring& path)
{
    // Find immediate children in our set
    std::wstring prefix = path + L"\\";

    for (auto& f : g_searchFolders)
    {
        if (f.size() <= prefix.size()) continue;
        if (f.substr(0, prefix.size()) != prefix) continue;
        std::wstring remainder = f.substr(prefix.size());
        if (remainder.find(L'\\') != std::wstring::npos) continue;
        if (IsExcludedPath(f)) continue;
        if (!g_showOneDrive && IsOneDrivePath(f)) continue;

        std::wstring displayName = remainder;

        // Check if this child has its own visible children
        std::wstring childPrefix = f + L"\\";
        int hasChildren = 0;
        for (auto& cf : g_searchFolders)
        {
            if (cf.size() > childPrefix.size() && cf.substr(0, childPrefix.size()) == childPrefix)
            {
                std::wstring childRem = cf.substr(childPrefix.size());
                if (childRem.find(L'\\') == std::wstring::npos)
                {
                    if (IsExcludedPath(cf)) continue;
                    if (!g_showOneDrive && IsOneDrivePath(cf)) continue;
                    hasChildren = 1;
                    break;
                }
            }
        }

        std::wstring* pathCopy = new std::wstring(f);
        g_treeItemPaths.push_back(pathCopy);

        TVINSERTSTRUCTW tvis = {};
        tvis.hParent = hParent;
        tvis.hInsertAfter = TVI_SORT;
        tvis.item.mask = TVIF_TEXT | TVIF_CHILDREN | TVIF_PARAM;
        tvis.item.pszText = const_cast<LPWSTR>(displayName.c_str());
        tvis.item.cChildren = hasChildren;
        tvis.item.lParam = reinterpret_cast<LPARAM>(pathCopy);
        TreeView_InsertItem(g_hTreeView, &tvis);
    }
}

// ---------------------------------------------------------------------------
// TreeView: get full path from a tree item via its stored LPARAM
// ---------------------------------------------------------------------------
static std::wstring GetTreeItemFullPath(HWND hTree, HTREEITEM hItem)
{
    TVITEMW tvi = {};
    tvi.mask = TVIF_PARAM;
    tvi.hItem = hItem;
    if (TreeView_GetItem(hTree, &tvi) && tvi.lParam)
    {
        return *reinterpret_cast<std::wstring*>(tvi.lParam);
    }
    return L"";
}

// ---------------------------------------------------------------------------
// Tray icon management
// ---------------------------------------------------------------------------
void InitTrayIcon(HWND hWnd)
{
    g_nid.cbSize = sizeof(NOTIFYICONDATA);
    g_nid.hWnd = hWnd;
    g_nid.uID = 1;
    g_nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    g_nid.uCallbackMessage = WM_TRAYICON;
    g_nid.hIcon = LoadIcon(hInst, MAKEINTRESOURCE(IDI_INTELLIGENTSEARCHSCOPEEXPANDER));
    wcscpy_s(g_nid.szTip, L"Intelligent Search Scope Expander");
    Shell_NotifyIconW(NIM_ADD, &g_nid);
}

void RemoveTrayIcon()
{
    Shell_NotifyIconW(NIM_DELETE, &g_nid);
}

void ShowTrayContextMenu(HWND hWnd)
{
    POINT pt;
    GetCursorPos(&pt);

    HMENU hMenu = CreatePopupMenu();
    AppendMenuW(hMenu, MF_STRING, IDM_TRAY_OPEN, L"Open");
    AppendMenuW(hMenu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(hMenu, MF_STRING, IDM_TRAY_EXIT, L"Exit");

    SetForegroundWindow(hWnd);
    TrackPopupMenu(hMenu, TPM_BOTTOMALIGN | TPM_LEFTALIGN, pt.x, pt.y, 0, hWnd, nullptr);
    DestroyMenu(hMenu);
}

void MinimizeToTray(HWND hWnd)
{
    ShowWindow(hWnd, SW_HIDE);
}

void RestoreFromTray(HWND hWnd)
{
    ShowWindow(hWnd, SW_SHOW);
    ShowWindow(hWnd, SW_RESTORE);
    SetForegroundWindow(hWnd);
}

// ---------------------------------------------------------------------------
// Cleanup on exit
// ---------------------------------------------------------------------------
static void DoExit(HWND hWnd)
{
    g_exiting = true;
    g_stopMonitor = true;
    if (g_monitorThread)
    {
        WaitForSingleObject(g_monitorThread, 5000);
        CloseHandle(g_monitorThread);
        g_monitorThread = nullptr;
    }
    SaveDatabase();
    RemoveTrayIcon();
    DestroyWindow(hWnd);
}

// ---------------------------------------------------------------------------
// WinMain
// ---------------------------------------------------------------------------
int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);
    UNREFERENCED_PARAMETER(nCmdShow);

    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    OleInitialize(nullptr);

    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(icex);
    icex.dwICC = ICC_TREEVIEW_CLASSES;
    InitCommonControlsEx(&icex);

    InitUIResources();

    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_INTELLIGENTSEARCHSCOPEEXPANDER, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    // Start minimized (SW_HIDE)
    if (!InitInstance(hInstance, SW_HIDE))
    {
        CoUninitialize();
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_INTELLIGENTSEARCHSCOPEEXPANDER));

    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    OleUninitialize();
    CoUninitialize();
    return (int)msg.wParam;
}

// ---------------------------------------------------------------------------
// Register window class
// ---------------------------------------------------------------------------
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;
    wcex.cbSize = sizeof(WNDCLASSEX);
    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.cbClsExtra     = 0;
    wcex.cbWndExtra     = 0;
    wcex.hInstance      = hInstance;
    wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_INTELLIGENTSEARCHSCOPEEXPANDER));
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW + 1);
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_INTELLIGENTSEARCHSCOPEEXPANDER);
    wcex.lpszClassName  = szWindowClass;
    wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));
    return RegisterClassExW(&wcex);
}

// ---------------------------------------------------------------------------
// Add discovered folders to Windows Search Indexer crawl scope
// Uses CrawlScopeManager COM APIs. Skips OneDrive folders.
// ---------------------------------------------------------------------------
static void AddToWindowsSearchScope(HWND hWndParent)
{
    // Confirmation dialog
    int result = MessageBoxW(hWndParent,
        L"This will add the discovered folders below to Windows Indexer, "
        L"incurring a cost to local storage. Do you wish to proceed?",
        L"Add to Windows Search Scope",
        MB_YESNO | MB_ICONQUESTION);
    if (result != IDYES)
        return;

    // Obtain CrawlScopeManager
    ISearchManager* pSearchMgr = nullptr;
    HRESULT hr = CoCreateInstance(__uuidof(CSearchManager), nullptr,
        CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&pSearchMgr));
    if (FAILED(hr))
    {
        MessageBoxW(hWndParent, L"Failed to connect to Windows Search service.",
            L"Error", MB_OK | MB_ICONERROR);
        return;
    }

    ISearchCatalogManager* pCatalogMgr = nullptr;
    hr = pSearchMgr->GetCatalog(L"SystemIndex", &pCatalogMgr);
    if (FAILED(hr))
    {
        pSearchMgr->Release();
        MessageBoxW(hWndParent, L"Failed to open SystemIndex catalog.",
            L"Error", MB_OK | MB_ICONERROR);
        return;
    }

    ISearchCrawlScopeManager* pCSM = nullptr;
    hr = pCatalogMgr->GetCrawlScopeManager(&pCSM);
    if (FAILED(hr))
    {
        pCatalogMgr->Release();
        pSearchMgr->Release();
        MessageBoxW(hWndParent, L"Failed to get Crawl Scope Manager.",
            L"Error", MB_OK | MB_ICONERROR);
        return;
    }

    int addedCount = 0;
    int skippedCount = 0;

    {
        std::lock_guard<std::mutex> lock(g_folderMutex);
        for (auto& folder : g_searchFolders)
        {
            // Skip OneDrive folders
            if (IsOneDrivePath(folder))
            {
                skippedCount++;
                continue;
            }
            // Skip drive roots and excluded paths
            if (IsDriveRoot(folder) || IsExcludedPath(folder))
            {
                skippedCount++;
                continue;
            }

            // Check if already in crawl scope (accepts filesystem path)
            BOOL isIncluded = FALSE;
            std::wstring checkUrl = L"file:///" + folder + L"\\";
            for (size_t ci = 8; ci < checkUrl.size(); ci++)
            {
                if (checkUrl[ci] == L'\\')
                    checkUrl[ci] = L'/';
            }
            hr = pCSM->IncludedInCrawlScope(checkUrl.c_str(), &isIncluded);
            if (SUCCEEDED(hr) && isIncluded)
            {
                skippedCount++;
                continue;
            }

            // Build file:/// URL — CrawlScopeManager expects file:/// URLs
            // with forward slashes and trailing slash for inclusion rules
            std::wstring fileUrl = checkUrl;

            hr = pCSM->AddUserScopeRule(fileUrl.c_str(), TRUE, FALSE, 0);
            if (SUCCEEDED(hr))
                addedCount++;
            else
                skippedCount++;
        }
    }

    // Commit changes
    hr = pCSM->SaveAll();

    pCSM->Release();
    pCatalogMgr->Release();
    pSearchMgr->Release();

    // Show result
    WCHAR msg[256];
    _snwprintf_s(msg, _countof(msg), _TRUNCATE,
        L"Added %d folders to Windows Search scope.\n"
        L"%d folders were skipped (already indexed, OneDrive, or excluded).",
        addedCount, skippedCount);
    MessageBoxW(hWndParent, msg, L"Add to Windows Search Scope", MB_OK | MB_ICONINFORMATION);
}

// ---------------------------------------------------------------------------
// Create main window and child controls
// ---------------------------------------------------------------------------
// ---------------------------------------------------------------------------
// Modern UI resource init / cleanup
// ---------------------------------------------------------------------------
static void InitUIResources()
{
    g_hFontUI = CreateFontW(-14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Segoe UI");
}

static void FreeUIResources()
{
    if (g_hFontUI) { DeleteObject(g_hFontUI); g_hFontUI = nullptr; }
}

BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
    hInst = hInstance;

    g_hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, 0, 900, 650, nullptr, nullptr, hInstance, nullptr);

    if (!g_hWnd)
        return FALSE;

    // Create "Show Search Folders" button
    g_hBtnShow = CreateWindowW(L"BUTTON", L"Show Search Folders",
        WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        16, 12, 220, 38,
        g_hWnd, (HMENU)IDC_BTN_SHOWFOLDERS, hInstance, nullptr);

    // Create OneDrive checkbox
    g_hChkOneDrive = CreateWindowW(L"BUTTON", L"Show OneDrive Folders",
        WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
        250, 18, 220, 24,
        g_hWnd, (HMENU)IDC_CHK_ONEDRIVE, hInstance, nullptr);
    SendMessageW(g_hChkOneDrive, BM_SETCHECK, BST_CHECKED, 0);

    // Create "Add to Windows Search Scope" button
    g_hBtnAddScope = CreateWindowW(L"BUTTON", L"Add to Windows Search Scope",
        WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        490, 12, 270, 38,
        g_hWnd, (HMENU)IDC_BTN_ADDSCOPE, hInstance, nullptr);

    // Create TreeView (initially hidden)
    g_hTreeView = CreateWindowExW(0, WC_TREEVIEWW, L"",
        WS_CHILD | WS_BORDER | WS_VSCROLL | WS_HSCROLL |
        TVS_HASLINES | TVS_HASBUTTONS | TVS_LINESATROOT | TVS_DISABLEDRAGDROP |
        TVS_FULLROWSELECT,
        16, 58, 850, 540,
        g_hWnd, (HMENU)IDC_TREEVIEW, hInstance, nullptr);

    // Apply font to all controls
    if (g_hFontUI)
    {
        SendMessageW(g_hBtnShow, WM_SETFONT, (WPARAM)g_hFontUI, TRUE);
        SendMessageW(g_hChkOneDrive, WM_SETFONT, (WPARAM)g_hFontUI, TRUE);
        SendMessageW(g_hBtnAddScope, WM_SETFONT, (WPARAM)g_hFontUI, TRUE);
        SendMessageW(g_hTreeView, WM_SETFONT, (WPARAM)g_hFontUI, TRUE);
    }

    // Init tray icon
    InitTrayIcon(g_hWnd);

    // Load database from previous session
    LoadDatabase();

    bool firstRun = g_searchFolders.empty();

    if (firstRun)
    {
        // First launch: collect search folders
        CollectSearchFolders();
        SaveDatabase();
    }

    // Start background monitoring thread
    g_monitorThread = CreateThread(nullptr, 0, MonitorThreadProc, nullptr, 0, nullptr);

    ShowWindow(g_hWnd, nCmdShow);
    UpdateWindow(g_hWnd);

    return TRUE;
}

// ---------------------------------------------------------------------------
// Window procedure
// ---------------------------------------------------------------------------
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_COMMAND:
    {
        int wmId = LOWORD(wParam);
        switch (wmId)
        {
        case IDC_BTN_SHOWFOLDERS:
        {
            if (!g_firstShowDone)
            {
                g_firstShowDone = true;
                g_treeVisible = true;
                ShowWindow(g_hTreeView, SW_SHOW);
                PopulateTreeView();
                SetWindowTextW(g_hBtnShow, L"Refresh Search Folders");
                InvalidateRect(g_hBtnShow, nullptr, TRUE);
            }
            else
            {
                CollectSearchFolders();
                PopulateTreeView();
            }
        }
        break;
        case IDC_CHK_ONEDRIVE:
        {
            g_showOneDrive = (SendMessageW(g_hChkOneDrive, BM_GETCHECK, 0, 0) == BST_CHECKED);
            if (g_treeVisible)
                PopulateTreeView();
        }
        break;
        case IDC_BTN_ADDSCOPE:
        {
            AddToWindowsSearchScope(hWnd);
        }
        break;
        case IDM_ABOUT:
            DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
            break;
        case IDM_EXIT:
            DoExit(hWnd);
            break;
        case IDM_TRAY_OPEN:
            RestoreFromTray(hWnd);
            break;
        case IDM_TRAY_EXIT:
            DoExit(hWnd);
            break;
        default:
            return DefWindowProc(hWnd, message, wParam, lParam);
        }
    }
    break;

    case WM_NOTIFY:
    {
        NMHDR* pnmh = (NMHDR*)lParam;
        if (pnmh->idFrom == IDC_TREEVIEW && pnmh->code == TVN_ITEMEXPANDINGW)
        {
            NMTREEVIEWW* pnmtv = (NMTREEVIEWW*)lParam;
            if (pnmtv->action == TVE_EXPAND)
            {
                HTREEITEM hItem = pnmtv->itemNew.hItem;
                // Check if children are already populated (first child is a dummy or real)
                HTREEITEM hChild = TreeView_GetChild(g_hTreeView, hItem);
                if (!hChild)
                {
                    // Lazy-load children
                    std::wstring fullPath = GetTreeItemFullPath(g_hTreeView, hItem);
                    if (!fullPath.empty())
                    {
                        std::lock_guard<std::mutex> lock(g_folderMutex);
                        AddFolderToTree(hItem, fullPath);
                    }
                }
            }
        }
    }
    break;

    case WM_SIZE:
    {
        if (wParam != SIZE_MINIMIZED)
        {
            int cx = LOWORD(lParam);
            int cy = HIWORD(lParam);
            if (g_hTreeView)
                MoveWindow(g_hTreeView, 16, 58, cx - 32, cy - 74, TRUE);
        }
    }
    break;

    case WM_CLOSE:
        // Close button minimizes to tray instead of exiting
        MinimizeToTray(hWnd);
        return 0;

    case WM_TRAYICON:
    {
        switch (LOWORD(lParam))
        {
        case WM_LBUTTONDBLCLK:
            RestoreFromTray(hWnd);
            break;
        case WM_RBUTTONUP:
            ShowTrayContextMenu(hWnd);
            break;
        }
    }
    break;

    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);
        UNREFERENCED_PARAMETER(hdc);
        EndPaint(hWnd, &ps);
    }
    break;

    case WM_DESTROY:
        ClearTreeItemPaths();
        FreeUIResources();
        if (!g_exiting)
        {
            g_exiting = true;
            g_stopMonitor = true;
            if (g_monitorThread)
            {
                WaitForSingleObject(g_monitorThread, 3000);
                CloseHandle(g_monitorThread);
                g_monitorThread = nullptr;
            }
            SaveDatabase();
            RemoveTrayIcon();
        }
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

// ---------------------------------------------------------------------------
// About dialog
// ---------------------------------------------------------------------------
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}
