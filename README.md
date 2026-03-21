# Intelligent Search Scope Expander

This application was created as an **experiment** to understand what should be the ideal method of scope expansion for local file searchability on Windows that is grounded on user activity signals. Rather than indexing the entire filesystem, the app observes which files and folders the user actually interacts with — via File Explorer history, Quick Access pins, and application jumplists — and uses those signals to build a focused, user-activity-driven search scope. The goal is to explore whether this approach produces a more relevant and efficient set of searchable locations than a blanket full-disk index.

The app is a Windows desktop application that automatically discovers and maintains a list of **Search Folders** based on the user's file activity. It starts minimized in the system tray and continuously monitors the user's recent file and folder usage to build and expand a comprehensive folder index.

---

## How It Works

### Overview

On first launch, the app scans three categories of Windows Shell data to discover which files and folders the user actively works with:

1. **File Explorer MRU (Most Recently Used)** — the shortcuts in the user's `Recent` folder
2. **File Explorer Folder Jumplist / Quick Access** — pinned and frequent folders accessible via the shell namespace
3. **Application Jumplists** — recent/pinned items from **Word, Excel, PowerPoint, Notepad, and Copilot** only (other applications are intentionally excluded to keep the scope focused on productivity signals)

For every discovered file path, the app extracts the **parent folder**. For every discovered folder path, it uses the folder directly. It then recursively enumerates **all subfolders** of each parent and adds them to the Search Folders list.

The app keeps running in the background, polling for changes every 30 seconds and adding any newly discovered folders to the list. On exit, the list is persisted to a local database file and reloaded on the next startup.

### User Interface

The app uses a **light-themed UI** with the Segoe UI font and standard Windows common controls.

| Element | Behavior |
|---|---|
| **System tray icon** | Always present while the app is running. Double-click to open the window. Right-click for a context menu with *Open* and *Exit*. |
| **Main window** | Standard light window with buttons, a checkbox, and a tree view control. The close (X) button **minimizes to tray** — it does not exit the app. |
| **"Show Search Folders" button** | First click: displays the tree view and changes its label to **"Refresh Search Folders"**. Subsequent clicks: re-collects folders from all sources and refreshes the tree. |
| **"Show OneDrive Folders" checkbox** | Toggles visibility of OneDrive-synced folders in the tree. Checked by default. Unchecking immediately hides all paths containing `OneDrive` from the tree (without removing them from the database). |
| **"Add to Windows Search Scope" button** | Adds all discovered folders to the Windows Search Indexer's crawl scope using the CrawlScopeManager COM API. OneDrive folders are always excluded from this operation. Shows a confirmation dialog before proceeding, then reports how many folders were added vs. skipped. |
| **Tree view** | Standard tree view with expandable hierarchy. Top-level items show their full path for clarity. |
| **Menu ? File ? Exit** | The only way to exit the app (besides the tray icon right-click ? Exit). Saves the database before shutting down. |

### Excluded Folders

The following system folders and all their subfolders are **always excluded** from the search folder list:

| Excluded Path | Reason |
|---|---|
| `C:\Windows` | OS system files |
| `C:\Program Files` | Installed application binaries |
| `C:\Program Files (x86)` | 32-bit application binaries |
| `%APPDATA%` (Roaming) | Application configuration data |
| `%LOCALAPPDATA%` | Local application data / caches |
| `%LOCALAPPDATA%Low` | Low-integrity app data |

Exclusion is applied at collection time (`AddFolderRecursive` and `getParentFolder`) and at display time (`PopulateTreeView` and `AddFolderToTree`).

### Windows Search Indexer Integration

The **"Add to Windows Search Scope"** button uses the Windows Search [CrawlScopeManager](https://learn.microsoft.com/en-us/windows/win32/search/-search-3x-wds-extidx-csm) COM APIs to add discovered folders to the system indexer:

1. Creates `CSearchManager` ? `ISearchCatalogManager` ("SystemIndex") ? `ISearchCrawlScopeManager`
2. For each folder in `g_searchFolders`:
   - **Skips** OneDrive paths, drive roots, and excluded system paths
   - Checks `ISearchCrawlScopeManager::IncludedInCrawlScope` to avoid duplicates
   - Calls `ISearchCrawlScopeManager::AddUserScopeRule` with a `file:///` URL for new folders
3. Calls `ISearchCrawlScopeManager::SaveAll` to commit all changes

The folders are added as **user scope rules** (inclusion), which means they appear in the Windows Indexing Options control panel and can be managed by the user. The indexer will begin crawling these locations in the background. **OneDrive folders are always excluded** from this operation since OneDrive content is indexed separately by the cloud file provider.

### Source Logging

Every time a **new** parent folder is discovered, the app appends a line to:
```
%LOCALAPPDATA%\IntelligentSearchScopeExpander\source_log.txt
```
Format: `[YYYY-MM-DD HH:MM:SS] <source>: <folder path>`

Source labels:
- `FE MRU` — File Explorer Most Recently Used (Recent folder `.lnk` files)
- `FE Folder Jumplist` — File Explorer Quick Access / Links folder / Explorer taskbar jumplist
- `Application Jumplist` — Word, Excel, PowerPoint, Notepad, Copilot only

---

## Architecture

### Project Structure

```
Intelligent Search Scope Expander/
??? Intelligent Search Scope Expander.cpp   # All application logic
??? Intelligent Search Scope Expander.h     # App header (includes resource.h)
??? framework.h                             # Precompiled header (Win32 + Shell + STL includes)
??? targetver.h                             # Windows SDK version targeting
??? Resource.h                              # Resource IDs (menus, icons, controls)
??? Intelligent Search Scope Expander.rc    # Resource script (menus, dialogs, icons, strings)
??? Intelligent Search Scope Expander.ico   # Main application icon (expand icon)
??? small.ico                               # Small icon variant
??? Intelligent Search Scope Expander.sln   # Visual Studio solution
??? Intelligent Search Scope Expander.vcxproj       # Project file
??? Intelligent Search Scope Expander.vcxproj.filters
```

The entire application is implemented in a single C++ source file (`Intelligent Search Scope Expander.cpp`) using the Win32 API with C++14.

### Component Breakdown

#### 1. Application Lifecycle (`wWinMain`, `InitInstance`)

```
wWinMain
  ?? CoInitializeEx (STA — required for Shell COM objects)
  ?? OleInitialize (full OLE support for shell operations)
  ?? InitCommonControlsEx (TreeView control registration)
  ?? MyRegisterClass (WNDCLASSEX with app icon)
  ?? InitInstance (SW_HIDE — starts hidden)
       ?? CreateWindowW (main window, 800×600)
       ?? Create "Show Search Folders" button
       ?? Create TreeView control (initially hidden)
       ?? InitTrayIcon (NOTIFYICONDATA ? Shell_NotifyIcon)
       ?? LoadDatabase (restore previous session)
       ?? CollectSearchFolders (if first run / empty DB)
       ?? CreateThread ? MonitorThreadProc (background polling)
```

The app starts **hidden** (`SW_HIDE`). The user interacts via the system tray icon.

> **Important**: COM is initialized with `COINIT_APARTMENTTHREADED` (STA), not MTA. Shell COM objects such as `IShellLink`, `IPersistFile`, `IShellFolder`, and `IStorage` are apartment-threaded and will silently fail if used from an MTA thread.

#### 2. Data Collection Pipeline

Three independent data sources feed into a single `std::set<std::wstring>` (`g_searchFolders`):

##### Source A: `GetRecentMRUPaths()`

- Reads the `FOLDERID_Recent` directory (typically `%APPDATA%\Microsoft\Windows\Recent`)
- Enumerates all `.lnk` shortcut files
- For each shortcut, creates an `IShellLink` object, loads it via `IPersistFile`, then calls `ResolveShellLinkTarget()` to get the actual file/folder path
- Returns all resolved target paths

##### Source B: `GetExplorerFolderJumplistPaths()`

- Enumerates `.lnk` shortcut files in the user's **Links folder** (`FOLDERID_Links`, typically `%USERPROFILE%\Links`)
- This folder contains Quick Access pinned items (e.g., Desktop, Downloads, and any user-pinned folders)
- Each shortcut is resolved via `IShellLink` + `IPersistFile::Load` ? `ResolveShellLinkTarget()`, identical to the MRU approach
- File Explorer's folder jumplist entries (right-click Explorer on the taskbar) are stored in the AutomaticDestinations files and are picked up by Source C below

##### Source C: `GetJumplistPaths()` + `ParseJumplistFile()`

- Scans `%APPDATA%\Microsoft\Windows\Recent\AutomaticDestinations\` for `.automaticDestinations-ms` files
- **AppID allowlist**: Only jumplist files from a curated set of applications are processed. Each `.automaticDestinations-ms` filename starts with a hex AppID. The function `GetJumplistAppTag()` matches against known AppIDs for:
  - **File Explorer** (tagged `FE Folder Jumplist`)
  - **Word, Excel, PowerPoint, Notepad, Copilot** (tagged `Application Jumplist`)
  - All other AppIDs are **skipped** to keep the scope focused on productivity signals
- Multiple AppID variants are listed per application (MSI, Microsoft 365, Store, alternate) since the ID depends on the installation method
- Opens each allowed file with `StgOpenStorage`, enumerates streams with `IEnumSTATSTG`, deserializes each stream as an `IShellLink` via `IPersistStream::Load`, then resolves the path with `ResolveShellLinkTarget()`
- **Locked file handling**: Jumplist files are actively locked by the Windows shell. The code tries `STGM_SHARE_DENY_WRITE` first, falls back to `STGM_SHARE_EXCLUSIVE`, and as a last resort copies the file to `%TEMP%` and opens the copy
- Scans `%APPDATA%\Microsoft\Windows\Recent\CustomDestinations\` for `.customDestinations-ms` files (same AppID filter applies)
- Parses allowed files by scanning for the shell link binary magic (`4C 00 00 00` + CLSID `00021401-0000-0000-C000-000000000046`), wrapping each found link in an `IStream` via `SHCreateMemStream`, and deserializing/resolving as above

##### Path Resolution: `ResolveShellLinkTarget()`

A unified helper that extracts the target path from any `IShellLink`:

1. **Primary**: `IShellLink::GetPath()` with default flags (0), then `ExpandEnvironmentStringsW()` to resolve `%USERPROFILE%` etc.
2. **Secondary**: `IShellLink::GetPath()` with `SLGP_RAWPATH` flag + environment expansion.
3. **PIDL Fallback**: `IShellLink::GetIDList()` ? `SHGetPathFromIDListW()` for folder shortcuts and virtual items that have no traditional file path.

This three-step approach is critical because folder shortcuts in File Explorer's jumplist typically store a PIDL rather than a file path, and different shortcut types respond to different `GetPath` flag combinations.

##### Folder Expansion: `AddFolderRecursive()`

For each collected path:
1. `CollectSearchFolders()` determines if the path is a file or directory. Files have their parent folder extracted.
2. **Drive root filtering**: If the parent folder is a drive root (e.g., `C:\`), it is skipped entirely. Recursing from a drive root would scan tens or hundreds of thousands of folders across the entire drive, which is not a meaningful search scope.
3. `AddFolderRecursive()` normalizes the path (lowercase, no trailing slash), checks `DirectoryExists()`, and inserts into `g_searchFolders` (guarded by `g_folderMutex`).
4. It then enumerates immediate subdirectories (via `FindFirstFileW/FindNextFileW`) and recurses, skipping:
   - Reparse points (junctions, symlinks) to avoid infinite loops
   - Hidden folders (e.g., `$Recycle.Bin`, `.git`)
   - System folders (e.g., `System Volume Information`)
5. If a folder is already in the set, recursion stops (deduplication).

#### 3. Background Monitoring (`MonitorThreadProc`)

A dedicated thread polls every **30 seconds**:

```
MonitorThreadProc
  ?? CoInitializeEx (STA — each thread needs its own COM init)
  ?? Loop:
  ?    ?? Sleep(30000)
  ?    ?? Check g_stopMonitor flag
  ?    ?? CollectSearchFolders()
  ?         ?? GetRecentMRUPaths() ? extract parents ? AddFolderRecursive()
  ?         ?? GetExplorerFolderJumplistPaths() ? extract parents ? AddFolderRecursive()
  ?         ?? GetJumplistPaths() ? extract parents ? AddFolderRecursive()
  ?? CoUninitialize
```

Since `AddFolderRecursive()` skips folders already in the set, repeated polls are efficient — only genuinely new folders cause filesystem enumeration.

Thread synchronization uses a `std::mutex` (`g_folderMutex`) protecting all reads/writes to `g_searchFolders`.

> **Note**: The monitor thread initializes COM independently with `COINIT_APARTMENTTHREADED` because Shell COM objects require STA and each thread must call `CoInitializeEx` separately.

#### 4. Persistence (`SaveDatabase` / `LoadDatabase`)

The database is a simple binary file at:
```
%LOCALAPPDATA%\IntelligentSearchScopeExpander\folders.dat
```

**Format**: A sequence of records, each consisting of:
| Field | Type | Description |
|---|---|---|
| Length | `DWORD` (4 bytes) | Number of `wchar_t` characters in the path |
| Path | `wchar_t[Length]` | The folder path (no null terminator) |

- **Save** occurs on app exit (both `IDM_EXIT` and `IDM_TRAY_EXIT`), under `g_folderMutex`.
- **Load** occurs at startup in `InitInstance`, before the first collection.
- On **first run** (empty database), the app immediately runs a full collection.

#### 5. TreeView Display (`PopulateTreeView` / `AddFolderToTree`)

The tree view identifies **top-level folders**: any folder in `g_searchFolders` that meets one of these criteria:
- Its parent directory is not in `g_searchFolders`
- Its parent is a drive root (e.g., `C:\`) — direct children of drive roots are promoted to top-level since drive roots themselves are excluded as non-meaningful entries
- It has no parent (edge case)

Bare drive roots (like `c:`) are always excluded from the tree.

Top-level items display their **full path** (e.g., `c:\experiments\testindexerdata`) for clarity, while child items show only their folder name.

Each tree item stores a pointer to its full path string in the Win32 `TVITEM::lParam` field. A global `std::vector<std::wstring*>` (`g_treeItemPaths`) owns these heap-allocated strings and frees them when the tree is cleared or the app exits.

For child items, `AddFolderToTree()` finds immediate children by prefix-matching on `path + "\\"` and checking the remainder has no additional backslashes. Each child checks for its own children to set the `cChildren` flag (enabling the expand button).

On expand (`TVN_ITEMEXPANDING`), the full path is read directly from the tree item's `lParam`.

#### 6. System Tray (`InitTrayIcon` / tray message handling)

| Action | Result |
|---|---|
| Tray double-click | `RestoreFromTray()` — `ShowWindow(SW_SHOW)` + `SetForegroundWindow()` |
| Tray right-click | `ShowTrayContextMenu()` — popup menu with **Open** and **Exit** |
| Window close (X) | `MinimizeToTray()` — `ShowWindow(SW_HIDE)` (does NOT exit) |
| Menu ? Exit or Tray ? Exit | `DoExit()` — stops monitor thread, saves database, removes tray icon, destroys window |

### Data Flow Diagram

```
???????????????????????????????????????????????????????????????
?                     Windows Shell Data                       ?
?                                                             ?
?  ????????????????  ????????????????????  ????????????????? ?
?  ? Recent Folder ?  ? AutomaticDest/   ?  ? Quick Access / ? ?
?  ? (.lnk files)  ?  ? CustomDest files ?  ? Links folder   ? ?
?  ????????????????  ????????????????????  ????????????????? ?
???????????????????????????????????????????????????????????????
          ?                  ?                     ?
          ?                  ?                     ?
   GetRecentMRUPaths   GetJumplistPaths   GetExplorerFolder
          ?            ParseJumplistFile   JumplistPaths
          ?                  ?                     ?
          ??????????????????????????????????????????
                 ?
                 ?  ResolveShellLinkTarget()
          ????????????????     (GetPath + PIDL fallback
          ? Raw paths    ?      + ExpandEnvironmentStrings)
          ? (files/dirs) ?
          ????????????????
                 ?  Extract parent folder (if file)
                 ?
          ?????????????????
          ? IsExcludedPath ???? Skip Windows, Program Files,
          ? + IsDriveRoot  ?    AppData, drive roots
          ?????????????????
                 ?
                 ?
          LogFolderSource()  ???? source_log.txt
                 ?
                 ?
        AddFolderRecursive()
          ????????????????
          ?g_searchFolders???? std::set<wstring> (mutex-protected)
          ?  (normalized) ?
          ????????????????
                 ?
          ???????????????????????????
          ?                         ?
          ?                         ?
   PopulateTreeView()        SaveDatabase()
   ?????????????????    (binary file persistence)
   ? IsOneDrivePath ?
   ?  (checkbox)    ?
   ?????????????????
           ?
    (UI tree display)
```

### Libraries Used

| Library | Purpose |
|---|---|
| `Shell32.lib` | Shell namespace, `SHGetKnownFolderPath`, `IShellLink`, `IShellItem`, `IShellFolder`, `SHCreateMemStream` |
| `Ole32.lib` | COM runtime (`CoCreateInstance`, `CoInitializeEx`), Structured Storage (`StgOpenStorage`) |
| `Shlwapi.lib` | Shell lightweight utility functions |
| `Comctl32.lib` | Common Controls (TreeView via `WC_TREEVIEW`) |
| `Propsys.lib` | Property system utilities |
| `OleAut32.lib` | OLE Automation support |

### Build Requirements

- **Visual Studio 2022** (v143 platform toolset)
- **C++14** standard
- **Windows 10 SDK** (10.0 or later)
- **Unicode** character set
- Targets **Win32**, **x64**, and **ARM64** platforms
