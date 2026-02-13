#ifndef UNICODE
#define UNICODE
#endif

#ifndef _UNICODE
#define _UNICODE
#endif
#include "json.hpp"
#include <windows.h>
#include <initguid.h>
#include <shellapi.h>
#include <string>
#include <algorithm>
#include <vector>
#include <shlobj.h>
#include <commctrl.h>
#include <shlguid.h>
#include <tchar.h>
#include <gdiplus.h>
#include "json.hpp"
#include <fstream>
#include <tlhelp32.h>
#include <map>
#include <winuser.h>
#include <shellapi.h>
#include <dwmapi.h>
#include <psapi.h>
#include <powrprof.h>
#ifdef NO_JSON
void SaveDesktopPositions() {}
void LoadDesktopPositions() {}
void SaveStartMenuPositions() {}
void LoadStartMenuPositions() {}
void SaveTaskbarConfig() {}
void LoadTaskbarConfig() {}
void LoadWindowPositions() {}
#endif

#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "uuid.lib")
#pragma comment(lib, "gdiplus.lib")
#pragma comment(lib, "dwmapi.lib")
#pragma comment(lib, "PowrProf.lib")

HWND g_hDesktopWnd = NULL;
HWND g_hTaskbarWnd = NULL;
HWND g_hStartMenuWnd = NULL;
HINSTANCE g_hInst = NULL;
int g_taskbarHeight = 64;
bool g_taskbarVisible = true; // Track taskbar visibility state
bool g_taskbarVisibleBeforeWinKey = true; // Track taskbar state before Windows key press
bool g_startMenuOpenedByWinKey = false; // Track if start menu was opened by Windows key

// Keyboard hook for Windows key
static HHOOK g_keyboardHook = NULL;
static bool g_winKeyPressed = false;

// AppBar callback message
#define APPBAR_CALLBACK (WM_APP + 1)

// Custom colors for taskbar and start menu
static COLORREF g_taskbarColor = RGB(45, 45, 48); // Default dark gray
static COLORREF g_startMenuColor = RGB(50, 100, 160); // Default blue

#define IDM_ITEM_OPEN 40001
#define IDM_ITEM_PROPERTIES 40002
#define IDM_DESKTOP_REFRESH 40003

// New command IDs
#define IDM_VIEW_LARGE        50001
#define IDM_VIEW_MEDIUM       50002
#define IDM_VIEW_SMALL        50003

#define IDM_SORT_NAME         50011
#define IDM_SORT_TYPE         50012
#define IDM_SORT_DATE         50013

#define IDM_TOGGLE_SHOW_ICONS 50021
#define IDM_AUTO_ARRANGE      50022
#define IDM_ALIGN_TO_GRID     50023
#define IDM_CHANGE_BACKGROUND 50024
#define IDM_CHANGE_TASKBAR_COLOR 50025
#define IDM_CHANGE_STARTMENU_COLOR 50026

int g_rightClickSelection = -1;

// Icon/layout constants
const int ICON_SIZE_DEFAULT = 2;
const int ICON_X_SPACING_DEFAULT = 140;
const int ICON_Y_SPACING_DEFAULT = 120;
const int ICON_X_START = 50;
const int ICON_Y_START = 20;

// Current view state
int g_viewMode = 1; // Medium by default
int g_iconSize = ICON_SIZE_DEFAULT;
int g_iconXSpacing = ICON_X_SPACING_DEFAULT;
int g_iconYSpacing = ICON_Y_SPACING_DEFAULT;
bool g_showIcons = true;
bool g_autoArrange = false;

static WNDPROC g_origDesktopProc = NULL;
static WNDPROC g_realDesktopProc = NULL; // The actual DesktopWndProc address
static IContextMenu2* g_pcm2 = NULL;
static IContextMenu3* g_pcm3 = NULL;
static Gdiplus::GdiplusStartupInput gdiplusStartupInput;
static ULONG_PTR gdiplusToken;

// Desktop background image
static Gdiplus::Bitmap* g_backgroundImage = NULL;
static std::wstring g_backgroundPath;

struct DesktopItem {
    wchar_t name[MAX_PATH];
    wchar_t path[MAX_PATH];
    HICON hIcon;
    FILETIME ftWrite;
    std::wstring extension;
    int x; // pixel position for custom placement
    int y;
    bool hasPos; // whether x/y have been assigned
};
std::vector<DesktopItem> g_desktopItems;

// Structure for sidebar items with positions
struct SidebarItemInfo {
    bool isFolder;
    std::wstring name;
    std::wstring fullPath;
    int yPos;  // Actual screen position
    int height;
};
static std::vector<SidebarItemInfo> g_sidebarItems;

// Structure for running apps
struct RunningApp {
    DWORD processId;
    HWND hwnd;
    std::wstring exePath;
    std::wstring displayName;
    HICON hIcon;
    bool visible;
};
static std::vector<RunningApp> g_runningApps;

// For Windows 11 taskbar behavior
static HWND g_lastForegroundWindow = NULL;
static std::map<HWND, bool> g_windowMinimizedState;

// ==================== SYSTEM TRAY IMPROVEMENTS ====================

// Structure for REAL system tray icons (using Windows Shell API)
struct RealTrayIcon {
    NOTIFYICONDATAW nid;          // Original notify data
    std::wstring tooltip;         // Tooltip text
    HICON hIcon;                  // Current icon
    DWORD processId;              // Process ID
    HWND hWnd;                    // Owner window
    UINT uID;                     // Icon ID
    GUID guid;                    // GUID (for modern apps)
    bool hidden;                  // Hidden in overflow
    int trayIndex;                // Index in tray display
    RECT displayRect;             // Display rectangle
    bool isModernApp;             // UWP/Modern app
    std::wstring appName;         // App name
};

static std::vector<RealTrayIcon> g_realTrayIcons;
static HWND g_trayNotifyWnd = NULL;
static bool g_trayIconsInitialized = false;
static int g_trayIconSize = 20;  // Smaller size for tray icons
static int g_trayIconSpacing = 8;  // Increased from 4 to 8 for better spacing
static bool g_showTrayIcons = true;

// Structure for our drawn tray icons
struct TrayIcon {
    std::wstring name;
    HICON hIcon;
    int x;
    int y;
    int width;
    int height;
    bool isRealIcon;      // Is this a real system tray icon?
    UINT realIconId;      // Original icon ID if real
    HWND realOwner;       // Owner window if real
};
static std::vector<TrayIcon> g_trayIcons;

// Hidden icons tracking
static int g_chevronX = 0;
static int g_chevronY = 0;
static int g_chevronSize = 40;
static bool g_hasHiddenIcons = false;
static std::vector<int> g_hiddenIconIndices; // Indices of hidden real icons

// ==================== HELPER FUNCTIONS ====================

static int CompareNoCaseW(const wchar_t* a, const wchar_t* b) {
    return _wcsicmp(a, b);
}

std::wstring GetAppDataPath() {
    PWSTR pszPath = NULL;
    std::wstring result;
    if (SUCCEEDED(SHGetKnownFolderPath(FOLDERID_RoamingAppData, 0, NULL, &pszPath))) {
        result = pszPath;
        result += L"\\CustomShell";
        CoTaskMemFree(pszPath);
    }
    return result;
}

void SaveDesktopPositions() {
    try {
        nlohmann::json jPositions = nlohmann::json::array();
        for (const auto& item : g_desktopItems) {
            nlohmann::json jItem;
            char pathBuf[MAX_PATH];
            WideCharToMultiByte(CP_UTF8, 0, item.path, -1, pathBuf, sizeof(pathBuf), NULL, NULL);
            jItem["path"] = pathBuf;
            jItem["x"] = item.x;
            jItem["y"] = item.y;
            jPositions.push_back(jItem);
        }

        std::wstring appDataPath = GetAppDataPath();
        CreateDirectoryW(appDataPath.c_str(), NULL);
        std::wstring configPath = appDataPath + L"\\positions.json";
        std::ofstream file(configPath.c_str());
        if (file.is_open()) {
            file << jPositions.dump(2);
            file.close();
        }
    } catch (...) {}
}

void LoadDesktopPositions() {
    try {
        std::wstring appDataPath = GetAppDataPath();
        std::wstring configPath = appDataPath + L"\\positions.json";
        std::ifstream file(configPath.c_str());
        if (file.is_open()) {
            nlohmann::json jPositions;
            file >> jPositions;
            file.close();

            for (const auto& jItem : jPositions) {
                std::string pathStr = jItem["path"];
                wchar_t pathW[MAX_PATH];
                MultiByteToWideChar(CP_UTF8, 0, pathStr.c_str(), -1, pathW, MAX_PATH);

                for (auto& item : g_desktopItems) {
                    if (_wcsicmp(item.path, pathW) == 0) {
                        item.x = jItem["x"];
                        item.y = jItem["y"];
                        item.hasPos = true;
                        break;
                    }
                }
            }
        }
    } catch (...) {}
}

void DrawIconWithGDIPlus(Gdiplus::Graphics* g, HICON hIcon, int x, int y, int size) {
    if (!g || !hIcon) return;

    ICONINFO iconInfo;
    if (GetIconInfo(hIcon, &iconInfo)) {
        HDC hdc = g->GetHDC();
        if (hdc) {
            DrawIconEx(hdc, x, y, hIcon, size, size, 0, NULL, DI_NORMAL);
            g->ReleaseHDC(hdc);
        }
        if (iconInfo.hbmColor) DeleteObject(iconInfo.hbmColor);
        if (iconInfo.hbmMask) DeleteObject(iconInfo.hbmMask);
    }
}

static std::vector<std::wstring> g_pinnedApps; // Paths of pinned apps in start menu
static std::vector<std::wstring> g_removedApps; // Apps removed from start menu (don't show)
static std::wstring g_currentSidebarFolder = L"C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs"; // Current sidebar folder
static int g_sidebarScrollOffset = 0;

// Structure for sidebar items (mixed folders and files)
struct SidebarItem {
    bool isFolder;
    std::wstring name;
    std::wstring fullPath;
};

// Structure for taskbar icons
struct TaskbarIcon {
    std::wstring path;
    HICON hIcon;
    int x; // x position
    int y;
    int width;
    int height;
    bool visible; // whether to show on taskbar
};
static std::vector<TaskbarIcon> g_taskbarIcons;
static std::vector<std::wstring> g_taskbarPinned; // Apps pinned to taskbar (persisted)

// Drag state for moving icons (desktop)
static bool g_dragging = false;
static int g_dragIndex = -1;
static int g_dragOffsetX = 0;
static int g_dragOffsetY = 0;

// Drag state for Start menu tiles
static bool g_startMenuDragging = false;
static int g_startMenuDragIndex = -1;
static int g_startMenuDragOffsetX = 0;
static int g_startMenuDragOffsetY = 0;
static std::vector<int> g_tilePositionsX;  // Store tile X positions
static std::vector<int> g_tilePositionsY;  // Store tile Y positions
static int g_startMenuScrollOffset = 0;  // Scroll offset for main area
static int g_hoveredIcon = -1; // Track which icon is under cursor

// Forward declarations
void ResizeShellWindows(int dpi = 96);
void ShowShellContextMenu(HWND hwnd, POINT pt, LPCWSTR pszPath);
void ShowDesktopContextMenu(HWND hwnd, POINT pt);
void LaunchApp(LPCWSTR path);
void LaunchExplorer();
void UpdateViewMode(int mode);
void SortByName();
void SortByTypeThenName();
void SortByDateDesc();
void AlignIconsToGrid();
void PopulateDesktopIcons();
void ResetDesktopPositions();
void SaveDesktopPositions();
void LoadDesktopPositions();
void PopulateStartMenuApps();
void DrawIconWithGDIPlus(Gdiplus::Graphics* g, HICON hIcon, int x, int y, int size);
std::wstring GetAppDataPath();
void RegisterAppBar(HWND hwnd);
void SetAppBarPos(HWND hwnd);
void UnregisterAppBar(HWND hwnd);

// New functions for running apps
void UpdateRunningAppsList();
void AddCommonTrayIcons();
std::wstring GetAppExeName(LPCWSTR appPath);
void DestroyTaskbar();
void RecreateTaskbar();
void ResizeShellWindows(int dpi);

// ==================== SYSTEM TRAY FUNCTIONS ====================

// Find the real Windows tray notification window
HWND FindTrayNotifyWnd() {
    HWND hTrayWnd = FindWindowW(L"Shell_TrayWnd", NULL);
    if (!hTrayWnd) return NULL;
    
    // Look for TrayNotifyWnd inside Shell_TrayWnd
    HWND hTrayNotifyWnd = FindWindowExW(hTrayWnd, NULL, L"TrayNotifyWnd", NULL);
    if (!hTrayNotifyWnd) {
        // Try alternative names for newer Windows versions
        hTrayNotifyWnd = FindWindowExW(hTrayWnd, NULL, L"Windows.UI.Composition.DesktopWindowContentBridge", NULL);
    }
    return hTrayNotifyWnd;
}

// Hook into Windows Shell to get tray icons
BOOL CALLBACK EnumTrayWindowsProc(HWND hwnd, LPARAM lParam) {
    wchar_t className[256];
    GetClassNameW(hwnd, className, 256);
    
    // Look for tooltip class (tray icons send notifications to these)
    if (wcsstr(className, L"Tooltip") || wcsstr(className, L"NotifyIconOverflowWindow")) {
        // Found a potential tray icon container
        return TRUE;
    }
    return TRUE;
}

// Get process name from PID
std::wstring GetProcessName(DWORD processId) {
    HANDLE hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, processId);
    if (hProcess) {
        wchar_t exePath[MAX_PATH] = {0};
        DWORD pathSize = MAX_PATH;
        
        if (QueryFullProcessImageNameW(hProcess, 0, exePath, &pathSize)) {
            std::wstring pathStr = exePath;
            size_t lastSlash = pathStr.find_last_of(L"\\");
            if (lastSlash != std::wstring::npos) {
                return pathStr.substr(lastSlash + 1);
            }
            return pathStr;
        }
        CloseHandle(hProcess);
    }
    return L"Unknown";
}

// Get window text safely
std::wstring GetWindowTextSafe(HWND hwnd) {
    wchar_t buffer[256];
    int len = GetWindowTextW(hwnd, buffer, 256);
    if (len > 0) {
        return std::wstring(buffer);
    }
    return L"";
}

// Simulate getting tray icons by enumerating windows with tray-like properties
void PopulateRealTrayIcons() {
    // Clear old icons
    for (auto& icon : g_realTrayIcons) {
        if (icon.hIcon) {
            DestroyIcon(icon.hIcon);
        }
    }
    g_realTrayIcons.clear();
    
    // Clear our drawn tray icons
    for (auto& icon : g_trayIcons) {
        if (icon.hIcon && !icon.isRealIcon) {
            DestroyIcon(icon.hIcon);
        }
    }
    g_trayIcons.clear();
    
    // ADD SYSTEM ICONS FIRST so they always show
    AddCommonTrayIcons();
    
    // Get all top-level windows and look for tray-like behavior
    std::vector<HWND> potentialTrayWindows;
    
    EnumWindows([](HWND hwnd, LPARAM lParam) -> BOOL {
        auto& windows = *reinterpret_cast<std::vector<HWND>*>(lParam);
        
        if (!IsWindowVisible(hwnd)) return TRUE;
        
        // Get window class
        wchar_t className[256];
        GetClassNameW(hwnd, className, 256);
        
        // Look for windows that might have tray icons
        // Many apps create hidden windows for tray notifications
        DWORD processId;
        GetWindowThreadProcessId(hwnd, &processId);
        
        // Check window style and extended style
        LONG style = GetWindowLongW(hwnd, GWL_STYLE);
        LONG exStyle = GetWindowLongW(hwnd, GWL_EXSTYLE);
        
        // Hidden windows with tooltip styles might be tray owners
        if ((style & WS_POPUP) && (exStyle & WS_EX_TOOLWINDOW)) {
            windows.push_back(hwnd);
        }
        
        return TRUE;
    }, reinterpret_cast<LPARAM>(&potentialTrayWindows));
    
    // Now try to communicate with these windows to get tray data
    // This is a simplified approach - real implementation would use Shell Hook
    for (HWND hwnd : potentialTrayWindows) {
        DWORD processId;
        GetWindowThreadProcessId(hwnd, &processId);
        
        // Try to get the window's icon
        HICON hIcon = (HICON)SendMessageW(hwnd, WM_GETICON, ICON_SMALL, 0);
        if (!hIcon) {
            hIcon = (HICON)SendMessageW(hwnd, WM_GETICON, ICON_BIG, 0);
        }
        if (!hIcon) {
            hIcon = (HICON)GetClassLongPtrW(hwnd, GCLP_HICONSM);
        }
        
        if (hIcon) {
            RealTrayIcon trayIcon = {};
            trayIcon.hWnd = hwnd;
            trayIcon.processId = processId;
            trayIcon.hIcon = CopyIcon(hIcon); // Copy the icon
            trayIcon.tooltip = GetWindowTextSafe(hwnd);
            trayIcon.appName = GetProcessName(processId);
            trayIcon.hidden = false;
            
            g_realTrayIcons.push_back(trayIcon);
        }
    }
    
    // NOW add the detected real tray icons AFTER system icons
    for (size_t i = 0; i < g_realTrayIcons.size(); ++i) {
        TrayIcon ti;
        ti.name = g_realTrayIcons[i].appName;
        ti.hIcon = CopyIcon(g_realTrayIcons[i].hIcon);
        ti.isRealIcon = true;
        ti.realOwner = g_realTrayIcons[i].hWnd;
        ti.width = g_trayIconSize;
        ti.height = g_trayIconSize;
        g_trayIcons.push_back(ti);
    }
}

// Add common system tray icons that are always present
void AddCommonTrayIcons() {
    // Volume icon
    {
        TrayIcon ti;
        ti.name = L"Volume";
        ti.isRealIcon = false;
        
        // Try to get the actual volume icon from Windows
        HMODULE hShell32 = LoadLibraryW(L"shell32.dll");
        if (hShell32) {
            ti.hIcon = (HICON)LoadImageW(hShell32, MAKEINTRESOURCE(168), IMAGE_ICON, 
                                        g_trayIconSize, g_trayIconSize, LR_SHARED);
            FreeLibrary(hShell32);
        }
        if (!ti.hIcon) {
            ti.hIcon = NULL; // Will draw colored rect
        }
        ti.width = g_trayIconSize;
        ti.height = g_trayIconSize;
        g_trayIcons.push_back(ti);
    }
    
    // Network icon
    {
        TrayIcon ti;
        ti.name = L"Network";
        ti.isRealIcon = false;
        
        HMODULE hShell32 = LoadLibraryW(L"shell32.dll");
        if (hShell32) {
            ti.hIcon = (HICON)LoadImageW(hShell32, MAKEINTRESOURCE(17), IMAGE_ICON, 
                                        g_trayIconSize, g_trayIconSize, LR_SHARED);
            FreeLibrary(hShell32);
        }
        if (!ti.hIcon) {
            ti.hIcon = NULL;
        }
        ti.width = g_trayIconSize;
        ti.height = g_trayIconSize;
        g_trayIcons.push_back(ti);
    }
    
    // Power/battery
    {
        TrayIcon ti;
        ti.name = L"Power";
        ti.isRealIcon = false;
        
        HMODULE hShell32 = LoadLibraryW(L"shell32.dll");
        if (hShell32) {
            ti.hIcon = (HICON)LoadImageW(hShell32, MAKEINTRESOURCE(244), IMAGE_ICON, 
                                        g_trayIconSize, g_trayIconSize, LR_SHARED);
            FreeLibrary(hShell32);
        }
        if (!ti.hIcon) {
            ti.hIcon = NULL;
        }
        ti.width = g_trayIconSize;
        ti.height = g_trayIconSize;
        g_trayIcons.push_back(ti);
    }
    
    // Clock (last, on the right)
    {
        TrayIcon ti;
        ti.name = L"Clock";
        ti.isRealIcon = false;
        ti.hIcon = NULL; // Will draw text
        ti.width = 55; // Wider for bigger time display (increased from 40)
        ti.height = g_trayIconSize;
        g_trayIcons.push_back(ti);
    }
}

// Draw time in tray
void DrawTrayTime(Gdiplus::Graphics* g, int x, int y, int width, int height) {
    SYSTEMTIME st;
    GetLocalTime(&st);
    
    wchar_t timeStr[64];
    wsprintfW(timeStr, L"%02d:%02d", st.wHour, st.wMinute);
    
    Gdiplus::FontFamily fontFamily(L"Segoe UI");
    Gdiplus::Font timeFont(&fontFamily, 13, Gdiplus::FontStyleRegular, Gdiplus::UnitPixel); // Increased from 11 to 13
    Gdiplus::SolidBrush textBrush(Gdiplus::Color(255, 240, 240, 240));
    
    Gdiplus::RectF textRect(x, y, width, height);
    Gdiplus::StringFormat format;
    format.SetAlignment(Gdiplus::StringAlignmentCenter);
    format.SetLineAlignment(Gdiplus::StringAlignmentCenter);
    
    g->DrawString(timeStr, -1, &timeFont, textRect, &format, &textBrush);
}

// Handle tray icon clicks
void HandleTrayIconClick(int index, bool rightClick) {
    if (index < 0 || index >= (int)g_trayIcons.size()) return;
    
    TrayIcon& icon = g_trayIcons[index];
    
    if (icon.isRealIcon && icon.realOwner) {
        // Send message to the real tray icon owner
        if (rightClick) {
            // Right click - show context menu
            SendMessageW(icon.realOwner, WM_CONTEXTMENU, (WPARAM)icon.realOwner, 
                        MAKELPARAM(icon.x + icon.width/2, icon.y + icon.height/2));
        } else {
            // Left click - simulate normal tray icon click
            SendMessageW(icon.realOwner, WM_USER + 1, 0, MAKELPARAM(WM_LBUTTONDOWN, 0));
            SendMessageW(icon.realOwner, WM_USER + 1, 0, MAKELPARAM(WM_LBUTTONUP, 0));
        }
    } else {
        // Handle our built-in icons
        if (icon.name == L"Volume") {
            if (rightClick) {
                // Open sound control panel
                ShellExecuteW(NULL, L"open", L"control.exe", L"mmsys.cpl", NULL, SW_SHOW);
            } else {
                // Open legacy volume control that works without explorer
                ShellExecuteW(NULL, L"open", L"sndvol.exe", NULL, NULL, SW_SHOW);
            }
        } else if (icon.name == L"Network") {
            // Open network connections control panel
            ShellExecuteW(NULL, L"open", L"control.exe", L"ncpa.cpl", NULL, SW_SHOW);
        } else if (icon.name == L"Power") {
            // Open power options control panel
            ShellExecuteW(NULL, L"open", L"control.exe", L"powercfg.cpl", NULL, SW_SHOW);
        } else if (icon.name == L"Clock") {
            // Open Date & Time control panel
            ShellExecuteW(NULL, L"open", L"control.exe", L"timedate.cpl", NULL, SW_SHOW);
        }
    }
}

// ==================== WINDOWS 11 MINIMIZE/RESTORE ====================

void FlashWindowIfMinimized(HWND hwnd) {
    if (IsWindow(hwnd) && IsIconic(hwnd)) {
        FLASHWINFO fwi;
        fwi.cbSize = sizeof(FLASHWINFO);
        fwi.hwnd = hwnd;
        fwi.dwFlags = FLASHW_STOP;
        fwi.uCount = 0;
        fwi.dwTimeout = 0;
        FlashWindowEx(&fwi);

        fwi.dwFlags = FLASHW_TRAY | FLASHW_TIMERNOFG;
        fwi.uCount = 3;
        fwi.dwTimeout = 0;
        FlashWindowEx(&fwi);
    }
}

// MODIFIED: Added Start Menu position saving
void SaveStartMenuPositions() {
    try {
        nlohmann::json jPositions = nlohmann::json::object();
        jPositions["pinnedApps"] = nlohmann::json::array();

        for (const auto& app : g_pinnedApps) {
            char pathBuf[MAX_PATH];
            WideCharToMultiByte(CP_UTF8, 0, app.c_str(), -1, pathBuf, sizeof(pathBuf), NULL, NULL);
            jPositions["pinnedApps"].push_back(pathBuf);
        }

        std::wstring appDataPath = GetAppDataPath();
        CreateDirectoryW(appDataPath.c_str(), NULL);
        std::wstring configPath = appDataPath + L"\\startmenu_positions.json";
        std::ofstream file(configPath.c_str());
        if (file.is_open()) {
            file << jPositions.dump(2);
            file.close();
        }
    } catch (...) {}
}

void LoadStartMenuPositions() {
    try {
        std::wstring appDataPath = GetAppDataPath();
        std::wstring configPath = appDataPath + L"\\startmenu_positions.json";
        std::ifstream file(configPath.c_str());
        if (file.is_open()) {
            nlohmann::json jPositions;
            file >> jPositions;
            file.close();

            if (jPositions.contains("pinnedApps") && jPositions["pinnedApps"].is_array()) {
                g_pinnedApps.clear();
                for (const auto& item : jPositions["pinnedApps"]) {
                    std::string pathStr = item;
                    wchar_t pathW[MAX_PATH];
                    MultiByteToWideChar(CP_UTF8, 0, pathStr.c_str(), -1, pathW, MAX_PATH);
                    g_pinnedApps.push_back(pathW);
                }
            }
        }
    } catch (...) {}
}

// NEW FUNCTION: Clean up explorer and shell from all lists
void CleanupExplorerAndShellFromLists() {
    // Clean up from start menu
    g_pinnedApps.erase(
        std::remove_if(g_pinnedApps.begin(), g_pinnedApps.end(),
            [](const std::wstring& appPath) {
                std::wstring fileName = appPath.substr(appPath.find_last_of(L"\\") + 1);
                std::wstring lowerName = fileName;
                std::transform(lowerName.begin(), lowerName.end(), lowerName.begin(), ::towlower);

                return (lowerName.find(L"explorer") != std::wstring::npos) ||
                       (lowerName.find(L"shell") != std::wstring::npos) ||
                       (lowerName.find(L"customshell") != std::wstring::npos);
            }),
        g_pinnedApps.end()
    );

    // Clean up from taskbar
    g_taskbarPinned.erase(
        std::remove_if(g_taskbarPinned.begin(), g_taskbarPinned.end(),
            [](const std::wstring& appPath) {
                std::wstring fileName = appPath.substr(appPath.find_last_of(L"\\") + 1);
                std::wstring lowerName = fileName;
                std::transform(lowerName.begin(), lowerName.end(), lowerName.begin(), ::towlower);

                return (lowerName.find(L"explorer") != std::wstring::npos) ||
                       (lowerName.find(L"shell") != std::wstring::npos) ||
                       (lowerName.find(L"customshell") != std::wstring::npos);
            }),
        g_taskbarPinned.end()
    );

    // Clean up from taskbar icons
    for (auto it = g_taskbarIcons.begin(); it != g_taskbarIcons.end(); ) {
        std::wstring fileName = it->path.substr(it->path.find_last_of(L"\\") + 1);
        std::wstring lowerName = fileName;
        std::transform(lowerName.begin(), lowerName.end(), lowerName.begin(), ::towlower);

        if (lowerName.find(L"explorer") != std::wstring::npos ||
            lowerName.find(L"shell") != std::wstring::npos ||
            lowerName.find(L"customshell") != std::wstring::npos) {
            if (it->hIcon) DestroyIcon(it->hIcon);
            it = g_taskbarIcons.erase(it);
        } else {
            ++it;
        }
    }
}

// MODIFIED: Updated PopulateStartMenuApps to filter explorer and shell
void PopulateStartMenuApps() {
    // First try to load saved positions
    LoadStartMenuPositions();

    // FILTER OUT EXPLORER AND SHELL FROM SAVED POSITIONS TOO
    CleanupExplorerAndShellFromLists();

    // If no saved positions, load default
    if (g_pinnedApps.empty()) {
        std::wstring startMenuPath = L"C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs";

        // Scan for .lnk files
        WIN32_FIND_DATAW findData;
        HANDLE hFind = FindFirstFileW((startMenuPath + L"\\*.lnk").c_str(), &findData);

        if (hFind != INVALID_HANDLE_VALUE) {
            do {
                if (!(findData.dwFileAttributes & FILE_ATTRIBUTE_HIDDEN)) {
                    std::wstring fullPath = startMenuPath + L"\\" + findData.cFileName;
                    std::wstring fileName = findData.cFileName;

                    // Convert to lowercase for case-insensitive comparison
                    std::wstring lowerName = fileName;
                    std::transform(lowerName.begin(), lowerName.end(), lowerName.begin(), ::towlower);

                    // FILTER OUT EXPLORER AND SHELL COMPLETELY
                    bool isExplorer =
                        lowerName.find(L"explorer") != std::wstring::npos ||
                        lowerName.find(L"file explorer") != std::wstring::npos ||
                        lowerName.find(L"windows explorer") != std::wstring::npos;

                    bool isThisShell =
                        lowerName.find(L"shell.exe") != std::wstring::npos ||
                        lowerName.find(L"shell.lnk") != std::wstring::npos ||
                        lowerName.find(L"customshell") != std::wstring::npos;

                    if (!isExplorer && !isThisShell) {
                        g_pinnedApps.push_back(fullPath);
                    }
                }
            } while (FindNextFileW(hFind, &findData));
            FindClose(hFind);
        }

        // Save filtered positions
        SaveStartMenuPositions();
    }
}

void SaveTaskbarConfig() {
    try {
        nlohmann::json jConfig = nlohmann::json::object();

        // Save pinned apps array
        nlohmann::json jApps = nlohmann::json::array();
        for (const auto& app : g_taskbarPinned) {
            jApps.push_back(app);
        }
        jConfig["apps"] = jApps;

        // Save window positions
        nlohmann::json jWindows = nlohmann::json::object();

        if (g_hTaskbarWnd) {
            RECT rect;
            GetWindowRect(g_hTaskbarWnd, &rect);
            jWindows["taskbar"] = {
                {"x", rect.left},
                {"y", rect.top},
                {"width", rect.right - rect.left},
                {"height", rect.bottom - rect.top}
            };
        }

        if (g_hStartMenuWnd) {
            RECT rect;
            GetWindowRect(g_hStartMenuWnd, &rect);
            jWindows["startmenu"] = {
                {"x", rect.left},
                {"y", rect.top},
                {"width", rect.right - rect.left},
                {"height", rect.bottom - rect.top}
            };
        }

        jConfig["windows"] = jWindows;
        
        // Save background image path
        if (!g_backgroundPath.empty()) {
            char bgPath[MAX_PATH];
            WideCharToMultiByte(CP_UTF8, 0, g_backgroundPath.c_str(), -1, bgPath, sizeof(bgPath), NULL, NULL);
            jConfig["background"] = bgPath;
        }
        
        // Save custom colors
        jConfig["taskbarColor"] = {
            {"r", GetRValue(g_taskbarColor)},
            {"g", GetGValue(g_taskbarColor)},
            {"b", GetBValue(g_taskbarColor)}
        };
        jConfig["startMenuColor"] = {
            {"r", GetRValue(g_startMenuColor)},
            {"g", GetGValue(g_startMenuColor)},
            {"b", GetBValue(g_startMenuColor)}
        };

        std::wstring appDataPath = GetAppDataPath();
        CreateDirectoryW(appDataPath.c_str(), NULL);

        char filePath[MAX_PATH];
        WideCharToMultiByte(CP_UTF8, 0, (appDataPath + L"\\shell_config.json").c_str(), -1, filePath, sizeof(filePath), NULL, NULL);

        std::ofstream file(filePath);
        file << jConfig.dump(2);
        file.close();
    } catch (...) {
        // Silently fail if can't save
    }
}

void LoadTaskbarConfig() {
    try {
        std::wstring appDataPath = GetAppDataPath();

        char filePath[MAX_PATH];
        WideCharToMultiByte(CP_UTF8, 0, (appDataPath + L"\\shell_config.json").c_str(), -1, filePath, sizeof(filePath), NULL, NULL);

        std::ifstream file(filePath);
        if (file.is_open()) {
            nlohmann::json jConfig;
            file >> jConfig;
            file.close();

            // Load pinned apps
            if (jConfig.contains("apps") && jConfig["apps"].is_array()) {
                g_taskbarPinned.clear();
                for (const auto& item : jConfig["apps"]) {
                    g_taskbarPinned.push_back(item.template get<std::wstring>());
                }
            }
        }
    } catch (...) {
        // Use defaults if can't load
    }
}

int g_taskbarX = 0, g_taskbarY = 0, g_taskbarWidth = 0;
int g_startMenuX = 10, g_startMenuY = 0, g_startMenuWidth = 0, g_startMenuHeightCfg = 0;

void LoadWindowPositions() {
    try {
        std::wstring appDataPath = GetAppDataPath();

        char filePath[MAX_PATH];
        WideCharToMultiByte(CP_UTF8, 0, (appDataPath + L"\\shell_config.json").c_str(), -1, filePath, sizeof(filePath), NULL, NULL);

        std::ifstream file(filePath);
        if (file.is_open()) {
            nlohmann::json jConfig;
            file >> jConfig;
            file.close();

            if (jConfig.contains("windows") && jConfig["windows"].is_object()) {
                auto& jWnd = jConfig["windows"];

                if (jWnd.contains("taskbar")) {
                    auto& tb = jWnd["taskbar"];
                    g_taskbarX = tb.value("x", 0);
                    g_taskbarY = tb.value("y", 0);
                    g_taskbarWidth = tb.value("width", 0);
                }

                if (jWnd.contains("startmenu")) {
                    auto& sm = jWnd["startmenu"];
                    g_startMenuX = sm.value("x", 10);
                    g_startMenuY = sm.value("y", 0);
                    g_startMenuWidth = sm.value("width", 0);
                    g_startMenuHeightCfg = sm.value("height", 0);
                }
            }
            
            // Load background image path
            if (jConfig.contains("background") && jConfig["background"].is_string()) {
                std::string bgPath = jConfig["background"].get<std::string>();
                if (!bgPath.empty()) {
                    wchar_t wPath[MAX_PATH];
                    MultiByteToWideChar(CP_UTF8, 0, bgPath.c_str(), -1, wPath, MAX_PATH);
                    g_backgroundPath = wPath;
                    
                    // Load the image
                    if (g_backgroundImage) {
                        delete g_backgroundImage;
                        g_backgroundImage = NULL;
                    }
                    g_backgroundImage = Gdiplus::Bitmap::FromFile(g_backgroundPath.c_str());
                    if (g_backgroundImage && g_backgroundImage->GetLastStatus() != Gdiplus::Ok) {
                        delete g_backgroundImage;
                        g_backgroundImage = NULL;
                    }
                }
            }
            
            // Load custom colors
            if (jConfig.contains("taskbarColor") && jConfig["taskbarColor"].is_object()) {
                auto& tc = jConfig["taskbarColor"];
                int r = tc.value("r", 45);
                int g = tc.value("g", 45);
                int b = tc.value("b", 48);
                g_taskbarColor = RGB(r, g, b);
            }
            if (jConfig.contains("startMenuColor") && jConfig["startMenuColor"].is_object()) {
                auto& sc = jConfig["startMenuColor"];
                int r = sc.value("r", 50);
                int g = sc.value("g", 100);
                int b = sc.value("b", 160);
                g_startMenuColor = RGB(r, g, b);
            }
        }
    } catch (...) {
        // Use defaults
    }
}

// MODIFIED: Updated to filter out explorer and shell from taskbar
void PopulateTaskbarIcons() {
    // Clear existing taskbar icons
    for (auto& icon : g_taskbarIcons) {
        if (icon.hIcon) {
            DestroyIcon(icon.hIcon);
        }
    }
    g_taskbarIcons.clear();

    // Clean up taskbar pinned list first
    CleanupExplorerAndShellFromLists();

    // If no pinned apps, use from start menu but skip explorer and shell
    if (g_taskbarPinned.empty()) {
        for (size_t i = 0; i < g_pinnedApps.size() && i < 6; ++i) {
            std::wstring appPath = g_pinnedApps[i];
            std::wstring fileName = appPath.substr(appPath.find_last_of(L"\\") + 1);
            std::wstring lowerName = fileName;
            std::transform(lowerName.begin(), lowerName.end(), lowerName.begin(), ::towlower);

            // SKIP EXPLORER AND SHELL
            if (lowerName.find(L"explorer") == std::wstring::npos &&
                lowerName.find(L"shell.exe") == std::wstring::npos &&
                lowerName.find(L"shell.lnk") == std::wstring::npos &&
                lowerName.find(L"customshell") == std::wstring::npos) {
                g_taskbarPinned.push_back(appPath);
            }
        }
        SaveTaskbarConfig();
    }

    // Load icons from pinned apps (excluding explorer and shell)
    for (const auto& appPath : g_taskbarPinned) {
        std::wstring fileName = appPath.substr(appPath.find_last_of(L"\\") + 1);
        std::wstring lowerName = fileName;
        std::transform(lowerName.begin(), lowerName.end(), lowerName.begin(), ::towlower);

        // DOUBLE CHECK: Don't add explorer or shell
        if (lowerName.find(L"explorer") == std::wstring::npos &&
            lowerName.find(L"shell.exe") == std::wstring::npos &&
            lowerName.find(L"shell.lnk") == std::wstring::npos &&
            lowerName.find(L"customshell") == std::wstring::npos) {
            TaskbarIcon tIcon;
            tIcon.path = appPath;
            tIcon.hIcon = NULL;
            tIcon.visible = true;

            // Extract icon from .lnk file
            SHFILEINFOW sfi = {};
            if (SHGetFileInfoW(appPath.c_str(), 0, &sfi, sizeof(sfi), SHGFI_ICON | SHGFI_LARGEICON)) {
                tIcon.hIcon = sfi.hIcon;
            }

            g_taskbarIcons.push_back(tIcon);
        }
    }
}

void DrawPolishedIconBackground(Gdiplus::Graphics* g, int x, int y, int iconSize, int spacing, bool hovered) {
    using namespace Gdiplus;

    int bgX = x - (spacing - iconSize) / 2;
    int bgY = y;
    int bgWidth = spacing;
    int bgHeight = spacing;

    // Draw shadow with rounded effect
    SolidBrush shadowBrush(Color(40, 0, 0, 0));
    RectF shadowRect(bgX + 2, bgY + 2, bgWidth, bgHeight);
    g->FillRectangle(&shadowBrush, shadowRect);

    // Draw background with rounded corners using a path
    Color bgColor;
    if (hovered) {
        bgColor = Color(180, 255, 255, 255);
    } else {
        bgColor = Color(140, 255, 255, 255);
    }
    SolidBrush bgBrush(bgColor);

    GraphicsPath path;
    float radius = 12.0f;
    RectF rect(bgX, bgY, bgWidth, bgHeight);
    path.AddArc(rect.X, rect.Y, radius, radius, 180, 90);
    path.AddArc(rect.X + rect.Width - radius, rect.Y, radius, radius, 270, 90);
    path.AddArc(rect.X + rect.Width - radius, rect.Y + rect.Height - radius, radius, radius, 0, 90);
    path.AddArc(rect.X, rect.Y + rect.Height - radius, radius, radius, 90, 90);
    path.CloseFigure();

    g->FillPath(&bgBrush, &path);

    Pen borderPen(Color(200, 200, 200, 200), 1.0f);
    g->DrawPath(&borderPen, &path);
}

LRESULT CALLBACK Desktop_SubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    if (g_pcm3) {
        LRESULT lr = 0;
        if (SUCCEEDED(g_pcm3->HandleMenuMsg2(msg, wParam, lParam, &lr))) {
            return lr;
        }
    }
    if (g_pcm2) {
        if (SUCCEEDED(g_pcm2->HandleMenuMsg(msg, wParam, lParam))) {
            return 0;
        }
    }
    return CallWindowProcW(g_origDesktopProc ? g_origDesktopProc : DefWindowProcW, hwnd, msg, wParam, lParam);
}

std::wstring GetAppExeName(LPCWSTR appPath) {
    std::wstring pathStr(appPath);
    size_t lastSlash = pathStr.find_last_of(L"\\");
    std::wstring fileName = (lastSlash != std::wstring::npos) ?
        pathStr.substr(lastSlash + 1) : pathStr;

    size_t dotPos = fileName.find(L".lnk");
    if (dotPos != std::wstring::npos) {
        fileName = fileName.substr(0, dotPos);
    }

    dotPos = fileName.find(L".exe");
    if (dotPos != std::wstring::npos) {
        fileName = fileName.substr(0, dotPos);
    }

    return fileName;
}

DWORD FindProcessByName(const std::wstring& appName) {
    HANDLE hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (hSnapshot == INVALID_HANDLE_VALUE) return 0;

    PROCESSENTRY32W pe32 = {};
    pe32.dwSize = sizeof(PROCESSENTRY32W);

    DWORD processId = 0;
    if (Process32FirstW(hSnapshot, &pe32)) {
        do {
            std::wstring exeName(pe32.szExeFile);
            if (_wcsicmp(exeName.c_str(), (appName + L".exe").c_str()) == 0 ||
                _wcsicmp(exeName.c_str(), appName.c_str()) == 0) {
                processId = pe32.th32ProcessID;
                break;
            }
        } while (Process32NextW(hSnapshot, &pe32));
    }

    CloseHandle(hSnapshot);
    return processId;
}

HWND FindWindowByProcessId(DWORD processId) {
    HWND hwnd = NULL;
    do {
        hwnd = FindWindowExW(NULL, hwnd, NULL, NULL);
        if (hwnd) {
            DWORD windowPID = 0;
            GetWindowThreadProcessId(hwnd, &windowPID);
            if (windowPID == processId) {
                if (IsWindowVisible(hwnd)) {
                    return hwnd;
                }
            }
        }
    } while (hwnd);

    return NULL;
}

void UpdateRunningAppsList() {
    static std::map<HWND, bool> prevMinimizedState;

    for (auto it = g_runningApps.begin(); it != g_runningApps.end();) {
        if (!IsWindow(it->hwnd) || !IsWindowVisible(it->hwnd)) {
            if (it->hIcon) DestroyIcon(it->hIcon);
            prevMinimizedState.erase(it->hwnd);
            it = g_runningApps.erase(it);
        } else {
            bool isMinimized = IsIconic(it->hwnd);
            if (prevMinimizedState[it->hwnd] != isMinimized) {
                prevMinimizedState[it->hwnd] = isMinimized;
                g_windowMinimizedState[it->hwnd] = isMinimized;
                InvalidateRect(g_hTaskbarWnd, NULL, FALSE);
            }
            ++it;
        }
    }

    EnumWindows([](HWND hwnd, LPARAM lParam) -> BOOL {
        if (!IsWindowVisible(hwnd)) return TRUE;

        DWORD processId;
        GetWindowThreadProcessId(hwnd, &processId);

        if (hwnd == g_hDesktopWnd || hwnd == g_hTaskbarWnd || hwnd == g_hStartMenuWnd) {
            return TRUE;
        }

        wchar_t title[256];
        GetWindowTextW(hwnd, title, 256);
        if (wcslen(title) == 0) return TRUE;

        bool found = false;
        for (auto& app : g_runningApps) {
            if (app.hwnd == hwnd) {
                found = true;
                break;
            }
        }

        if (!found) {
            HANDLE hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, processId);
            if (hProcess) {
                wchar_t exePath[MAX_PATH] = {0};
                DWORD pathSize = MAX_PATH;

                if (QueryFullProcessImageNameW(hProcess, 0, exePath, &pathSize)) {
                    RunningApp app;
                    app.processId = processId;
                    app.hwnd = hwnd;
                    app.exePath = exePath;
                    app.visible = true;
                    app.displayName = title;

                    SHFILEINFOW sfi = {};
                    if (SHGetFileInfoW(exePath, 0, &sfi, sizeof(sfi), SHGFI_ICON | SHGFI_SMALLICON)) {
                        app.hIcon = sfi.hIcon;
                    } else {
                        app.hIcon = NULL;
                    }

                    g_runningApps.push_back(app);
                    g_windowMinimizedState[hwnd] = IsIconic(hwnd);
                }
                CloseHandle(hProcess);
            }
        }

        return TRUE;
    }, 0);
}

// MODIFIED: Block explorer and shell completely
void LaunchApp(LPCWSTR path) {
    std::wstring appPath = path;
    std::wstring fileName = appPath.substr(appPath.find_last_of(L"\\") + 1);
    std::wstring lowerName = fileName;
    std::transform(lowerName.begin(), lowerName.end(), lowerName.begin(), ::towlower);

    // BLOCK EXPLORER AND SHELL COMPLETELY
    if (lowerName.find(L"explorer") != std::wstring::npos ||
        lowerName.find(L"shell.exe") != std::wstring::npos ||
        lowerName.find(L"shell.lnk") != std::wstring::npos ||
        lowerName.find(L"customshell") != std::wstring::npos) {
        return;
    }

    std::wstring targetPath = path;

    if (targetPath.find(L".lnk") != std::wstring::npos) {
        IShellLinkW* psl = NULL;
        IPersistFile* ppf = NULL;
        wchar_t resolvedPath[MAX_PATH] = {0};

        if (SUCCEEDED(CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER,
                                      IID_IShellLinkW, (void**)&psl))) {
            if (SUCCEEDED(psl->QueryInterface(IID_IPersistFile, (void**)&ppf))) {
                if (SUCCEEDED(ppf->Load(path, STGM_READ))) {
                    psl->GetPath(resolvedPath, MAX_PATH, NULL, 0);
                    targetPath = resolvedPath;
                }
                ppf->Release();
            }
            psl->Release();
        }
    }

    for (const auto& app : g_runningApps) {
        if (_wcsicmp(app.exePath.c_str(), targetPath.c_str()) == 0) {
            if (IsIconic(app.hwnd)) {
                ShowWindow(app.hwnd, SW_RESTORE);
            }
            SetForegroundWindow(app.hwnd);
            SetWindowPos(app.hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
            return;
        }
    }

    SHELLEXECUTEINFOW sei = {0};
    sei.cbSize = sizeof(sei);
    sei.fMask = SEE_MASK_NOCLOSEPROCESS;
    sei.lpVerb = L"open";
    sei.lpFile = path;
    sei.nShow = SW_SHOWNORMAL;

    if (ShellExecuteExW(&sei)) {
        if (sei.hProcess) {
            CloseHandle(sei.hProcess);
        }
    }
}

bool IsAppRunning(LPCWSTR appPath) {
    std::wstring appName = GetAppExeName(appPath);
    return FindProcessByName(appName) != 0;
}

void LaunchExplorer() {
    ShellExecuteW(NULL, L"open", L"explorer.exe", NULL, NULL, SW_SHOWDEFAULT);
}

void UpdateViewMode(int mode) {
    g_viewMode = mode;
    switch (mode) {
    case 0:
        g_iconSize = 32;
        g_iconXSpacing = 100;
        g_iconYSpacing = 80;
        break;
    case 1:
        g_iconSize = 48;
        g_iconXSpacing = 120;
        g_iconYSpacing = 100;
        break;
    case 2:
    default:
        g_iconSize = 64;
        g_iconXSpacing = 140;
        g_iconYSpacing = 120;
        break;
    }
    InvalidateRect(g_hDesktopWnd, NULL, TRUE);
}

void SortByName() {
    std::sort(g_desktopItems.begin(), g_desktopItems.end(),
        [](const DesktopItem& a, const DesktopItem& b) {
            return _wcsicmp(a.name, b.name) < 0;
        });
}

void SortByTypeThenName() {
    std::sort(g_desktopItems.begin(), g_desktopItems.end(),
        [](const DesktopItem& a, const DesktopItem& b) {
            int extCmp = _wcsicmp(a.extension.c_str(), b.extension.c_str());
            if (extCmp != 0) return extCmp < 0;
            return _wcsicmp(a.name, b.name) < 0;
        });
}

void SortByDateDesc() {
    std::sort(g_desktopItems.begin(), g_desktopItems.end(),
        [](const DesktopItem& a, const DesktopItem& b) {
            return CompareFileTime(&a.ftWrite, &b.ftWrite) > 0;
        });
}

void ResetDesktopPositions() {
    SortByName();
    int screenW = GetSystemMetrics(SM_CXSCREEN);
    int curX = ICON_X_START;
    int curY = ICON_Y_START;
    for (auto &it : g_desktopItems) {
        it.x = curX;
        it.y = curY;
        it.hasPos = true;

        curX += g_iconXSpacing;
        if (curX + g_iconXSpacing > screenW) {
            curX = ICON_X_START;
            curY += g_iconYSpacing;
        }
    }
    InvalidateRect(g_hDesktopWnd, NULL, TRUE);
}

void AlignIconsToGrid() {
    InvalidateRect(g_hDesktopWnd, NULL, TRUE);
}

void PopulateDesktopIcons() {
    std::vector<std::wstring> desktopPaths;

    PWSTR pszPathUser = NULL;
    if (S_OK == SHGetKnownFolderPath(FOLDERID_Desktop, 0, NULL, &pszPathUser)) {
        desktopPaths.push_back(pszPathUser);
        CoTaskMemFree(pszPathUser);
    }

    PWSTR pszPathCommon = NULL;
    if (S_OK == SHGetKnownFolderPath(FOLDERID_PublicDesktop, 0, NULL, &pszPathCommon)) {
        desktopPaths.push_back(pszPathCommon);
        CoTaskMemFree(pszPathCommon);
    }

    for (auto &it : g_desktopItems) {
        if (it.hIcon) {
            DestroyIcon(it.hIcon);
            it.hIcon = NULL;
        }
    }
    g_desktopItems.clear();

    for (const auto& desktopDir : desktopPaths) {
        std::wstring searchPath = desktopDir + L"\\*";
        WIN32_FIND_DATAW findData;
        HANDLE hFind = FindFirstFileW(searchPath.c_str(), &findData);
        if (hFind != INVALID_HANDLE_VALUE) {
            do {
                if (!(findData.dwFileAttributes & FILE_ATTRIBUTE_HIDDEN) &&
                    wcscmp(findData.cFileName, L".") != 0 &&
                    wcscmp(findData.cFileName, L"..") != 0)
                {
                    DesktopItem item = {};
                    SHFILEINFOW sfi = {};
                    std::wstring fullPath = desktopDir + L"\\" + findData.cFileName;

                    if (SHGetFileInfoW(fullPath.c_str(), 0, &sfi, sizeof(sfi),
                                       SHGFI_DISPLAYNAME | SHGFI_ICON | SHGFI_LARGEICON)) {
                        wcsncpy_s(item.name, MAX_PATH, sfi.szDisplayName, _TRUNCATE);
                        wcsncpy_s(item.path, MAX_PATH, fullPath.c_str(), _TRUNCATE);
                        item.hIcon = sfi.hIcon;

                        WIN32_FILE_ATTRIBUTE_DATA fad = {};
                        if (GetFileAttributesExW(fullPath.c_str(), GetFileExInfoStandard, &fad)) {
                            item.ftWrite = fad.ftLastWriteTime;
                        } else {
                            item.ftWrite.dwLowDateTime = 0;
                            item.ftWrite.dwHighDateTime = 0;
                        }

                        const wchar_t* dot = wcsrchr(fullPath.c_str(), L'.');
                        item.extension = dot ? dot : L"";
                        item.hasPos = false;
                        g_desktopItems.push_back(item);
                    }
                }
            } while (FindNextFileW(hFind, &findData) != 0);
            FindClose(hFind);
        }
    }
    
    if (g_autoArrange) SortByName();

    LoadDesktopPositions();

    int screenW = GetSystemMetrics(SM_CXSCREEN);
    int curX = ICON_X_START;
    int curY = ICON_Y_START;
    for (auto &it : g_desktopItems) {
        if (!it.hasPos) {
            it.x = curX;
            it.y = curY;
            it.hasPos = true;
        }

        curX += g_iconXSpacing;
        if (curX + g_iconXSpacing > screenW) {
            curX = ICON_X_START;
            curY += g_iconYSpacing;
        }
    }

    InvalidateRect(g_hDesktopWnd, NULL, TRUE);
}

LRESULT CALLBACK DesktopWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);

        RECT rcClient;
        GetClientRect(hwnd, &rcClient);

        HDC hdcMem = CreateCompatibleDC(hdc);
        HBITMAP hbmMem = CreateCompatibleBitmap(hdc, rcClient.right, rcClient.bottom);
        HBITMAP hbmOld = (HBITMAP)SelectObject(hdcMem, hbmMem);

        // Draw background - either image or solid color
        if (g_backgroundImage) {
            // Draw the background image stretched to fill the screen
            Gdiplus::Graphics g(hdcMem);
            g.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);
            g.DrawImage(g_backgroundImage, 0, 0, rcClient.right, rcClient.bottom);
        } else {
            // Fallback to solid color
            HBRUSH br = CreateSolidBrush(RGB(32, 70, 130));
            FillRect(hdcMem, &rcClient, br);
            DeleteObject(br);
        }

        if (g_showIcons && !g_desktopItems.empty()) {
            Gdiplus::Graphics g(hdcMem);
            g.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);
            g.SetInterpolationMode(Gdiplus::InterpolationModeHighQuality);
            g.SetTextRenderingHint(Gdiplus::TextRenderingHintAntiAlias);

            for (size_t i = 0; i < g_desktopItems.size(); ++i) {
                const auto& item = g_desktopItems[i];
                int ix = item.x;
                int iy = item.y;

                if (item.hIcon) {
                    DrawIconWithGDIPlus(&g, item.hIcon, ix, iy, g_iconSize);
                }

                RECT rcText;
                rcText.left = ix - (g_iconXSpacing - g_iconSize) / 2;
                rcText.top = iy + g_iconSize;
                rcText.right = rcText.left + g_iconXSpacing;
                rcText.bottom = iy + g_iconYSpacing;

                SetTextColor(hdcMem, RGB(255, 255, 255));
                SetBkMode(hdcMem, TRANSPARENT);
                DrawTextW(hdcMem, item.name, -1, &rcText, DT_CENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
            }
        }

        BitBlt(hdc, 0, 0, rcClient.right, rcClient.bottom, hdcMem, 0, 0, SRCCOPY);

        SelectObject(hdcMem, hbmOld);
        DeleteObject(hbmMem);
        DeleteDC(hdcMem);

        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_DISPLAYCHANGE:
        ResizeShellWindows();
        return 0;
    case WM_DPICHANGED: {
        RECT* const pr = reinterpret_cast<RECT*>(lParam);
        if (pr) {
            SetWindowPos(g_hDesktopWnd, NULL, pr->left, pr->top,
                         pr->right - pr->left, pr->bottom - pr->top,
                         SWP_NOZORDER | SWP_NOACTIVATE);
        }
        ResizeShellWindows((int)wParam);

        PopulateDesktopIcons();
        return 0;
    }
    case WM_RBUTTONDOWN: {
        int clickX = LOWORD(lParam);
        int clickY = HIWORD(lParam);

        int sel = -1;
        if (g_showIcons) {
            for (size_t i = 0; i < g_desktopItems.size(); ++i) {
                const auto &it = g_desktopItems[i];
                RECT rcIcon;
                rcIcon.left = it.x - (g_iconXSpacing - g_iconSize) / 2;
                rcIcon.top = it.y;
                rcIcon.right = rcIcon.left + g_iconXSpacing;
                rcIcon.bottom = it.y + g_iconYSpacing;

                if (clickX >= rcIcon.left && clickX < rcIcon.right &&
                    clickY >= rcIcon.top && clickY < rcIcon.bottom) {
                    sel = (int)i;
                    break;
                }
            }
        }

        g_rightClickSelection = sel;

        POINT pt = { clickX, clickY };
        ClientToScreen(hwnd, &pt);

        if (sel >= 0) {
            ShowShellContextMenu(hwnd, pt, g_desktopItems[sel].path);
        } else {
            ShowDesktopContextMenu(hwnd, pt);
        }
        return 0;
    }
    case WM_COMMAND: {
        switch (LOWORD(wParam)) {
        case IDM_ITEM_OPEN:
            if (g_rightClickSelection >= 0 && g_rightClickSelection < (int)g_desktopItems.size()) {
                LaunchApp(g_desktopItems[g_rightClickSelection].path);
            }
            break;
        case IDM_ITEM_PROPERTIES:
            if (g_rightClickSelection >= 0 && g_rightClickSelection < (int)g_desktopItems.size()) {
                ShellExecuteW(NULL, L"properties", g_desktopItems[g_rightClickSelection].path, NULL, NULL, SW_SHOWDEFAULT);
            }
            break;
        case IDM_DESKTOP_REFRESH:
            PopulateDesktopIcons();
            ResetDesktopPositions();
            break;

        case IDM_VIEW_LARGE:
            UpdateViewMode(2);
            break;
        case IDM_VIEW_MEDIUM:
            UpdateViewMode(1);
            break;
        case IDM_VIEW_SMALL:
            UpdateViewMode(0);
            break;
        case IDM_SORT_NAME:
            SortByName();
            InvalidateRect(hwnd, NULL, TRUE);
            break;
        case IDM_SORT_TYPE:
            SortByTypeThenName();
            InvalidateRect(hwnd, NULL, TRUE);
            break;
        case IDM_SORT_DATE:
            SortByDateDesc();
            InvalidateRect(hwnd, NULL, TRUE);
            break;
        case IDM_TOGGLE_SHOW_ICONS:
            g_showIcons = !g_showIcons;
            InvalidateRect(hwnd, NULL, TRUE);
            break;
        case IDM_AUTO_ARRANGE:
            g_autoArrange = !g_autoArrange;
            if (g_autoArrange) SortByName();
            InvalidateRect(hwnd, NULL, TRUE);
            break;
        case IDM_ALIGN_TO_GRID:
            AlignIconsToGrid();
            break;
        case IDM_CHANGE_TASKBAR_COLOR: {
            // Choose color for taskbar
            CHOOSECOLORW cc = {};
            static COLORREF customColors[16] = {0};
            
            cc.lStructSize = sizeof(cc);
            cc.hwndOwner = hwnd;
            cc.lpCustColors = customColors;
            cc.rgbResult = g_taskbarColor;
            cc.Flags = CC_FULLOPEN | CC_RGBINIT;
            
            if (ChooseColorW(&cc)) {
                g_taskbarColor = cc.rgbResult;
                SaveTaskbarConfig();
                InvalidateRect(g_hTaskbarWnd, NULL, TRUE);
            }
            break;
        }
        case IDM_CHANGE_STARTMENU_COLOR: {
            // Choose color for start menu
            CHOOSECOLORW cc = {};
            static COLORREF customColors2[16] = {0};
            
            cc.lStructSize = sizeof(cc);
            cc.hwndOwner = hwnd;
            cc.lpCustColors = customColors2;
            cc.rgbResult = g_startMenuColor;
            cc.Flags = CC_FULLOPEN | CC_RGBINIT;
            
            if (ChooseColorW(&cc)) {
                g_startMenuColor = cc.rgbResult;
                SaveTaskbarConfig();
                InvalidateRect(g_hStartMenuWnd, NULL, TRUE);
            }
            break;
        }
        case IDM_CHANGE_BACKGROUND: {
            // Open file dialog to select background image
            OPENFILENAMEW ofn = {};
            wchar_t szFile[MAX_PATH] = L"";
            
            ofn.lStructSize = sizeof(ofn);
            ofn.hwndOwner = hwnd;
            ofn.lpstrFile = szFile;
            ofn.nMaxFile = MAX_PATH;
            ofn.lpstrFilter = L"Image Files\0*.jpg;*.jpeg;*.png;*.bmp;*.gif\0All Files\0*.*\0";
            ofn.nFilterIndex = 1;
            ofn.lpstrTitle = L"Select Background Image";
            ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;
            
            if (GetOpenFileNameW(&ofn)) {
                // Load the new background image
                if (g_backgroundImage) {
                    delete g_backgroundImage;
                    g_backgroundImage = NULL;
                }
                
                g_backgroundImage = Gdiplus::Bitmap::FromFile(szFile);
                if (g_backgroundImage && g_backgroundImage->GetLastStatus() == Gdiplus::Ok) {
                    g_backgroundPath = szFile;
                    SaveTaskbarConfig();
                    InvalidateRect(hwnd, NULL, TRUE);
                } else {
                    if (g_backgroundImage) {
                        delete g_backgroundImage;
                        g_backgroundImage = NULL;
                    }
                    MessageBoxW(hwnd, L"Failed to load image file.", L"Error", MB_OK | MB_ICONERROR);
                }
            }
            break;
        }
        }
        return 0;
    }
    case WM_LBUTTONDBLCLK: {
        int clickX = LOWORD(lParam);
        int clickY = HIWORD(lParam);

        for (const auto& item : g_desktopItems) {
            RECT rcIcon;
            rcIcon.left = item.x - (g_iconXSpacing - g_iconSize) / 2;
            rcIcon.top = item.y;
            rcIcon.right = rcIcon.left + g_iconXSpacing;
            rcIcon.bottom = item.y + g_iconYSpacing;

            if (clickX >= rcIcon.left && clickX < rcIcon.right &&
                clickY >= rcIcon.top && clickY < rcIcon.bottom)
            {
                LaunchApp(item.path);
                return 0;
            }
        }

        return 0;
    }
    case WM_LBUTTONDOWN: {
        int clickX = LOWORD(lParam);
        int clickY = HIWORD(lParam);

        if (IsWindowVisible(g_hStartMenuWnd)) {
            ShowWindow(g_hStartMenuWnd, SW_HIDE);
            InvalidateRect(g_hTaskbarWnd, NULL, TRUE);
        }

        int sel = -1;
        if (g_showIcons) {
            for (size_t i = 0; i < g_desktopItems.size(); ++i) {
                const auto &it = g_desktopItems[i];
                RECT rcIcon;
                rcIcon.left = it.x - (g_iconXSpacing - g_iconSize) / 2;
                rcIcon.top = it.y;
                rcIcon.right = rcIcon.left + g_iconXSpacing;
                rcIcon.bottom = it.y + g_iconYSpacing;

                if (clickX >= rcIcon.left && clickX < rcIcon.right &&
                    clickY >= rcIcon.top && clickY < rcIcon.bottom) {
                    sel = (int)i;
                    break;
                }
            }
        }

        g_rightClickSelection = sel;
        if (sel >= 0) {
            g_dragIndex = sel;
            g_dragging = true;
            SetCapture(hwnd);
            g_dragOffsetX = clickX - g_desktopItems[sel].x;
            g_dragOffsetY = clickY - g_desktopItems[sel].y;
        }
        InvalidateRect(hwnd, NULL, TRUE);
        return 0;
    }
    case WM_MOUSEMOVE: {
        if (g_dragging && g_dragIndex >= 0 && g_dragIndex < (int)g_desktopItems.size()) {
            int mx = LOWORD(lParam);
            int my = HIWORD(lParam);
            g_desktopItems[g_dragIndex].x = mx - g_dragOffsetX;
            g_desktopItems[g_dragIndex].y = my - g_dragOffsetY;
            return 0;
        } else if (!g_dragging && g_showIcons) {
            int mx = LOWORD(lParam);
            int my = HIWORD(lParam);
            int newHovered = -1;

            for (size_t i = 0; i < g_desktopItems.size(); ++i) {
                const auto &it = g_desktopItems[i];
                RECT rcIcon;
                rcIcon.left = it.x - (g_iconXSpacing - g_iconSize) / 2;
                rcIcon.top = it.y;
                rcIcon.right = rcIcon.left + g_iconXSpacing;
                rcIcon.bottom = it.y + g_iconYSpacing;

                if (mx >= rcIcon.left && mx < rcIcon.right &&
                    my >= rcIcon.top && my < rcIcon.bottom) {
                    newHovered = (int)i;
                    break;
                }
            }

            if (newHovered != g_hoveredIcon) {
                g_hoveredIcon = newHovered;
                InvalidateRect(g_hDesktopWnd, NULL, FALSE);
            }
        }
        break;
    }
    case WM_LBUTTONUP: {
        if (g_dragging) {
            ReleaseCapture();
            g_dragging = false;
            if (g_dragIndex >= 0 && g_dragIndex < (int)g_desktopItems.size()) {
                if (!g_autoArrange) {
                    auto &it = g_desktopItems[g_dragIndex];
                    int relX = it.x - ICON_X_START;
                    int relY = it.y - ICON_Y_START;
                    int col = (relX + g_iconXSpacing/2) / g_iconXSpacing;
                    int row = (relY + g_iconYSpacing/2) / g_iconYSpacing;
                    if (col < 0) col = 0;
                    if (row < 0) row = 0;
                    it.x = ICON_X_START + col * g_iconXSpacing;
                    it.y = ICON_Y_START + row * g_iconYSpacing;
                }
                SaveDesktopPositions();
            }
            g_dragIndex = -1;
            InvalidateRect(hwnd, NULL, TRUE);
        }
        return 0;
    }
    case WM_USER + 1: {
        // Delayed refresh after file operations
        PopulateDesktopIcons();
        InvalidateRect(hwnd, NULL, TRUE);
        return 0;
    }
    case WM_TIMER: {
        if (wParam == 999) {
            // Kill the timer
            KillTimer(hwnd, 999);
            // Do the delayed refresh
            PopulateDesktopIcons();
            InvalidateRect(hwnd, NULL, TRUE);
        }
        return 0;
    }
    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

LRESULT CALLBACK TaskbarWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);

        RECT rcClient;
        GetClientRect(hwnd, &rcClient);

        // Create buffer for double-buffering
        HDC hdcMem = CreateCompatibleDC(hdc);
        HBITMAP hbmMem = CreateCompatibleBitmap(hdc, rcClient.right, rcClient.bottom);
        HBITMAP hbmOld = (HBITMAP)SelectObject(hdcMem, hbmMem);

        // Modern semi-transparent dark background (using custom color)
        HBRUSH br = CreateSolidBrush(g_taskbarColor);
        FillRect(hdcMem, &rcClient, br);
        DeleteObject(br);

        // Draw with GDI+
        Gdiplus::Graphics g(hdcMem);
        g.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);

        // Draw Start button as blue square (no text)
        int btnX = 6, btnY = 6, btnW = 55, btnH = g_taskbarHeight - 12;
        Gdiplus::RectF btnRect(btnX, btnY, btnW, btnH);
        Gdiplus::SolidBrush btnBrush(Gdiplus::Color(255, 0, 120, 220));

        // Rounded square button
        Gdiplus::GraphicsPath btnPath;
        float radius = 4.0f;
        btnPath.AddArc(btnRect.X, btnRect.Y, radius, radius, 180, 90);
        btnPath.AddArc(btnRect.X + btnRect.Width - radius, btnRect.Y, radius, radius, 270, 90);
        btnPath.AddArc(btnRect.X + btnRect.Width - radius, btnRect.Y + btnRect.Height - radius, radius, radius, 0, 90);
        btnPath.AddArc(btnRect.X, btnRect.Y + btnRect.Height - radius, radius, radius, 90, 90);
        btnPath.CloseFigure();

        g.FillPath(&btnBrush, &btnPath);

        // Draw app icons next to Start button (Windows 11 style)
        int iconSize = 40;
        int iconSpacing = 10;
        int iconStartX = btnX + btnW + 10;
        int currentX = iconStartX;

        // Draw pinned apps with Windows 11 style indicators
        for (size_t i = 0; i < g_taskbarIcons.size(); ++i) {
            if (!g_taskbarIcons[i].visible) continue;

            int y = (g_taskbarHeight - iconSize) / 2;
            g_taskbarIcons[i].x = currentX;
            g_taskbarIcons[i].y = y;
            g_taskbarIcons[i].width = iconSize;
            g_taskbarIcons[i].height = iconSize;

            if (g_taskbarIcons[i].hIcon) {
                DrawIconEx(hdcMem, currentX, y, g_taskbarIcons[i].hIcon, iconSize, iconSize, 0, NULL, DI_NORMAL);

                // WINDOWS 11 STYLE: Check if app is running
                bool isRunning = false;
                bool isMinimized = false;
                bool isActive = false;
                HWND appWindow = NULL;

                for (const auto& app : g_runningApps) {
                    if (_wcsicmp(GetAppExeName(app.exePath.c_str()).c_str(),
                                GetAppExeName(g_taskbarIcons[i].path.c_str()).c_str()) == 0) {
                        isRunning = true;
                        appWindow = app.hwnd;
                        if (appWindow) {
                            isMinimized = IsIconic(appWindow);
                            isActive = (GetForegroundWindow() == appWindow);
                        }
                        break;
                    }
                }

                // WINDOWS 11 STYLE INDICATORS
                if (isRunning) {
                    Gdiplus::Pen* indicatorPen = nullptr;

                    if (isActive) {
                        indicatorPen = new Gdiplus::Pen(Gdiplus::Color(255, 0, 120, 215), 3.0f);
                    } else if (isMinimized) {
                        indicatorPen = new Gdiplus::Pen(Gdiplus::Color(150, 150, 150, 150), 2.0f);
                    } else {
                        indicatorPen = new Gdiplus::Pen(Gdiplus::Color(200, 100, 180, 255), 2.5f);
                    }

                    int indicatorY = y + iconSize - 4;
                    g.DrawLine(indicatorPen, currentX + 8, indicatorY,
                               currentX + iconSize - 8, indicatorY);

                    delete indicatorPen;
                }
            }
            currentX += iconSize + iconSpacing;
        }

        // Draw running apps that aren't pinned (to the right of pinned apps)
        int runningAppsX = currentX + 20;

        for (size_t i = 0; i < g_runningApps.size(); ++i) {
            bool alreadyPinned = false;
            for (const auto& pinned : g_taskbarIcons) {
                if (_wcsicmp(GetAppExeName(g_runningApps[i].exePath.c_str()).c_str(),
                            GetAppExeName(pinned.path.c_str()).c_str()) == 0) {
                    alreadyPinned = true;
                    break;
                }
            }

            if (!alreadyPinned) {
                if (runningAppsX + iconSize > rcClient.right - 150) break;

                int y = (g_taskbarHeight - iconSize) / 2;

                if (g_runningApps[i].hIcon) {
                    DrawIconEx(hdcMem, runningAppsX, y, g_runningApps[i].hIcon,
                              iconSize, iconSize, 0, NULL, DI_NORMAL);
                } else {
                    Gdiplus::SolidBrush colorBrush(Gdiplus::Color(255, 100, 150, 200));
                    g.FillRectangle(&colorBrush, runningAppsX, y, iconSize, iconSize);
                }

                g_runningApps[i].visible = true;

                HWND foreground = GetForegroundWindow();
                bool isMinimized = IsIconic(g_runningApps[i].hwnd);
                bool isActive = (foreground == g_runningApps[i].hwnd);

                Gdiplus::Pen* indicatorPen = nullptr;

                if (isActive) {
                    indicatorPen = new Gdiplus::Pen(Gdiplus::Color(255, 255, 200, 0), 3.0f);
                } else if (isMinimized) {
                    indicatorPen = new Gdiplus::Pen(Gdiplus::Color(150, 100, 100, 100), 2.0f);
                } else {
                    indicatorPen = new Gdiplus::Pen(Gdiplus::Color(255, 0, 180, 255), 2.5f);
                }

                int indicatorY = y + iconSize - 4;
                g.DrawLine(indicatorPen, runningAppsX + 8, indicatorY,
                           runningAppsX + iconSize - 8, indicatorY);

                delete indicatorPen;

                runningAppsX += iconSize + iconSpacing;
            }
        }

        // ==================== NEW: IMPROVED TRAY ICONS AREA ====================
        
        // Calculate tray area width based on number of icons
        int trayAreaWidth = 180; // Increased from 120 to fit all icons with proper spacing
        
        // Draw background for tray area (same as taskbar color)
        BYTE tr = GetRValue(g_taskbarColor);
        BYTE tg = GetGValue(g_taskbarColor);
        BYTE tb = GetBValue(g_taskbarColor);
        Gdiplus::SolidBrush trayBgBrush(Gdiplus::Color(255, tr, tg, tb));
        g.FillRectangle(&trayBgBrush, rcClient.right - trayAreaWidth, 0, trayAreaWidth, rcClient.bottom);
        
        // Draw separator line (subtle)
        Gdiplus::Pen separatorPen(Gdiplus::Color(50, 100, 100, 100), 1.0f);
        g.DrawLine(&separatorPen, rcClient.right - trayAreaWidth, 5, 
                   rcClient.right - trayAreaWidth, rcClient.bottom - 5);
        
        // Position tray icons horizontally (Windows 11 style)
        int trayStartX = rcClient.right - trayAreaWidth + 8;
        int trayY = (rcClient.bottom - g_trayIconSize) / 2;
        int currentTrayX = trayStartX;
        
        // Draw each tray icon
        for (size_t i = 0; i < g_trayIcons.size(); ++i) {
            // Use the icon's actual width (clock is wider)
            int iconWidth = g_trayIcons[i].width;
            
            if (currentTrayX + iconWidth > rcClient.right - 5) {
                // Not enough space - this icon will be hidden
                break;
            }
            
            // Store position for click detection
            g_trayIcons[i].x = currentTrayX;
            g_trayIcons[i].y = trayY;
            // Width is already set in AddCommonTrayIcons
            g_trayIcons[i].height = g_trayIconSize;
            
            // Draw icon
            if (g_trayIcons[i].name == L"Clock") {
                // Special handling for clock (draw text) - use actual width
                DrawTrayTime(&g, currentTrayX, trayY, g_trayIcons[i].width, g_trayIconSize);
            } else if (g_trayIcons[i].hIcon) {
                // Draw actual icon
                DrawIconEx(hdcMem, currentTrayX, trayY, g_trayIcons[i].hIcon, 
                          g_trayIconSize, g_trayIconSize, 0, NULL, DI_NORMAL);
            } else {
                // Fallback: draw colored rectangle with app initial
                Gdiplus::Color colors[] = {
                    Gdiplus::Color(200, 52, 152, 219),
                    Gdiplus::Color(200, 230, 126, 34),
                    Gdiplus::Color(200, 46, 204, 113),
                    Gdiplus::Color(200, 231, 76, 60),
                    Gdiplus::Color(200, 155, 89, 182)
                };
                
                Gdiplus::SolidBrush iconBgBrush(colors[i % 5]);
                g.FillRectangle(&iconBgBrush, currentTrayX, trayY, g_trayIconSize, g_trayIconSize);
                
                // Draw first letter of app name
                if (!g_trayIcons[i].name.empty()) {
                    Gdiplus::FontFamily fontFamily(L"Segoe UI");
                    Gdiplus::Font iconFont(&fontFamily, 10, Gdiplus::FontStyleBold, Gdiplus::UnitPixel);
                    Gdiplus::SolidBrush textBrush(Gdiplus::Color(255, 255, 255, 255));
                    
                    wchar_t firstChar = g_trayIcons[i].name[0];
                    wchar_t charStr[2] = { firstChar, 0 };
                    
                    Gdiplus::RectF textRect(currentTrayX, trayY, g_trayIconSize, g_trayIconSize);
                    Gdiplus::StringFormat format;
                    format.SetAlignment(Gdiplus::StringAlignmentCenter);
                    format.SetLineAlignment(Gdiplus::StringAlignmentCenter);
                    
                    g.DrawString(charStr, 1, &iconFont, textRect, &format, &textBrush);
                }
            }
            
            // Draw subtle border (but not for clock)
            if (g_trayIcons[i].name != L"Clock") {
                Gdiplus::Pen iconBorderPen(Gdiplus::Color(80, 180, 180, 180), 0.5f);
                g.DrawRectangle(&iconBorderPen, currentTrayX, trayY, g_trayIconSize, g_trayIconSize);
            }
            
            // Advance by the icon's actual width plus spacing
            currentTrayX += g_trayIcons[i].width + g_trayIconSpacing;
        }
        
        // Draw chevron (v) for hidden icons if needed
        if (currentTrayX + g_trayIconSize > rcClient.right - 5) {
            g_hasHiddenIcons = true;
            g_chevronX = rcClient.right - g_chevronSize - 5;
            g_chevronY = (rcClient.bottom - g_chevronSize) / 2;
            
            // Draw chevron background (slightly darker than taskbar)
            BYTE r = (std::max)(0, (int)GetRValue(g_taskbarColor) - 20);
            BYTE g_val = (std::max)(0, (int)GetGValue(g_taskbarColor) - 20);
            BYTE b = (std::max)(0, (int)GetBValue(g_taskbarColor) - 20);
            Gdiplus::SolidBrush chevronBgBrush(Gdiplus::Color(150, r, g_val, b));
            g.FillRectangle(&chevronBgBrush, g_chevronX, g_chevronY, g_chevronSize, g_chevronSize);
            
            // Draw chevron arrow (v)
            Gdiplus::FontFamily chevronFontFamily(L"Segoe UI Symbol");
            Gdiplus::Font chevronFont(&chevronFontFamily, 14, Gdiplus::FontStyleBold, Gdiplus::UnitPixel);
            Gdiplus::SolidBrush chevronBrush(Gdiplus::Color(255, 200, 200, 200));
            
            Gdiplus::RectF chevronRect(g_chevronX, g_chevronY, g_chevronSize, g_chevronSize);
            Gdiplus::StringFormat chevronFormat;
            chevronFormat.SetAlignment(Gdiplus::StringAlignmentCenter);
            chevronFormat.SetLineAlignment(Gdiplus::StringAlignmentCenter);
            
            g.DrawString(L"▼", -1, &chevronFont, chevronRect, &chevronFormat, &chevronBrush);
        } else {
            g_hasHiddenIcons = false;
        }
        
        // Draw system info on right-most side (optional)
        /*
        SYSTEMTIME st;
        GetLocalTime(&st);
        wchar_t timeStr[64];
        wsprintfW(timeStr, L"%02d:%02d", st.wHour, st.wMinute);
        
        Gdiplus::FontFamily timeFontFamily(L"Segoe UI");
        Gdiplus::Font timeFont(&timeFontFamily, 11, Gdiplus::FontStyleRegular, Gdiplus::UnitPixel);
        Gdiplus::SolidBrush timeBrush(Gdiplus::Color(255, 220, 220, 220));
        
        Gdiplus::RectF timeRect(rcClient.right - 80, trayY, 75, g_trayIconSize);
        Gdiplus::StringFormat timeFormat;
        timeFormat.SetAlignment(Gdiplus::StringAlignmentFar);
        timeFormat.SetLineAlignment(Gdiplus::StringAlignmentCenter);
        
        g.DrawString(timeStr, -1, &timeFont, timeRect, &timeFormat, &timeBrush);
        */

        // Blit buffer
        BitBlt(hdc, 0, 0, rcClient.right, rcClient.bottom, hdcMem, 0, 0, SRCCOPY);
        SelectObject(hdcMem, hbmOld);
        DeleteObject(hbmMem);
        DeleteDC(hdcMem);

        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_RBUTTONDOWN: {
        int clickX = LOWORD(lParam);
        int clickY = HIWORD(lParam);

        // Check if click is on a tray icon
        for (size_t i = 0; i < g_trayIcons.size(); ++i) {
            if (clickX >= g_trayIcons[i].x && clickX < g_trayIcons[i].x + g_trayIcons[i].width &&
                clickY >= g_trayIcons[i].y && clickY < g_trayIcons[i].y + g_trayIcons[i].height) {
                
                HandleTrayIconClick(i, true); // Right click
                return 0;
            }
        }

        // Existing taskbar right-click menu code...
        HMENU hMenu = CreatePopupMenu();
        AppendMenuW(hMenu, MF_STRING, 40010, L"Pin another app to taskbar");
        AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);

        for (size_t i = 0; i < g_taskbarIcons.size() && i < 10; ++i) {
            if (g_taskbarIcons[i].visible &&
                clickX >= g_taskbarIcons[i].x && clickX < g_taskbarIcons[i].x + g_taskbarIcons[i].width &&
                clickY >= g_taskbarIcons[i].y && clickY < g_taskbarIcons[i].y + g_taskbarIcons[i].height) {

                std::wstring appPath = g_taskbarIcons[i].path;
                size_t lastSlash = appPath.find_last_of(L"\\");
                std::wstring appName = (lastSlash != std::wstring::npos) ?
                    appPath.substr(lastSlash + 1) : appPath;

                AppendMenuW(hMenu, MF_STRING, 40000 + i, (L"Unpin \"" + appName + L"\" from taskbar").c_str());
                break;
            }
        }

        POINT pt;
        pt.x = clickX;
        pt.y = clickY;
        ClientToScreen(hwnd, &pt);

        int cmd = TrackPopupMenu(hMenu, TPM_RETURNCMD, pt.x, pt.y, 0, hwnd, NULL);

        if (cmd >= 40000 && cmd < 40010) {
            int idx = cmd - 40000;
            if (idx >= 0 && idx < (int)g_taskbarIcons.size()) {
                g_taskbarPinned.erase(
                    std::remove(g_taskbarPinned.begin(), g_taskbarPinned.end(), g_taskbarIcons[idx].path),
                    g_taskbarPinned.end()
                );
                SaveTaskbarConfig();
                PopulateTaskbarIcons();
                InvalidateRect(hwnd, NULL, TRUE);
            }
        } else if (cmd == 40010) {
            ShowWindow(g_hStartMenuWnd, IsWindowVisible(g_hStartMenuWnd) ? SW_HIDE : SW_SHOW);
        }

        DestroyMenu(hMenu);
        return 0;
    }
    case WM_LBUTTONDOWN: {
        int clickX = LOWORD(lParam);
        int clickY = HIWORD(lParam);

        // ==================== NEW: Handle tray icon clicks ====================
        for (size_t i = 0; i < g_trayIcons.size(); ++i) {
            if (clickX >= g_trayIcons[i].x && clickX < g_trayIcons[i].x + g_trayIcons[i].width &&
                clickY >= g_trayIcons[i].y && clickY < g_trayIcons[i].y + g_trayIcons[i].height) {
                
                HandleTrayIconClick(i, false); // Left click
                return 0;
            }
        }

        // Check chevron (hidden icons) click
        if (g_hasHiddenIcons &&
            clickX >= g_chevronX && clickX < g_chevronX + g_chevronSize &&
            clickY >= g_chevronY && clickY < g_chevronY + g_chevronSize) {

            // Show hidden tray icons popup menu
            HMENU hMenu = CreatePopupMenu();
            
            // Find which icons are hidden (index where we stopped drawing)
            int lastVisibleIndex = -1;
            for (size_t i = 0; i < g_trayIcons.size(); ++i) {
                if (g_trayIcons[i].x + g_trayIcons[i].width > g_chevronX) {
                    lastVisibleIndex = i - 1;
                    break;
                }
            }
            
            if (lastVisibleIndex >= 0 && lastVisibleIndex < (int)g_trayIcons.size() - 1) {
                // Add hidden icons to menu
                for (size_t i = lastVisibleIndex + 1; i < g_trayIcons.size(); ++i) {
                    std::wstring menuText = g_trayIcons[i].name;
                    if (menuText.empty()) menuText = L"Icon " + std::to_wstring(i);
                    
                    AppendMenuW(hMenu, MF_STRING, 51000 + i, menuText.c_str());
                }
                
                POINT pt;
                GetCursorPos(&pt);
                
                int cmd = TrackPopupMenu(hMenu, TPM_RETURNCMD, pt.x, pt.y, 0, hwnd, NULL);
                
                if (cmd >= 51000 && cmd < 51000 + (int)g_trayIcons.size()) {
                    int idx = cmd - 51000;
                    if (idx >= 0 && idx < (int)g_trayIcons.size()) {
                        HandleTrayIconClick(idx, false);
                    }
                }
            }
            
            DestroyMenu(hMenu);
            return 0;
        }

        // Check Start button clicked
        if (clickX >= 6 && clickX < 61 && clickY >= 6 && clickY < g_taskbarHeight - 6) {
            ShowWindow(g_hStartMenuWnd, IsWindowVisible(g_hStartMenuWnd) ? SW_HIDE : SW_SHOW);
            InvalidateRect(hwnd, NULL, TRUE);
            return 0;
        }

        // WINDOWS 11 TASK BEHAVIOR FOR PINNED APPS
        for (size_t i = 0; i < g_taskbarIcons.size(); ++i) {
            if (g_taskbarIcons[i].visible &&
                clickX >= g_taskbarIcons[i].x && clickX < g_taskbarIcons[i].x + g_taskbarIcons[i].width &&
                clickY >= g_taskbarIcons[i].y && clickY < g_taskbarIcons[i].y + g_taskbarIcons[i].height) {

                HWND targetWindow = NULL;
                bool isRunning = false;

                for (const auto& app : g_runningApps) {
                    if (_wcsicmp(GetAppExeName(app.exePath.c_str()).c_str(),
                                GetAppExeName(g_taskbarIcons[i].path.c_str()).c_str()) == 0) {
                        targetWindow = app.hwnd;
                        isRunning = true;
                        break;
                    }
                }

                if (isRunning && targetWindow && IsWindow(targetWindow)) {
                    bool isMinimized = IsIconic(targetWindow);
                    HWND foregroundWindow = GetForegroundWindow();

                    if (isMinimized) {
                        ShowWindow(targetWindow, SW_RESTORE);
                        SetForegroundWindow(targetWindow);
                        SetWindowPos(targetWindow, HWND_TOP, 0, 0, 0, 0,
                                    SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
                        FlashWindowIfMinimized(targetWindow);
                    }
                    else if (foregroundWindow == targetWindow) {
                        ShowWindow(targetWindow, SW_MINIMIZE);
                        FlashWindowIfMinimized(targetWindow);
                    }
                    else {
                        SetForegroundWindow(targetWindow);
                        SetWindowPos(targetWindow, HWND_TOP, 0, 0, 0, 0,
                                    SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);

                        FLASHWINFO fwi;
                        fwi.cbSize = sizeof(FLASHWINFO);
                        fwi.hwnd = targetWindow;
                        fwi.dwFlags = FLASHW_TRAY;
                        fwi.uCount = 1;
                        fwi.dwTimeout = 0;
                        FlashWindowEx(&fwi);
                    }
                }
                else {
                    LaunchApp(g_taskbarIcons[i].path.c_str());
                }
                return 0;
            }
        }

        // WINDOWS 11 BEHAVIOR FOR RUNNING APPS (NOT PINNED)
        int iconSize = 40;
        int iconSpacing = 10;
        int runningAppsStartX = 6 + 55 + 10;

        for (size_t i = 0; i < g_taskbarIcons.size(); ++i) {
            if (g_taskbarIcons[i].visible) {
                runningAppsStartX += iconSize + iconSpacing;
            }
        }
        runningAppsStartX += 20;

        int currentX = runningAppsStartX;
        for (size_t i = 0; i < g_runningApps.size(); ++i) {
            bool alreadyPinned = false;
            for (const auto& pinned : g_taskbarIcons) {
                if (_wcsicmp(GetAppExeName(g_runningApps[i].exePath.c_str()).c_str(),
                            GetAppExeName(pinned.path.c_str()).c_str()) == 0) {
                    alreadyPinned = true;
                    break;
                }
            }

            if (!alreadyPinned) {
                int y = (g_taskbarHeight - iconSize) / 2;

                if (clickX >= currentX && clickX < currentX + iconSize &&
                    clickY >= y && clickY < y + iconSize) {

                    HWND targetWindow = g_runningApps[i].hwnd;
                    bool isMinimized = IsIconic(targetWindow);
                    HWND foregroundWindow = GetForegroundWindow();

                    if (isMinimized) {
                        ShowWindow(targetWindow, SW_RESTORE);
                        SetForegroundWindow(targetWindow);
                        SetWindowPos(targetWindow, HWND_TOP, 0, 0, 0, 0,
                                    SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
                    }
                    else if (foregroundWindow == targetWindow) {
                        ShowWindow(targetWindow, SW_MINIMIZE);
                        FlashWindowIfMinimized(targetWindow);
                    }
                    else {
                        SetForegroundWindow(targetWindow);
                        SetWindowPos(targetWindow, HWND_TOP, 0, 0, 0, 0,
                                    SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
                    }
                    return 0;
                }
                currentX += iconSize + iconSpacing;
            }
        }

        return 0;
    }
    case WM_TIMER: {
        // Update window activation
        HWND currentForeground = GetForegroundWindow();
        if (currentForeground != g_lastForegroundWindow) {
            g_lastForegroundWindow = currentForeground;
            InvalidateRect(hwnd, NULL, FALSE);
        }

        // Keep taskbar on top of normal windows (but not topmost)
        SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);

        // Continuously re-apply work area to force Windows to respect it
        static int workAreaCounter = 0;
        if (++workAreaCounter % 4 == 0) { // Every 2 seconds (500ms * 4)
            int screenW = GetSystemMetrics(SM_CXSCREEN);
            int screenH = GetSystemMetrics(SM_CYSCREEN);
            RECT workArea;
            workArea.left = 0;
            workArea.top = 0;
            workArea.right = screenW;
            workArea.bottom = screenH - g_taskbarHeight;
            SystemParametersInfoW(SPI_SETWORKAREA, 0, &workArea, SPIF_SENDCHANGE | SPIF_UPDATEINIFILE);
            
            // Broadcast WM_SETTINGCHANGE to all windows
            SendMessageTimeoutW(HWND_BROADCAST, WM_SETTINGCHANGE, SPI_SETWORKAREA, 
                              (LPARAM)L"WorkArea", SMTO_ABORTIFHUNG, 100, NULL);
        }

        // Update running apps list
        UpdateRunningAppsList();
        
        // Periodically refresh tray icons (every 5 seconds)
        static DWORD lastTrayUpdate = 0;
        DWORD currentTime = GetTickCount();
        if (currentTime - lastTrayUpdate > 5000) { // 5 seconds
            PopulateRealTrayIcons();
            lastTrayUpdate = currentTime;
            InvalidateRect(hwnd, NULL, FALSE);
        } else {
            InvalidateRect(hwnd, NULL, FALSE);
        }
        return 0;
    }
    case APPBAR_CALLBACK:
    {
        switch (wParam)
        {
            case ABN_FULLSCREENAPP:
                // A fullscreen app opened or closed
                if (lParam) {
                    // Fullscreen app is opening - hide taskbar
                    ShowWindow(hwnd, SW_HIDE);
                } else {
                    // Fullscreen app closed - show taskbar
                    ShowWindow(hwnd, SW_SHOW);
                }
                break;
            case ABN_POSCHANGED:
                SetAppBarPos(hwnd);
                break;
        }
        return 0;
    }
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

// ==================== START MENU (unchanged) ====================

LRESULT CALLBACK StartMenuWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        RECT rcClient;
        GetClientRect(hwnd, &rcClient);

        HDC hdcMem = CreateCompatibleDC(hdc);
        HBITMAP hbmMem = CreateCompatibleBitmap(hdc, rcClient.right, rcClient.bottom);
        HBITMAP hbmOld = (HBITMAP)SelectObject(hdcMem, hbmMem);

        HBRUSH br = CreateSolidBrush(g_startMenuColor);
        FillRect(hdcMem, &rcClient, br);
        DeleteObject(br);

        Gdiplus::Graphics g(hdcMem);
        g.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);
        g.SetTextRenderingHint(Gdiplus::TextRenderingHintAntiAlias);

        g_sidebarItems.clear();

        int sidebarWidth = 350;
        // Use the same color as start menu background
        BYTE r = GetRValue(g_startMenuColor);
        BYTE g_val = GetGValue(g_startMenuColor);
        BYTE b = GetBValue(g_startMenuColor);
        Gdiplus::SolidBrush sidebarBrush(Gdiplus::Color(255, r, g_val, b));
        g.FillRectangle(&sidebarBrush, 0, 0, sidebarWidth, rcClient.bottom);

        int sidebarY = 15 - g_sidebarScrollOffset;

        if (g_currentSidebarFolder != L"C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs") {
            if (sidebarY >= 0 && sidebarY < rcClient.bottom) {
                SidebarItemInfo backBtn;
                backBtn.isFolder = true;
                backBtn.name = L"..";
                backBtn.fullPath = L"PARENT";
                backBtn.yPos = sidebarY;
                backBtn.height = 32;
                g_sidebarItems.push_back(backBtn);

                Gdiplus::RectF backRect(10, sidebarY, 24, 24);
                Gdiplus::SolidBrush backBrush(Gdiplus::Color(255, 180, 180, 180));
                g.FillRectangle(&backBrush, backRect);

                Gdiplus::FontFamily fontFamily(L"Segoe UI");
                Gdiplus::Font sidebarFont(&fontFamily, 11, Gdiplus::FontStyleRegular, Gdiplus::UnitPixel);
                Gdiplus::SolidBrush textBrush(Gdiplus::Color(255, 220, 220, 220));

                Gdiplus::RectF backTextRect(10, sidebarY, sidebarWidth - 20, 24);
                Gdiplus::StringFormat centerFormat;
                centerFormat.SetAlignment(Gdiplus::StringAlignmentCenter);
                centerFormat.SetLineAlignment(Gdiplus::StringAlignmentCenter);
                g.DrawString(L"..", -1, &sidebarFont, backTextRect, &centerFormat, &textBrush);
            }
            sidebarY += 32;
        }

        WIN32_FIND_DATAW findData;
        HANDLE hFind = FindFirstFileW((g_currentSidebarFolder + L"\\*").c_str(), &findData);

        std::vector<std::pair<std::wstring, bool>> items;

        if (hFind != INVALID_HANDLE_VALUE) {
            do {
                if (!(findData.dwFileAttributes & FILE_ATTRIBUTE_HIDDEN) &&
                    wcscmp(findData.cFileName, L".") != 0 &&
                    wcscmp(findData.cFileName, L"..") != 0) {

                    bool isFolder = (findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                    std::wstring name = findData.cFileName;

                    if (isFolder || wcsstr(findData.cFileName, L".lnk")) {
                        items.push_back({name, isFolder});
                    }
                }
            } while (FindNextFileW(hFind, &findData));
            FindClose(hFind);
        }

        std::sort(items.begin(), items.end(),
            [](const std::pair<std::wstring, bool>& a, const std::pair<std::wstring, bool>& b) {
                return _wcsicmp(a.first.c_str(), b.first.c_str()) < 0;
            });

        Gdiplus::FontFamily fontFamily(L"Segoe UI");
        Gdiplus::Font sidebarFont(&fontFamily, 11, Gdiplus::FontStyleRegular, Gdiplus::UnitPixel);
        Gdiplus::SolidBrush textBrush(Gdiplus::Color(255, 220, 220, 220));
        Gdiplus::SolidBrush folderBrush(Gdiplus::Color(255, 150, 150, 200));

        for (const auto& item : items) {
            if (sidebarY > rcClient.bottom) break;
            if (sidebarY + 32 < 0) {
                sidebarY += 32;
                continue;
            }

            std::wstring fullPath = g_currentSidebarFolder + L"\\" + item.first;

            SidebarItemInfo info;
            info.isFolder = item.second;
            info.name = item.first;
            info.fullPath = fullPath;
            info.yPos = sidebarY;
            info.height = 32;
            g_sidebarItems.push_back(info);

            HICON hIcon = NULL;
            SHFILEINFOW sfi = {};
            DWORD flags = SHGFI_ICON | (item.second ? SHGFI_SMALLICON : SHGFI_LARGEICON);
            if (SHGetFileInfoW(fullPath.c_str(), 0, &sfi, sizeof(sfi), flags)) {
                hIcon = sfi.hIcon;
            }

            if (hIcon) {
                DrawIconWithGDIPlus(&g, hIcon, 10, sidebarY, item.second ? 20 : 24);
                DestroyIcon(hIcon);
            }

            if (item.second) {
                Gdiplus::RectF folderIndicator(34, sidebarY + 8, 4, 4);
                Gdiplus::SolidBrush indicatorBrush(Gdiplus::Color(255, 100, 200, 255));
                g.FillEllipse(&indicatorBrush, folderIndicator);
            }

            std::wstring displayName = item.first;
            if (item.second == false && displayName.find(L".lnk") != std::wstring::npos) {
                displayName = displayName.substr(0, displayName.length() - 4);
            }
            if (displayName.length() > 18) {
                displayName = displayName.substr(0, 15) + L"...";
            }

            Gdiplus::RectF textRect(40, sidebarY, sidebarWidth - 50, 24);
            Gdiplus::StringFormat format;
            format.SetAlignment(Gdiplus::StringAlignmentNear);
            format.SetLineAlignment(Gdiplus::StringAlignmentCenter);

            Gdiplus::SolidBrush* brush = item.second ? &folderBrush : &textBrush;
            g.DrawString(displayName.c_str(), -1, &sidebarFont, textRect, &format, brush);

            sidebarY += 32;
        }

        // Draw power buttons at the bottom of sidebar
        int powerButtonY = rcClient.bottom - 60; // 60 pixels from bottom
        int powerButtonWidth = 100;
        int powerButtonHeight = 40;
        int powerButtonSpacing = 10;
        int totalButtonWidth = (powerButtonWidth * 3) + (powerButtonSpacing * 2);
        int powerButtonStartX = (sidebarWidth - totalButtonWidth) / 2; // Center buttons

        Gdiplus::Font powerFont(&fontFamily, 10, Gdiplus::FontStyleBold, Gdiplus::UnitPixel);
        Gdiplus::SolidBrush powerTextBrush(Gdiplus::Color(255, 255, 255, 255));
        Gdiplus::StringFormat centerFormat;
        centerFormat.SetAlignment(Gdiplus::StringAlignmentCenter);
        centerFormat.SetLineAlignment(Gdiplus::StringAlignmentCenter);

        // Shutdown button
        Gdiplus::RectF shutdownRect(powerButtonStartX, powerButtonY, powerButtonWidth, powerButtonHeight);
        Gdiplus::SolidBrush shutdownBrush(Gdiplus::Color(255, 200, 50, 50));
        g.FillRectangle(&shutdownBrush, shutdownRect);
        Gdiplus::Pen shutdownPen(Gdiplus::Color(255, 150, 30, 30), 2);
        g.DrawRectangle(&shutdownPen, shutdownRect);
        g.DrawString(L"Shutdown", -1, &powerFont, shutdownRect, &centerFormat, &powerTextBrush);

        // Restart button
        Gdiplus::RectF restartRect(powerButtonStartX + powerButtonWidth + powerButtonSpacing, powerButtonY, powerButtonWidth, powerButtonHeight);
        Gdiplus::SolidBrush restartBrush(Gdiplus::Color(255, 50, 150, 200));
        g.FillRectangle(&restartBrush, restartRect);
        Gdiplus::Pen restartPen(Gdiplus::Color(255, 30, 100, 150), 2);
        g.DrawRectangle(&restartPen, restartRect);
        g.DrawString(L"Restart", -1, &powerFont, restartRect, &centerFormat, &powerTextBrush);

        // Sleep button
        Gdiplus::RectF sleepRect(powerButtonStartX + (powerButtonWidth + powerButtonSpacing) * 2, powerButtonY, powerButtonWidth, powerButtonHeight);
        Gdiplus::SolidBrush sleepBrush(Gdiplus::Color(255, 100, 100, 100));
        g.FillRectangle(&sleepBrush, sleepRect);
        Gdiplus::Pen sleepPen(Gdiplus::Color(255, 60, 60, 60), 2);
        g.DrawRectangle(&sleepPen, sleepRect);
        g.DrawString(L"Sleep", -1, &powerFont, sleepRect, &centerFormat, &powerTextBrush);

        int contentX = sidebarWidth + 20;
        int contentWidth = rcClient.right - contentX - 20;

        int tileSize = 100;
        int tileSpacing = 120;
        int tilesPerRow = contentWidth / tileSpacing;
        if (tilesPerRow < 1) tilesPerRow = 1;

        Gdiplus::Font appNameFont(&fontFamily, 10, Gdiplus::FontStyleRegular, Gdiplus::UnitPixel);
        Gdiplus::SolidBrush tileTextBrush(Gdiplus::Color(255, 240, 240, 240));

        g_tilePositionsX.clear();
        g_tilePositionsY.clear();

        for (size_t i = 0; i < g_pinnedApps.size(); ++i) {
            int row = i / tilesPerRow;
            int col = i % tilesPerRow;
            int tileX = contentX + col * tileSpacing;
            int tileRectY = 30 + row * tileSpacing - g_startMenuScrollOffset;

            g_tilePositionsX.push_back(tileX);
            g_tilePositionsY.push_back(tileRectY);

            if (tileRectY + tileSize < 0 || tileRectY > rcClient.bottom) continue;

            Gdiplus::RectF tileRect(tileX, tileRectY, tileSize, tileSize);
            Gdiplus::Color bgColor = (g_startMenuDragIndex == (int)i) ?
                Gdiplus::Color(100, 120, 180, 240) : Gdiplus::Color(50, 100, 150, 200);
            Gdiplus::SolidBrush tileBgBrush(bgColor);
            g.FillRectangle(&tileBgBrush, tileRect);

            Gdiplus::Pen borderPen(Gdiplus::Color(100, 150, 200, 250), 1);
            g.DrawRectangle(&borderPen, tileRect);

            HICON hIcon = NULL;
            SHFILEINFOW sfi = {};
            if (SHGetFileInfoW(g_pinnedApps[i].c_str(), 0, &sfi, sizeof(sfi), SHGFI_ICON | SHGFI_LARGEICON)) {
                hIcon = sfi.hIcon;
            }

            if (hIcon) {
                DrawIconWithGDIPlus(&g, hIcon, tileX + 10, tileRectY + 10, 60);
                DestroyIcon(hIcon);
            } else {
                Gdiplus::Color colors[] = {
                    Gdiplus::Color(200, 52, 152, 219),
                    Gdiplus::Color(200, 230, 126, 34),
                    Gdiplus::Color(200, 46, 204, 113),
                    Gdiplus::Color(200, 231, 76, 60),
                    Gdiplus::Color(200, 155, 89, 182),
                    Gdiplus::Color(200, 22, 160, 133)
                };
                Gdiplus::SolidBrush colorBrush(colors[i % 6]);
                g.FillRectangle(&colorBrush, Gdiplus::RectF(tileX + 10, tileRectY + 10, 60, 60));
            }

            std::wstring appName = g_pinnedApps[i];
            size_t lastSlash = appName.find_last_of(L"\\");
            if (lastSlash != std::wstring::npos) {
                appName = appName.substr(lastSlash + 1);
            }
            size_t dotPos = appName.find_last_of(L".");
            if (dotPos != std::wstring::npos) {
                appName = appName.substr(0, dotPos);
            }

            Gdiplus::RectF textRect(tileX, tileRectY + 75, tileSize, 20);
            Gdiplus::StringFormat format;
            format.SetAlignment(Gdiplus::StringAlignmentCenter);
            format.SetLineAlignment(Gdiplus::StringAlignmentCenter);
            format.SetFormatFlags(Gdiplus::StringFormatFlagsNoWrap);
            g.DrawString(appName.c_str(), -1, &appNameFont, textRect, &format, &tileTextBrush);
        }

        BitBlt(hdc, 0, 0, rcClient.right, rcClient.bottom, hdcMem, 0, 0, SRCCOPY);
        SelectObject(hdcMem, hbmOld);
        DeleteObject(hbmMem);
        DeleteDC(hdcMem);

        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_MOUSEWHEEL: {
        int clickX = LOWORD(lParam);
        int clickY = HIWORD(lParam);
        int zDelta = GET_WHEEL_DELTA_WPARAM(wParam);
        int scrollAmount = 36;
        int sidebarWidth = 350;

        if (clickX < sidebarWidth) {
            if (zDelta > 0) {
                g_sidebarScrollOffset -= scrollAmount;
                if (g_sidebarScrollOffset < 0) g_sidebarScrollOffset = 0;
            } else {
                g_sidebarScrollOffset += scrollAmount;
            }
            InvalidateRect(hwnd, NULL, FALSE);
            return 0;
        }

        RECT rcClient;
        GetClientRect(hwnd, &rcClient);
        int contentWidth = rcClient.right - sidebarWidth - 40;
        int tilesPerRow = contentWidth / 120;
        if (tilesPerRow < 1) tilesPerRow = 1;
        int totalTiles = g_pinnedApps.size();
        int totalRows = (totalTiles + tilesPerRow - 1) / tilesPerRow;
        int totalHeight = totalRows * 120;
        int contentHeight = rcClient.bottom - 60;
        int maxScroll = (totalHeight > contentHeight) ? (totalHeight - contentHeight) : 0;

        if (zDelta > 0) {
            g_startMenuScrollOffset -= scrollAmount;
            if (g_startMenuScrollOffset < 0) g_startMenuScrollOffset = 0;
        } else {
            g_startMenuScrollOffset += scrollAmount;
            if (g_startMenuScrollOffset > maxScroll) g_startMenuScrollOffset = maxScroll;
        }
        InvalidateRect(hwnd, NULL, FALSE);
        return 0;
    }
    case WM_LBUTTONDOWN: {
        int clickX = LOWORD(lParam);
        int clickY = HIWORD(lParam);
        int sidebarWidth = 350;

        if (clickX < sidebarWidth) {
            RECT rcClient;
            GetClientRect(hwnd, &rcClient);
            
            // Check if clicking on power buttons
            int powerButtonY = rcClient.bottom - 60;
            int powerButtonWidth = 100;
            int powerButtonHeight = 40;
            int powerButtonSpacing = 10;
            int totalButtonWidth = (powerButtonWidth * 3) + (powerButtonSpacing * 2);
            int powerButtonStartX = (sidebarWidth - totalButtonWidth) / 2;

            // Shutdown button
            if (clickX >= powerButtonStartX && clickX < powerButtonStartX + powerButtonWidth &&
                clickY >= powerButtonY && clickY < powerButtonY + powerButtonHeight) {
                // Shutdown the computer
                system("shutdown /s /t 0");
                return 0;
            }

            // Restart button
            int restartX = powerButtonStartX + powerButtonWidth + powerButtonSpacing;
            if (clickX >= restartX && clickX < restartX + powerButtonWidth &&
                clickY >= powerButtonY && clickY < powerButtonY + powerButtonHeight) {
                // Restart the computer
                system("shutdown /r /t 0");
                return 0;
            }

            // Sleep button
            int sleepX = powerButtonStartX + (powerButtonWidth + powerButtonSpacing) * 2;
            if (clickX >= sleepX && clickX < sleepX + powerButtonWidth &&
                clickY >= powerButtonY && clickY < powerButtonY + powerButtonHeight) {
                // Put computer to sleep
                SetSuspendState(FALSE, FALSE, FALSE);
                return 0;
            }

            // Handle sidebar item clicks
            for (const auto& item : g_sidebarItems) {
                if (clickY >= item.yPos && clickY < item.yPos + item.height) {
                    if (item.name == L"..") {
                        size_t lastSlash = g_currentSidebarFolder.find_last_of(L"\\");
                        if (lastSlash != std::wstring::npos) {
                            g_currentSidebarFolder = g_currentSidebarFolder.substr(0, lastSlash);
                        }
                    } else if (item.isFolder) {
                        g_currentSidebarFolder = item.fullPath;
                    } else {
                        LaunchApp(item.fullPath.c_str());
                    }
                    InvalidateRect(hwnd, NULL, FALSE);
                    return 0;
                }
            }
            return 0;
        }

        RECT rcClient;
        GetClientRect(hwnd, &rcClient);
        int contentX = sidebarWidth + 20;
        int contentWidth = rcClient.right - contentX - 20;
        int tileSize = 100;
        int tileSpacing = 120;
        int tilesPerRow = contentWidth / tileSpacing;
        if (tilesPerRow < 1) tilesPerRow = 1;

        for (size_t i = 0; i < g_pinnedApps.size(); ++i) {
            int row = i / tilesPerRow;
            int col = i % tilesPerRow;
            int tileX = contentX + col * tileSpacing;
            int tileY = 30 + row * tileSpacing - g_startMenuScrollOffset;

            if (clickX >= tileX && clickX < tileX + tileSize &&
                clickY >= tileY && clickY < tileY + tileSize) {
                LaunchApp(g_pinnedApps[i].c_str());
                return 0;
            }
        }
        return 0;
    }
    case WM_RBUTTONDOWN: {
        int clickX = LOWORD(lParam);
        int clickY = HIWORD(lParam);
        int sidebarWidth = 350;

        if (clickX < sidebarWidth) {
            for (const auto& item : g_sidebarItems) {
                if (clickY >= item.yPos && clickY < item.yPos + item.height && !item.isFolder && item.name != L"..") {
                    HMENU hMenu = CreatePopupMenu();

                    bool inStartMenu = false;
                    bool inTaskbar = false;

                    for (const auto& app : g_pinnedApps) {
                        if (_wcsicmp(app.c_str(), item.fullPath.c_str()) == 0) {
                            inStartMenu = true;
                            break;
                        }
                    }

                    for (const auto& app : g_taskbarPinned) {
                        if (_wcsicmp(app.c_str(), item.fullPath.c_str()) == 0) {
                            inTaskbar = true;
                            break;
                        }
                    }

                    if (inStartMenu) {
                        AppendMenuW(hMenu, MF_STRING, 51000, L"Remove from Start Menu");
                    } else {
                        AppendMenuW(hMenu, MF_STRING, 51001, L"Pin to Start Menu");
                    }

                    AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);

                    if (inTaskbar) {
                        AppendMenuW(hMenu, MF_STRING, 51002, L"Unpin from Taskbar");
                    } else {
                        AppendMenuW(hMenu, MF_STRING, 51003, L"Pin to Taskbar");
                    }

                    POINT pt;
                    GetCursorPos(&pt);

                    int cmd = TrackPopupMenu(hMenu, TPM_RETURNCMD, pt.x, pt.y, 0, hwnd, NULL);

                    if (cmd == 51000) {
                        g_pinnedApps.erase(
                            std::remove_if(g_pinnedApps.begin(), g_pinnedApps.end(),
                                [&item](const std::wstring& app) {
                                    return _wcsicmp(app.c_str(), item.fullPath.c_str()) == 0;
                                }),
                            g_pinnedApps.end()
                        );
                        SaveStartMenuPositions();
                        InvalidateRect(hwnd, NULL, FALSE);
                    } else if (cmd == 51001) {
                        g_pinnedApps.push_back(item.fullPath);
                        SaveStartMenuPositions();
                        InvalidateRect(hwnd, NULL, FALSE);
                    } else if (cmd == 51002) {
                        g_taskbarPinned.erase(
                            std::remove_if(g_taskbarPinned.begin(), g_taskbarPinned.end(),
                                [&item](const std::wstring& app) {
                                    return _wcsicmp(app.c_str(), item.fullPath.c_str()) == 0;
                                }),
                            g_taskbarPinned.end()
                        );
                        SaveTaskbarConfig();
                        PopulateTaskbarIcons();
                        InvalidateRect(g_hTaskbarWnd, NULL, TRUE);
                    } else if (cmd == 51003) {
                        g_taskbarPinned.push_back(item.fullPath);
                        SaveTaskbarConfig();
                        PopulateTaskbarIcons();
                        InvalidateRect(g_hTaskbarWnd, NULL, TRUE);
                    }

                    DestroyMenu(hMenu);
                    return 0;
                }
            }
            return 0;
        }

        RECT rcClient;
        GetClientRect(hwnd, &rcClient);
        int contentX = sidebarWidth + 20;
        int contentWidth = rcClient.right - contentX - 20;
        int tileSize = 100;
        int tileSpacing = 120;
        int tilesPerRow = contentWidth / tileSpacing;
        if (tilesPerRow < 1) tilesPerRow = 1;

        for (size_t i = 0; i < g_pinnedApps.size(); ++i) {
            int row = i / tilesPerRow;
            int col = i % tilesPerRow;
            int tileX = contentX + col * tileSpacing;
            int tileY = 30 + row * tileSpacing - g_startMenuScrollOffset;

            if (clickX >= tileX && clickX < tileX + tileSize &&
                clickY >= tileY && clickY < tileY + tileSize) {

                HMENU hMenu = CreatePopupMenu();

                bool alreadyPinned = std::find(g_taskbarPinned.begin(), g_taskbarPinned.end(),
                    g_pinnedApps[i]) != g_taskbarPinned.end();

                if (alreadyPinned) {
                    AppendMenuW(hMenu, MF_STRING, 50000, L"Unpin from taskbar");
                } else {
                    AppendMenuW(hMenu, MF_STRING, 50000, L"Pin to taskbar");
                }

                AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
                AppendMenuW(hMenu, MF_STRING, 50001, L"Remove from Start");

                POINT pt;
                GetCursorPos(&pt);

                int cmd = TrackPopupMenu(hMenu, TPM_RETURNCMD, pt.x, pt.y, 0, hwnd, NULL);

                if (cmd == 50000) {
                    if (alreadyPinned) {
                        g_taskbarPinned.erase(
                            std::remove(g_taskbarPinned.begin(), g_taskbarPinned.end(), g_pinnedApps[i]),
                            g_taskbarPinned.end()
                        );
                    } else {
                        g_taskbarPinned.push_back(g_pinnedApps[i]);
                    }
                    SaveTaskbarConfig();
                    PopulateTaskbarIcons();
                    InvalidateRect(g_hTaskbarWnd, NULL, TRUE);
                } else if (cmd == 50001) {
                    g_pinnedApps.erase(g_pinnedApps.begin() + i);
                    SaveStartMenuPositions();
                    InvalidateRect(hwnd, NULL, FALSE);
                }

                DestroyMenu(hMenu);
                return 0;
            }
        }
        return 0;
    }
    case WM_MOUSEMOVE: {
        if (g_startMenuDragging && g_startMenuDragIndex >= 0) {
            InvalidateRect(hwnd, NULL, FALSE);
        }
        return 0;
    }
    case WM_LBUTTONUP: {
        if (g_startMenuDragging && g_startMenuDragIndex >= 0) {
            g_startMenuDragging = false;
            g_startMenuDragIndex = -1;
            InvalidateRect(hwnd, NULL, FALSE);
        }
        return 0;
    }
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

void DestroyTaskbar() {
    if (g_hTaskbarWnd && IsWindow(g_hTaskbarWnd)) {
        // Kill the timer first
        KillTimer(g_hTaskbarWnd, 1);
        // Destroy the window
        DestroyWindow(g_hTaskbarWnd);
        g_hTaskbarWnd = NULL;
        
        // Resize desktop to take full screen
        ResizeShellWindows();
    }
}

void RecreateTaskbar() {
    if (!g_hTaskbarWnd) {
        // Recreate taskbar window
        g_hTaskbarWnd = CreateWindowEx(
            WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE,
            L"CustomTaskbarClass",
            L"Taskbar",
            WS_POPUP,
            0, 0, 0, 0,
            NULL,
            NULL,
            g_hInst,
            NULL
        );
        
        if (g_hTaskbarWnd) {
            // Restore timer
            SetTimer(g_hTaskbarWnd, 1, 500, NULL);
            
            // Normal z-order - games can go on top
            SetWindowPos(g_hTaskbarWnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
            
            // Resize to proper position
            ResizeShellWindows();
            
            // Show the taskbar
            ShowWindow(g_hTaskbarWnd, SW_SHOW);
            
            // Force redraw
            InvalidateRect(g_hTaskbarWnd, NULL, TRUE);
        }
    }
}

void ResizeShellWindows(int dpi) {
    int w = GetSystemMetrics(SM_CXSCREEN);
    int h = GetSystemMetrics(SM_CYSCREEN);

    int tb = g_taskbarHeight * std::max(1, dpi / 96);

    // Desktop full width - adjust based on taskbar visibility
    if (g_hTaskbarWnd && IsWindow(g_hTaskbarWnd)) {
        SetWindowPos(g_hDesktopWnd, HWND_BOTTOM, 0, 0, w, h - tb, SWP_NOACTIVATE);
        
        // Taskbar at bottom - normal z-order so games can cover it
        SetWindowPos(g_hTaskbarWnd, HWND_BOTTOM, 0, h - tb, w, tb, SWP_NOACTIVATE | SWP_SHOWWINDOW);
    } else {
        // No taskbar - desktop takes full screen
        SetWindowPos(g_hDesktopWnd, HWND_BOTTOM, 0, 0, w, h, SWP_NOACTIVATE);
    }

    // Start menu on left side, 45% width with top margin only
    int smW = (w * 45) / 100;
    int smY = 50; // 50px from top
    int smH = h - (g_hTaskbarWnd && IsWindow(g_hTaskbarWnd) ? tb : 0) - smY; // Fill to taskbar
    SetWindowPos(g_hStartMenuWnd, HWND_TOP, 0, smY, smW, smH, SWP_NOZORDER);
}

// Low-level keyboard hook to capture Windows key
LRESULT CALLBACK LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode == HC_ACTION) {
        KBDLLHOOKSTRUCT* pKeyboard = (KBDLLHOOKSTRUCT*)lParam;
        
        // Check for Windows key (left or right)
        if (pKeyboard->vkCode == VK_LWIN || pKeyboard->vkCode == VK_RWIN) {
            if (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN) {
                // Only trigger on key down, not repeat
                if (!g_winKeyPressed) {
                    g_winKeyPressed = true;
                    
                    // Toggle taskbar and start menu
                    bool startMenuCurrentlyVisible = IsWindowVisible(g_hStartMenuWnd);
                    
                    if (!startMenuCurrentlyVisible) {
                        // Opening start menu - save current taskbar state and show both
                        g_taskbarVisibleBeforeWinKey = g_taskbarVisible;
                        
                        // Make sure taskbar is visible and on top
                        if (!g_taskbarVisible) {
                            RecreateTaskbar();
                            g_taskbarVisible = true;
                        }
                        SetWindowPos(g_hTaskbarWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
                        
                        // Show start menu
                        ShowWindow(g_hStartMenuWnd, SW_SHOW);
                        SetForegroundWindow(g_hStartMenuWnd);
                        g_startMenuOpenedByWinKey = true;
                    } else {
                        // Closing start menu - restore taskbar to previous state
                        ShowWindow(g_hStartMenuWnd, SW_HIDE);
                        
                        // Restore taskbar z-order to normal (not topmost)
                        SetWindowPos(g_hTaskbarWnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
                        
                        // Restore taskbar visibility to state before Windows key was pressed
                        if (!g_taskbarVisibleBeforeWinKey && g_taskbarVisible) {
                            DestroyTaskbar();
                            g_taskbarVisible = false;
                        }
                        g_startMenuOpenedByWinKey = false;
                    }
                    
                    // Block the Windows key from reaching Windows
                    return 1;
                }
            } else if (wParam == WM_KEYUP || wParam == WM_SYSKEYUP) {
                g_winKeyPressed = false;
                return 1; // Block key up as well
            }
        }
    }
    
    return CallNextHookEx(g_keyboardHook, nCode, wParam, lParam);
}

// Register taskbar as AppBar
void RegisterAppBar(HWND hwnd)
{
    APPBARDATA abd = { 0 };
    abd.cbSize = sizeof(APPBARDATA);
    abd.hWnd = hwnd;
    abd.uCallbackMessage = APPBAR_CALLBACK;
    SHAppBarMessage(ABM_NEW, &abd);
}

// Set AppBar position
void SetAppBarPos(HWND hwnd)
{
    RECT rc;
    GetWindowRect(hwnd, &rc);
    APPBARDATA abd = { 0 };
    abd.cbSize = sizeof(APPBARDATA);
    abd.hWnd = hwnd;
    abd.uEdge = ABE_BOTTOM;
    abd.rc.left   = 0;
    abd.rc.right  = GetSystemMetrics(SM_CXSCREEN);
    abd.rc.bottom = GetSystemMetrics(SM_CYSCREEN);
    abd.rc.top    = abd.rc.bottom - (rc.bottom - rc.top);
    SHAppBarMessage(ABM_QUERYPOS, &abd);
    SHAppBarMessage(ABM_SETPOS, &abd);
    MoveWindow(hwnd,
        abd.rc.left,
        abd.rc.top,
        abd.rc.right - abd.rc.left,
        abd.rc.bottom - abd.rc.top,
        TRUE);
    
    // Also manually set the work area
    RECT workArea;
    workArea.left = 0;
    workArea.top = 0;
    workArea.right = abd.rc.right;
    workArea.bottom = abd.rc.top; // Top of taskbar = bottom of work area
    SystemParametersInfoW(SPI_SETWORKAREA, 0, &workArea, SPIF_SENDCHANGE | SPIF_UPDATEINIFILE);
    
    // Broadcast the change to all top-level windows
    SendMessageTimeoutW(HWND_BROADCAST, WM_SETTINGCHANGE, SPI_SETWORKAREA, 0, SMTO_ABORTIFHUNG, 1000, NULL);
}

// Unregister AppBar
void UnregisterAppBar(HWND hwnd)
{
    APPBARDATA abd = { 0 };
    abd.cbSize = sizeof(APPBARDATA);
    abd.hWnd = hwnd;
    SHAppBarMessage(ABM_REMOVE, &abd);
    
    // Restore work area to full screen
    int w = GetSystemMetrics(SM_CXSCREEN);
    int h = GetSystemMetrics(SM_CYSCREEN);
    RECT workArea = {0, 0, w, h};
    SystemParametersInfoW(SPI_SETWORKAREA, 0, &workArea, SPIF_SENDCHANGE);
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, PWSTR, int) {
    // Hide console window if it exists
    HWND hConsole = GetConsoleWindow();
    if (hConsole) {
        ShowWindow(hConsole, SW_HIDE);
    }

    // HIDE THE REAL WINDOWS TASKBAR
    HWND hRealTaskbar = FindWindowW(L"Shell_TrayWnd", NULL);
    if (hRealTaskbar) {
        ShowWindow(hRealTaskbar, SW_HIDE);
    }

    // HIDE THE START MENU BUTTON AREA
    HWND hStartButton = FindWindowW(L"Button", L"Start");
    if (hStartButton) {
        ShowWindow(hStartButton, SW_HIDE);
    }

    g_hInst = hInstance;

    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

    // Initialize GDI+
    Gdiplus::GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);

    INITCOMMONCONTROLSEX icex = {};
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_WIN95_CLASSES;
    InitCommonControlsEx(&icex);

    HMODULE user32 = LoadLibraryW(L"user32.dll");
    if (user32) {
        typedef BOOL (WINAPI *SetProcessDpiAwarenessContextFn)(DPI_AWARENESS_CONTEXT);
        SetProcessDpiAwarenessContextFn f = (SetProcessDpiAwarenessContextFn) GetProcAddress(user32, "SetProcessDpiAwarenessContext");
        if (f) {
            f(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
        } else {
            SetProcessDPIAware();
        }
        FreeLibrary(user32);
    }

    WNDCLASS wc = {};
    wc.lpfnWndProc = DesktopWndProc;
    wc.style = CS_DBLCLKS;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"CustomDesktopClass";
    wc.hbrBackground = NULL;
    RegisterClass(&wc);

    wc = {};
    wc.lpfnWndProc = TaskbarWndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"CustomTaskbarClass";
    wc.hbrBackground = (HBRUSH) CreateSolidBrush(RGB(45,45,48));
    RegisterClass(&wc);

    wc = {};
    wc.lpfnWndProc = StartMenuWndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"CustomStartMenuClass";
    wc.hbrBackground = (HBRUSH) CreateSolidBrush(RGB(40,40,40));
    RegisterClass(&wc);

    // FIXED: Create desktop as regular window
    g_hDesktopWnd = CreateWindowEx(
        WS_EX_TOOLWINDOW,
        L"CustomDesktopClass",
        L"Desktop",
        WS_POPUP,
        0,0,0,0,
        NULL,
        NULL,
        hInstance,
        NULL
    );
    
    // Store the real DesktopWndProc address for safe restoration
    g_realDesktopProc = (WNDPROC)GetWindowLongPtrW(g_hDesktopWnd, GWLP_WNDPROC);

    // Taskbar as normal window (can be covered by fullscreen apps)
    g_hTaskbarWnd = CreateWindowEx(
        WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE,
        L"CustomTaskbarClass",
        L"Taskbar",
        WS_POPUP,
        0,0,0,0,
        NULL,
        NULL,
        hInstance,
        NULL
    );

    g_hStartMenuWnd = CreateWindowEx(
        WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE,
        L"CustomStartMenuClass",
        L"StartMenu",
        WS_POPUP,
        10,10,300,300,
        NULL,  // NO parent
        NULL,
        hInstance,
        NULL
    );

    // Ensure desktop is behind everything
    SetWindowPos(g_hDesktopWnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);

    // Load window positions from config
    LoadWindowPositions();

    ShowWindow(g_hDesktopWnd, SW_SHOW);
    ShowWindow(g_hTaskbarWnd, SW_SHOW);
    ShowWindow(g_hStartMenuWnd, SW_HIDE);

    // Set timer to refresh taskbar for running app indicators (every 500ms)
    SetTimer(g_hTaskbarWnd, 1, 500, NULL);

    // Register taskbar as AppBar and set position
    RegisterAppBar(g_hTaskbarWnd);
    SetAppBarPos(g_hTaskbarWnd);

    ResizeShellWindows();

    PopulateDesktopIcons();

    // Load Start Menu apps from Windows
    PopulateStartMenuApps();

    // Clean up explorer and shell from lists
    CleanupExplorerAndShellFromLists();

    // Load taskbar configuration and icons
    LoadTaskbarConfig();
    PopulateTaskbarIcons();
    
    // Initialize system tray icons
    PopulateRealTrayIcons();

    // apply default view mode (large)
    UpdateViewMode(g_viewMode);

    if (!RegisterHotKey(NULL, 9001, MOD_CONTROL | MOD_ALT, 'E')) {
        // Failed to register Ctrl+Alt+E
    }

    // Register Ctrl + Alt + T hotkey for toggling taskbar visibility (Windows+T is reserved by Windows)
    if (!RegisterHotKey(NULL, 9002, MOD_CONTROL | MOD_ALT, 'T')) {
        // Failed to register Ctrl+Alt+T
    }

    // Install low-level keyboard hook to capture Windows key
    g_keyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, LowLevelKeyboardProc, hInstance, 0);
    if (!g_keyboardHook) {
        // Failed to install keyboard hook
    }

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        if (msg.message == WM_HOTKEY && msg.wParam == 9001) {
            LaunchExplorer();
        }
        else if (msg.message == WM_HOTKEY && msg.wParam == 9002) {
            // Toggle taskbar by destroying/recreating it
            g_taskbarVisible = !g_taskbarVisible;
            if (g_taskbarVisible) {
                RecreateTaskbar();
            } else {
                DestroyTaskbar();
            }
        }
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    UnregisterHotKey(NULL, 9001);
    UnregisterHotKey(NULL, 9002); // Unregister Ctrl+Alt+T hotkey
    
    // Unhook keyboard hook
    if (g_keyboardHook) {
        UnhookWindowsHookEx(g_keyboardHook);
        g_keyboardHook = NULL;
    }

    // Unregister AppBar
    if (g_hTaskbarWnd) {
        UnregisterAppBar(g_hTaskbarWnd);
    }

    // Save window positions before closing
    SaveTaskbarConfig();

    // Save Start Menu positions
    SaveStartMenuPositions();

    // Clean up tray icons
    for (auto& icon : g_trayIcons) {
        if (icon.hIcon && !icon.isRealIcon) {
            DestroyIcon(icon.hIcon);
        }
    }
    g_trayIcons.clear();
    
    for (auto& icon : g_realTrayIcons) {
        if (icon.hIcon) {
            DestroyIcon(icon.hIcon);
        }
    }
    g_realTrayIcons.clear();

    // SHOW THE REAL TASKBAR AGAIN ON EXIT
    if (hRealTaskbar) {
        ShowWindow(hRealTaskbar, SW_SHOW);
    }

    // SHOW START BUTTON AGAIN
    if (hStartButton) {
        ShowWindow(hStartButton, SW_SHOW);
    }

    for (auto &it : g_desktopItems) {
        if (it.hIcon) {
            DestroyIcon(it.hIcon);
            it.hIcon = NULL;
        }
    }

    // Clean up background image
    if (g_backgroundImage) {
        delete g_backgroundImage;
        g_backgroundImage = NULL;
    }

    Gdiplus::GdiplusShutdown(gdiplusToken);

    CoUninitialize();

    return 0;
}

// Show the native Explorer context menu for a file path at screen point `pt`.
void ShowShellContextMenu(HWND hwnd, POINT pt, LPCWSTR pszPath) {
    PIDLIST_ABSOLUTE pidl = NULL;
    HRESULT hr = SHParseDisplayName(pszPath, NULL, &pidl, 0, NULL);
    if (SUCCEEDED(hr) && pidl) {
        IShellFolder* psfParent = NULL;
        PCUITEMID_CHILD pidlChild = NULL;
        hr = SHBindToParent(pidl, IID_IShellFolder, (void**)&psfParent, &pidlChild);
        if (SUCCEEDED(hr) && psfParent && pidlChild) {
            IContextMenu* pcm = NULL;
            hr = psfParent->GetUIObjectOf(hwnd, 1, &pidlChild, IID_IContextMenu, NULL, (void**)&pcm);
            if (SUCCEEDED(hr) && pcm) {
                IContextMenu2* pcm2 = NULL;
                IContextMenu3* pcm3 = NULL;
                pcm->QueryInterface(IID_IContextMenu2, (void**)&pcm2);
                pcm->QueryInterface(IID_IContextMenu3, (void**)&pcm3);

                IContextMenu2* prev_pcm2 = g_pcm2;
                IContextMenu3* prev_pcm3 = g_pcm3;
                WNDPROC prev_proc = g_origDesktopProc;

                g_pcm2 = pcm2;
                g_pcm3 = pcm3;
                
                // Only subclass if not already subclassed
                WNDPROC currentProc = (WNDPROC)GetWindowLongPtrW(hwnd, GWLP_WNDPROC);
                if (currentProc != Desktop_SubclassProc) {
                    g_origDesktopProc = (WNDPROC)SetWindowLongPtrW(hwnd, GWLP_WNDPROC, (LONG_PTR)Desktop_SubclassProc);
                }

                HMENU hMenu = CreatePopupMenu();
                if (hMenu) {
                    hr = pcm->QueryContextMenu(hMenu, 0, 1, 0x7fff, CMF_NORMAL);
                    if (SUCCEEDED(hr)) {
                        int baseId = 1;
                        int count = GetMenuItemCount(hMenu);
                        UINT refreshId = (UINT)(baseId + count);
                        InsertMenuW(hMenu, -1, MF_BYPOSITION, refreshId, L"Refresh");

                        int id = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, 0, hwnd, NULL);
                        if (id) {
                            if ((UINT)id == refreshId) {
                                PopulateDesktopIcons();
                                InvalidateRect(hwnd, NULL, TRUE);
                            } else {
                                CMINVOKECOMMANDINFOEX ici = {};
                                ici.cbSize = sizeof(ici);
                                ici.fMask = CMIC_MASK_UNICODE;
                                ici.hwnd = hwnd;
                                ici.lpVerb = MAKEINTRESOURCEA(id - baseId);
                                ici.lpVerbW = MAKEINTRESOURCEW(id - baseId);
                                ici.nShow = SW_SHOWNORMAL;
                                pcm->InvokeCommand((LPCMINVOKECOMMANDINFO)&ici);
                                // Don't auto-refresh - let user manually refresh if needed
                            }
                        }
                    }
                    DestroyMenu(hMenu);
                }

                // Always restore to the real desktop window proc
                SetWindowLongPtrW(hwnd, GWLP_WNDPROC, (LONG_PTR)g_realDesktopProc);
                g_origDesktopProc = NULL;

                if (g_pcm3) { g_pcm3->Release(); g_pcm3 = NULL; }
                if (g_pcm2) { g_pcm2->Release(); g_pcm2 = NULL; }

                g_pcm2 = prev_pcm2;
                g_pcm3 = prev_pcm3;

                pcm->Release();
            }
            psfParent->Release();
        }
        CoTaskMemFree(pidl);
    }
}

// Show the native Explorer context menu for the desktop background at screen point `pt`.
void ShowDesktopContextMenu(HWND hwnd, POINT pt) {
    IShellFolder* psfDesktop = NULL;
    if (SUCCEEDED(SHGetDesktopFolder(&psfDesktop)) && psfDesktop) {
        IContextMenu* pcm = NULL;
        HRESULT hr = psfDesktop->GetUIObjectOf(hwnd, 0, NULL, IID_IContextMenu, NULL, (void**)&pcm);
        if (SUCCEEDED(hr) && pcm) {
            pcm->QueryInterface(IID_IContextMenu2, (void**)&g_pcm2);
            pcm->QueryInterface(IID_IContextMenu3, (void**)&g_pcm3);

            // Only subclass if not already subclassed
            WNDPROC currentProc = (WNDPROC)GetWindowLongPtrW(hwnd, GWLP_WNDPROC);
            if (currentProc != Desktop_SubclassProc) {
                g_origDesktopProc = (WNDPROC)SetWindowLongPtrW(hwnd, GWLP_WNDPROC, (LONG_PTR)Desktop_SubclassProc);
            }

            HMENU hMenu = CreatePopupMenu();
            if (hMenu) {
                hr = pcm->QueryContextMenu(hMenu, 0, 1, 0x7fff, CMF_NORMAL | CMF_EXPLORE);
                if (SUCCEEDED(hr)) {

                    AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
                    AppendMenuW(hMenu, MF_STRING, IDM_DESKTOP_REFRESH, L"Refresh");

                    HMENU hView = CreatePopupMenu();
                    AppendMenuW(hView, MF_STRING | (g_viewMode == 2 ? MF_CHECKED : 0), IDM_VIEW_LARGE, L"Large icons");
                    AppendMenuW(hView, MF_STRING | (g_viewMode == 1 ? MF_CHECKED : 0), IDM_VIEW_MEDIUM, L"Medium icons");
                    AppendMenuW(hView, MF_STRING | (g_viewMode == 0 ? MF_CHECKED : 0), IDM_VIEW_SMALL, L"Small icons");
                    AppendMenuW(hMenu, MF_POPUP, (UINT_PTR)hView, L"View");

                    HMENU hSort = CreatePopupMenu();
                    AppendMenuW(hSort, MF_STRING, IDM_SORT_NAME, L"Name");
                    AppendMenuW(hSort, MF_STRING, IDM_SORT_TYPE, L"Type");
                    AppendMenuW(hSort, MF_STRING, IDM_SORT_DATE, L"Date modified");
                    AppendMenuW(hMenu, MF_POPUP, (UINT_PTR)hSort, L"Sort by");

                    AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
                    AppendMenuW(hMenu, MF_STRING | (g_showIcons ? MF_CHECKED : 0), IDM_TOGGLE_SHOW_ICONS, L"Show desktop icons");
                    AppendMenuW(hMenu, MF_STRING | (g_autoArrange ? MF_CHECKED : 0), IDM_AUTO_ARRANGE, L"Auto arrange icons");
                    AppendMenuW(hMenu, MF_STRING, IDM_ALIGN_TO_GRID, L"Align icons to grid");
                    AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
                    AppendMenuW(hMenu, MF_STRING, IDM_CHANGE_TASKBAR_COLOR, L"Change Taskbar Color...");
                    AppendMenuW(hMenu, MF_STRING, IDM_CHANGE_STARTMENU_COLOR, L"Change Start Menu Color...");
                    AppendMenuW(hMenu, MF_STRING, IDM_CHANGE_BACKGROUND, L"Change Background...");

                    SetForegroundWindow(hwnd);

                    int cmd = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, 0, hwnd, NULL);
                    PostMessage(hwnd, WM_NULL, 0, 0);
                    if (cmd) {
                        if (cmd == IDM_DESKTOP_REFRESH) {
                            PopulateDesktopIcons();
                            ResetDesktopPositions();
                            InvalidateRect(hwnd, NULL, TRUE);
                        } else if (cmd == IDM_VIEW_LARGE) {
                            UpdateViewMode(2);
                        } else if (cmd == IDM_VIEW_MEDIUM) {
                            UpdateViewMode(1);
                        } else if (cmd == IDM_VIEW_SMALL) {
                            UpdateViewMode(0);
                        } else if (cmd == IDM_SORT_NAME) {
                            SortByName();
                            InvalidateRect(hwnd, NULL, TRUE);
                        } else if (cmd == IDM_SORT_TYPE) {
                            SortByTypeThenName();
                            InvalidateRect(hwnd, NULL, TRUE);
                        } else if (cmd == IDM_SORT_DATE) {
                            SortByDateDesc();
                            InvalidateRect(hwnd, NULL, TRUE);
                        } else if (cmd == IDM_TOGGLE_SHOW_ICONS) {
                            g_showIcons = !g_showIcons;
                            InvalidateRect(hwnd, NULL, TRUE);
                        } else if (cmd == IDM_AUTO_ARRANGE) {
                            g_autoArrange = !g_autoArrange;
                            if (g_autoArrange) SortByName();
                            InvalidateRect(hwnd, NULL, TRUE);
                        } else if (cmd == IDM_ALIGN_TO_GRID) {
                            AlignIconsToGrid();
                        } else if (cmd == IDM_CHANGE_TASKBAR_COLOR) {
                            // Choose color for taskbar
                            CHOOSECOLORW cc = {};
                            static COLORREF customColors[16] = {0};
                            
                            cc.lStructSize = sizeof(cc);
                            cc.hwndOwner = hwnd;
                            cc.lpCustColors = customColors;
                            cc.rgbResult = g_taskbarColor;
                            cc.Flags = CC_FULLOPEN | CC_RGBINIT;
                            
                            if (ChooseColorW(&cc)) {
                                g_taskbarColor = cc.rgbResult;
                                SaveTaskbarConfig();
                                InvalidateRect(g_hTaskbarWnd, NULL, TRUE);
                            }
                        } else if (cmd == IDM_CHANGE_STARTMENU_COLOR) {
                            // Choose color for start menu
                            CHOOSECOLORW cc = {};
                            static COLORREF customColors2[16] = {0};
                            
                            cc.lStructSize = sizeof(cc);
                            cc.hwndOwner = hwnd;
                            cc.lpCustColors = customColors2;
                            cc.rgbResult = g_startMenuColor;
                            cc.Flags = CC_FULLOPEN | CC_RGBINIT;
                            
                            if (ChooseColorW(&cc)) {
                                g_startMenuColor = cc.rgbResult;
                                SaveTaskbarConfig();
                                InvalidateRect(g_hStartMenuWnd, NULL, TRUE);
                            }
                        } else if (cmd == IDM_CHANGE_BACKGROUND) {
                            // Open file dialog to select background image
                            OPENFILENAMEW ofn = {};
                            wchar_t szFile[MAX_PATH] = L"";
                            
                            ofn.lStructSize = sizeof(ofn);
                            ofn.hwndOwner = hwnd;
                            ofn.lpstrFile = szFile;
                            ofn.nMaxFile = MAX_PATH;
                            ofn.lpstrFilter = L"Image Files\0*.jpg;*.jpeg;*.png;*.bmp;*.gif\0All Files\0*.*\0";
                            ofn.nFilterIndex = 1;
                            ofn.lpstrTitle = L"Select Background Image";
                            ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;
                            
                            if (GetOpenFileNameW(&ofn)) {
                                // Load the new background image
                                if (g_backgroundImage) {
                                    delete g_backgroundImage;
                                    g_backgroundImage = NULL;
                                }
                                
                                g_backgroundImage = Gdiplus::Bitmap::FromFile(szFile);
                                if (g_backgroundImage && g_backgroundImage->GetLastStatus() == Gdiplus::Ok) {
                                    g_backgroundPath = szFile;
                                    SaveTaskbarConfig();
                                    InvalidateRect(hwnd, NULL, TRUE);
                                } else {
                                    if (g_backgroundImage) {
                                        delete g_backgroundImage;
                                        g_backgroundImage = NULL;
                                    }
                                    MessageBoxW(hwnd, L"Failed to load image file.", L"Error", MB_OK | MB_ICONERROR);
                                }
                            }
                        } else {
                            CMINVOKECOMMANDINFOEX ici = {};
                            ici.cbSize = sizeof(ici);
                            ici.fMask = CMIC_MASK_UNICODE;
                            ici.hwnd = hwnd;
                            ici.lpVerb = MAKEINTRESOURCEA(cmd - 1);
                            ici.lpVerbW = MAKEINTRESOURCEW(cmd - 1);
                            ici.nShow = SW_SHOWNORMAL;
                            pcm->InvokeCommand((LPCMINVOKECOMMANDINFO)&ici);
                            // Don't auto-refresh - let user manually refresh if needed
                        }
                    }

                    DestroyMenu(hView);
                    DestroyMenu(hSort);
                }
                DestroyMenu(hMenu);
            }

            // Always restore to the real desktop window proc
            SetWindowLongPtrW(hwnd, GWLP_WNDPROC, (LONG_PTR)g_realDesktopProc);
            g_origDesktopProc = NULL;

            if (g_pcm3) { g_pcm3->Release(); g_pcm3 = NULL; }
            if (g_pcm2) { g_pcm2->Release(); g_pcm2 = NULL; }

            pcm->Release();
        }
        psfDesktop->Release();
    }
}
