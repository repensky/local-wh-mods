// ==WindhawkMod==
// @id              discord-balloon-notifs
// @name            Discord Balloon Notifications
// @description     Converts Discord toast notifications to classic Windows balloon notifications
// @version         1.0
// @author          repensky
// @include         Discord.exe
// @include         Discord*.exe
// @compilerOptions -lole32 -loleaut32 -lruntimeobject -lshlwapi -lshell32
// ==/WindhawkMod==

// ==WindhawkModReadme==
/*
# Discord Balloon Notifications
Converts Discord's modern Windows toast notifications into balloon tips that appear from the system tray using Shell_NotifyIconW.

![Picture](https://raw.githubusercontent.com/repensky/local-wh-mods/refs/heads/main/image.png)

Features:
- Notification sender and message is shown in the balloon
- Optionally shows the sender's profile picture as an icon in the balloon
- Configurable PFP icon size (Small 16px, Medium 24px, Large 32px)
- Clicking the balloon or tray icon brings Discord to the foreground
- Tray icon automatically removed when Discord exits

**Note:** Enable this mod before starting Discord. If Discord is already
running when you enable the mod, restart Discord manually. Otherwise, Discord may **crash** if not restarted on next notification.

**Note 2:** Balloon display duration is controlled by Windows
Control Panel > Ease of Access > Use the computer without a display > How long should Windows notification dialog boxes stay open

*/
// ==/WindhawkModReadme==

// ==WindhawkModSettings==
/*
- showProfilePicture: true
  $name: Show profile pictures in balloon
  $description: Display the sender's Discord profile picture as an icon in the balloon notification
- iconSize: small
  $name: Profile picture size
  $description: Size of the profile picture icon in the balloon (only applies when profile pictures are enabled)
  $options:
  - small: Small (16x16)
  - medium: Medium (24x24)
  - large: Large (32x32)
*/
// ==/WindhawkModSettings==

#include <windows.h>
#include <windhawk_api.h>
#include <shellapi.h>
#include <roapi.h>
#include <winstring.h>
#include <tlhelp32.h>
#include <cwchar>

// Settings

enum IconSizeOption {
    ICON_SIZE_SMALL,
    ICON_SIZE_MEDIUM,
    ICON_SIZE_LARGE
};

struct {
    bool showProfilePicture;
    IconSizeOption iconSize;
} g_settings;

static void LoadSettings() {
    g_settings.showProfilePicture = Wh_GetIntSetting(L"showProfilePicture", 1) != 0;
    
    LPCWSTR iconSizeSetting = Wh_GetStringSetting(L"iconSize");
    if (wcscmp(iconSizeSetting, L"small") == 0) {
        g_settings.iconSize = ICON_SIZE_SMALL;
    } else if (wcscmp(iconSizeSetting, L"medium") == 0) {
        g_settings.iconSize = ICON_SIZE_MEDIUM;
    } else {
        g_settings.iconSize = ICON_SIZE_LARGE;
    }
    Wh_FreeStringSetting(iconSizeSetting);
    
    Wh_Log(L"Settings: showProfilePicture=%s, iconSize=%d", 
           g_settings.showProfilePicture ? L"true" : L"false",
           g_settings.iconSize);
}

// Helpers

#ifndef MIN_VAL
#define MIN_VAL(a, b) (((a) < (b)) ? (a) : (b))
#endif

// Discord Process Monitor

static DWORD g_currentPid = 0;
static HANDLE g_hMonitorThread = nullptr;
static volatile bool g_stopMonitor = false;

static bool IsAnyDiscordRunning() {
    bool found = false;
    HANDLE hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (hSnap == INVALID_HANDLE_VALUE) return true;
    
    PROCESSENTRY32W pe = {};
    pe.dwSize = sizeof(pe);
    
    if (Process32FirstW(hSnap, &pe)) {
        do {
            if (_wcsicmp(pe.szExeFile, L"Discord.exe") == 0) {
                found = true;
                break;
            }
        } while (Process32NextW(hSnap, &pe));
    }
    
    CloseHandle(hSnap);
    return found;
}

// Focus Discord on Balloon Click

struct EnumData {
    HWND result;
    DWORD pid;
};

static BOOL CALLBACK FindDiscordWindowProc(HWND h, LPARAM lp) {
    EnumData* d = (EnumData*)lp;
    DWORD pid = 0;
    GetWindowThreadProcessId(h, &pid);
    if (pid == d->pid && IsWindowVisible(h)) {
        WCHAR cls[256];
        GetClassNameW(h, cls, 256);
        if (wcsstr(cls, L"Chrome_WidgetWin") || wcsstr(cls, L"Discord")) {
            WCHAR title[256];
            if (GetWindowTextW(h, title, 256) > 0) {
                d->result = h;
                return FALSE;
            }
        }
    }
    return TRUE;
}

static void FocusDiscordWindow() {
    HWND hWnd = nullptr;
    EnumData data = { nullptr, GetCurrentProcessId() };
    EnumWindows(FindDiscordWindowProc, (LPARAM)&data);
    hWnd = data.result;
    if (!hWnd) hWnd = FindWindowW(L"Discord", nullptr);
    if (!hWnd) return;

    if (IsIconic(hWnd))
        ShowWindow(hWnd, SW_RESTORE);

    INPUT inputs[2] = {};
    inputs[0].type = INPUT_KEYBOARD;
    inputs[0].ki.wVk = VK_MENU;
    inputs[1].type = INPUT_KEYBOARD;
    inputs[1].ki.wVk = VK_MENU;
    inputs[1].ki.dwFlags = KEYEVENTF_KEYUP;
    SendInput(2, inputs, sizeof(INPUT));

    SetForegroundWindow(hWnd);
}

// Tray Constructor

#define WM_BALLOON_SHOW   (WM_USER + 300)
#define WM_TRAY_CALLBACK  (WM_USER + 200)
#define WM_CHECK_DISCORD  (WM_USER + 301)

static HWND g_hBalloonWnd = nullptr;
static bool g_iconAdded = false;
static CRITICAL_SECTION g_cs;
static HICON g_hAppIcon = nullptr;
static HANDLE g_hThread = nullptr;

static WCHAR g_pendingTitle[256] = {};
static WCHAR g_pendingBody[512] = {};

// PFP icon
static WCHAR g_lastImagePath[MAX_PATH] = {};
static HICON g_lastNotifIcon = nullptr;

static void RemoveTrayIcon();

static void EnsureTrayIcon() {
    if (g_iconAdded || !g_hBalloonWnd) return;

    NOTIFYICONDATAW nid = {};
    nid.cbSize = sizeof(nid);
    nid.hWnd = g_hBalloonWnd;
    nid.uID = 1;
    nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    nid.uCallbackMessage = WM_TRAY_CALLBACK;
    nid.hIcon = g_hAppIcon ? g_hAppIcon : LoadIconW(nullptr, IDI_APPLICATION);
    wcscpy_s(nid.szTip, L"Discord");

    if (Shell_NotifyIconW(NIM_ADD, &nid)) {
        g_iconAdded = true;
        nid.uVersion = NOTIFYICON_VERSION_4;
        Shell_NotifyIconW(NIM_SETVERSION, &nid);
    }
}

// GDI+ PNG to HICON

typedef int (__stdcall *GdiplusStartup_t)(ULONG_PTR*, const void*, void*);
typedef void (__stdcall *GdiplusShutdown_t)(ULONG_PTR);
typedef int (__stdcall *GdipCreateBitmapFromFile_t)(const WCHAR*, void**);
typedef int (__stdcall *GdipCreateHICONFromBitmap_t)(void*, HICON*);
typedef int (__stdcall *GdipDisposeImage_t)(void*);
typedef int (__stdcall *GdipGetImageWidth_t)(void*, UINT*);
typedef int (__stdcall *GdipGetImageHeight_t)(void*, UINT*);
typedef int (__stdcall *GdipCreateBitmapFromScan0_t)(INT, INT, INT, INT, BYTE*, void**);
typedef int (__stdcall *GdipGetImageGraphicsContext_t)(void*, void**);
typedef int (__stdcall *GdipDrawImageRectI_t)(void*, void*, INT, INT, INT, INT);
typedef int (__stdcall *GdipDeleteGraphics_t)(void*);
typedef int (__stdcall *GdipSetInterpolationMode_t)(void*, int);
typedef int (__stdcall *GdipGraphicsClear_t)(void*, UINT);

static HICON LoadPngAsIcon(const WCHAR* path, int targetSize, int drawSize, int offsetX, int offsetY) {
    if (!path || !path[0]) return nullptr;

    DWORD attrs = GetFileAttributesW(path);
    if (attrs == INVALID_FILE_ATTRIBUTES) {
        Wh_Log(L"Image file not found: %s", path);
        return nullptr;
    }

    HMODULE hGdiPlus = LoadLibraryW(L"gdiplus.dll");
    if (!hGdiPlus) return nullptr;

    auto pStartup = (GdiplusStartup_t)GetProcAddress(hGdiPlus, "GdiplusStartup");
    auto pShutdown = (GdiplusShutdown_t)GetProcAddress(hGdiPlus, "GdiplusShutdown");
    auto pFromFile = (GdipCreateBitmapFromFile_t)GetProcAddress(hGdiPlus, "GdipCreateBitmapFromFile");
    auto pCreateHICON = (GdipCreateHICONFromBitmap_t)GetProcAddress(hGdiPlus, "GdipCreateHICONFromBitmap");
    auto pDispose = (GdipDisposeImage_t)GetProcAddress(hGdiPlus, "GdipDisposeImage");
    auto pFromScan0 = (GdipCreateBitmapFromScan0_t)GetProcAddress(hGdiPlus, "GdipCreateBitmapFromScan0");
    auto pGetGfx = (GdipGetImageGraphicsContext_t)GetProcAddress(hGdiPlus, "GdipGetImageGraphicsContext");
    auto pDrawRect = (GdipDrawImageRectI_t)GetProcAddress(hGdiPlus, "GdipDrawImageRectI");
    auto pDelGfx = (GdipDeleteGraphics_t)GetProcAddress(hGdiPlus, "GdipDeleteGraphics");
    auto pSetInterp = (GdipSetInterpolationMode_t)GetProcAddress(hGdiPlus, "GdipSetInterpolationMode");
    auto pClear = (GdipGraphicsClear_t)GetProcAddress(hGdiPlus, "GdipGraphicsClear");

    if (!pStartup || !pShutdown || !pFromFile || !pCreateHICON || !pDispose || 
        !pFromScan0 || !pGetGfx || !pDrawRect || !pDelGfx) {
        FreeLibrary(hGdiPlus);
        return nullptr;
    }

    struct { UINT32 ver; void* cb; BOOL noThread; BOOL noCodecs; } input = {1, nullptr, FALSE, FALSE};
    ULONG_PTR token = 0;

    if (pStartup(&token, &input, nullptr) != 0) {
        FreeLibrary(hGdiPlus);
        return nullptr;
    }

    HICON hIcon = nullptr;
    void* pBitmap = nullptr;

    if (pFromFile(path, &pBitmap) == 0 && pBitmap) {
        void* pTarget = nullptr;
        if (pFromScan0(targetSize, targetSize, 0, 0x0026200A, nullptr, &pTarget) == 0 && pTarget) {
            void* pGfx = nullptr;
            if (pGetGfx(pTarget, &pGfx) == 0 && pGfx) {
                if (pClear) pClear(pGfx, 0x00000000);
                if (pSetInterp) pSetInterp(pGfx, 7);

                // Draw the source image scaled to drawSize at offset position
                pDrawRect(pGfx, pBitmap, offsetX, offsetY, drawSize, drawSize);
                pDelGfx(pGfx);
            }
            pCreateHICON(pTarget, &hIcon);
            pDispose(pTarget);
        }
        pDispose(pBitmap);
    }

    pShutdown(token);
    return hIcon;
}

static HICON LoadPngAsIconSimple(const WCHAR* path, int size) {
    return LoadPngAsIcon(path, size, size, 0, 0);
}

// Balloon Constructor

static void DoShowBalloon(const WCHAR* title, const WCHAR* body) {
    if (!g_hBalloonWnd) return;
    EnsureTrayIcon();

    if (g_lastNotifIcon) {
        DestroyIcon(g_lastNotifIcon);
        g_lastNotifIcon = nullptr;
    }

    bool useCustomIcon = false;
    bool useLargeIconFlag = false;
    
    if (g_settings.showProfilePicture && g_lastImagePath[0]) {
        switch (g_settings.iconSize) {
            case ICON_SIZE_SMALL:
                g_lastNotifIcon = LoadPngAsIconSimple(g_lastImagePath, 16);
                useLargeIconFlag = false;
                break;
                
            case ICON_SIZE_MEDIUM:
                g_lastNotifIcon = LoadPngAsIcon(g_lastImagePath, 32, 24, 4, 0);
                useLargeIconFlag = true;
                break;
                
            case ICON_SIZE_LARGE:
            default:
                g_lastNotifIcon = LoadPngAsIconSimple(g_lastImagePath, 32);
                useLargeIconFlag = true;
                break;
        }
        
        if (g_lastNotifIcon) {
            useCustomIcon = true;
            //Wh_Log(L"Loaded PFP icon (mode=%d) from %s", g_settings.iconSize, g_lastImagePath);
        }
    }

    NOTIFYICONDATAW nid = {};
    nid.cbSize = sizeof(nid);
    nid.hWnd = g_hBalloonWnd;
    nid.uID = 1;
    nid.uFlags = NIF_INFO;

    if (useCustomIcon) {
        nid.dwInfoFlags = NIIF_USER;
        if (useLargeIconFlag) {
            nid.dwInfoFlags |= NIIF_LARGE_ICON;
        }
        nid.hBalloonIcon = g_lastNotifIcon;
    } else {
        nid.dwInfoFlags = NIIF_INFO;
    }

    wcsncpy(nid.szInfoTitle,
        (title && title[0]) ? title : L"Discord",
        ARRAYSIZE(nid.szInfoTitle) - 1);
    wcsncpy(nid.szInfo,
        (body && body[0]) ? body : L"New message",
        ARRAYSIZE(nid.szInfo) - 1);

    Shell_NotifyIconW(NIM_MODIFY, &nid);
    Wh_Log(L"Balloon: \"%s\" - \"%s\" (icon=%s, mode=%d, largeFlag=%s)",
        nid.szInfoTitle, nid.szInfo,
        useCustomIcon ? L"PFP" : L"default",
        g_settings.iconSize,
        useLargeIconFlag ? L"yes" : L"no");
}

static void ShowBalloonNotification(const WCHAR* title, const WCHAR* body) {
    EnterCriticalSection(&g_cs);
    wcsncpy(g_pendingTitle, title ? title : L"Discord", ARRAYSIZE(g_pendingTitle) - 1);
    wcsncpy(g_pendingBody, body ? body : L"New message", ARRAYSIZE(g_pendingBody) - 1);
    LeaveCriticalSection(&g_cs);

    if (g_hBalloonWnd)
        PostMessageW(g_hBalloonWnd, WM_BALLOON_SHOW, 0, 0);
}

static void RemoveTrayIcon() {
    if (g_iconAdded && g_hBalloonWnd) {
        NOTIFYICONDATAW nid = {};
        nid.cbSize = sizeof(nid);
        nid.hWnd = g_hBalloonWnd;
        nid.uID = 1;
        Shell_NotifyIconW(NIM_DELETE, &nid);
        g_iconAdded = false;
        Wh_Log(L"Tray icon removed");
    }
}

static LRESULT CALLBACK BalloonWndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    if (msg == WM_BALLOON_SHOW) {
        WCHAR title[256], body[512];
        EnterCriticalSection(&g_cs);
        wcscpy_s(title, g_pendingTitle);
        wcscpy_s(body, g_pendingBody);
        LeaveCriticalSection(&g_cs);
        DoShowBalloon(title, body);
        return 0;
    }

    if (msg == WM_TRAY_CALLBACK) {
        UINT event = LOWORD(lParam);
        if (event == NIN_BALLOONUSERCLICK ||
            event == WM_LBUTTONUP ||
            event == WM_LBUTTONDBLCLK)
        {
            FocusDiscordWindow();
        }
        return 0;
    }
    
    if (msg == WM_CHECK_DISCORD) {
        if (!IsAnyDiscordRunning()) {
            Wh_Log(L"No Discord processes found, removing tray icon");
            RemoveTrayIcon();
        }
        return 0;
    }

    return DefWindowProcW(hWnd, msg, wParam, lParam);
}

static DWORD WINAPI DiscordMonitorThread(LPVOID) {
    Wh_Log(L"Discord monitor thread started");
    
    while (!g_stopMonitor) {
        Sleep(2000); // Check every 2 seconds
        
        if (g_stopMonitor) break;
        
        if (g_iconAdded && g_hBalloonWnd) {
            PostMessageW(g_hBalloonWnd, WM_CHECK_DISCORD, 0, 0);
        }
    }
    
    Wh_Log(L"Discord monitor thread stopped");
    return 0;
}

static DWORD WINAPI BalloonThread(LPVOID) {
    WNDCLASSW wc = {};
    wc.lpfnWndProc = BalloonWndProc;
    wc.hInstance = GetModuleHandleW(nullptr);
    wc.lpszClassName = L"WindhawkDiscordBalloonClass";
    RegisterClassW(&wc);

    g_hBalloonWnd = CreateWindowExW(0, wc.lpszClassName, L"", 0,
        0, 0, 0, 0, HWND_MESSAGE, nullptr, wc.hInstance, nullptr);

    if (!g_hBalloonWnd) return 1;

    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    RemoveTrayIcon();
    return 0;
}

// XML string parsing

static void DecodeXmlEntities(WCHAR* str) {
    if (!str) return;

    WCHAR* read = str;
    WCHAR* write = str;

    while (*read) {
        if (*read == L'&') {
            if (wcsncmp(read, L"&amp;", 5) == 0) {
                *write++ = L'&'; read += 5;
            } else if (wcsncmp(read, L"&lt;", 4) == 0) {
                *write++ = L'<'; read += 4;
            } else if (wcsncmp(read, L"&gt;", 4) == 0) {
                *write++ = L'>'; read += 4;
            } else if (wcsncmp(read, L"&apos;", 6) == 0) {
                *write++ = L'\''; read += 6;
            } else if (wcsncmp(read, L"&quot;", 6) == 0) {
                *write++ = L'"'; read += 6;
            } else if (wcsncmp(read, L"&#x", 3) == 0 || wcsncmp(read, L"&#X", 3) == 0) {
                WCHAR* afterNum = nullptr;
                unsigned long val = wcstoul(read + 3, &afterNum, 16);
                if (afterNum && *afterNum == L';' && val > 0 && val <= 0xFFFF) {
                    *write++ = (WCHAR)val;
                    read = afterNum + 1;
                } else {
                    *write++ = *read++;
                }
            } else if (wcsncmp(read, L"&#", 2) == 0) {
                WCHAR* afterNum = nullptr;
                unsigned long val = wcstoul(read + 2, &afterNum, 10);
                if (afterNum && *afterNum == L';' && val > 0 && val <= 0xFFFF) {
                    *write++ = (WCHAR)val;
                    read = afterNum + 1;
                } else {
                    *write++ = *read++;
                }
            } else {
                *write++ = *read++;
            }
        } else {
            *write++ = *read++;
        }
    }
    *write = L'\0';
}

static void StripInvisibleChars(WCHAR* str) {
    if (!str) return;

    WCHAR* read = str;
    WCHAR* write = str;

    while (*read) {
        WCHAR c = *read;

        if (c >= 0xFE00 && c <= 0xFE0F) { read++; continue; }
        if (c >= 0x200B && c <= 0x200F) { read++; continue; }
        if (c >= 0x2028 && c <= 0x202F) { read++; continue; }
        if (c >= 0x2060 && c <= 0x2069) { read++; continue; }
        if (c == 0xFEFF) { read++; continue; }
        if (c == 0xFFFC || c == 0xFFFD) { read++; continue; }

        if (c >= 0xD800 && c <= 0xDBFF) {
            if (*(read + 1) >= 0xDC00 && *(read + 1) <= 0xDFFF) {
                read += 2;
            } else {
                read++;
            }
            continue;
        }
        if (c >= 0xDC00 && c <= 0xDFFF) { read++; continue; }

        *write++ = *read++;
    }
    *write = L'\0';

    WCHAR* start = str;
    while (*start == L' ') start++;
    if (start != str) {
        WCHAR* dst = str;
        while (*start) *dst++ = *start++;
        *dst = L'\0';
    }

    int len = (int)wcslen(str);
    while (len > 0 && str[len - 1] == L' ') {
        str[--len] = L'\0';
    }

    read = str;
    write = str;
    bool lastWasSpace = false;
    while (*read) {
        if (*read == L' ') {
            if (!lastWasSpace) *write++ = *read;
            lastWasSpace = true;
        } else {
            *write++ = *read;
            lastWasSpace = false;
        }
        read++;
    }
    *write = L'\0';
}

static WCHAR g_lastXmlTitle[256] = {};
static WCHAR g_lastXmlBody[512] = {};
static bool g_haveXmlText = false;

static void ParseTextFromXmlString(const WCHAR* xml) {
    g_lastXmlTitle[0] = L'\0';
    g_lastXmlBody[0] = L'\0';
    g_lastImagePath[0] = L'\0';
    g_haveXmlText = false;
    if (!xml) return;

    // Extract image src
    const WCHAR* imgTag = wcsstr(xml, L"<image");
    if (imgTag) {
        const WCHAR* srcAttr = wcsstr(imgTag, L"src='");
        if (!srcAttr) srcAttr = wcsstr(imgTag, L"src=\"");
        if (srcAttr) {
            srcAttr += 5;
            WCHAR quote = *(srcAttr - 1);
            const WCHAR* srcEnd = wcschr(srcAttr, quote);
            if (srcEnd) {
                int len = (int)(srcEnd - srcAttr);
                int cl = MIN_VAL(len, (int)ARRAYSIZE(g_lastImagePath) - 1);
                wcsncpy(g_lastImagePath, srcAttr, cl);
                g_lastImagePath[cl] = L'\0';
                DecodeXmlEntities(g_lastImagePath);
                for (WCHAR* p = g_lastImagePath; *p; p++) {
                    if (*p == L'/') *p = L'\\';
                }
                Wh_Log(L"Notification image: \"%s\"", g_lastImagePath);
            }
        }
    }

    // Extract <text> elements
    int textIndex = 0;
    const WCHAR* pos = xml;

    while (pos && textIndex < 3) {
        const WCHAR* tagStart = wcsstr(pos, L"<text");
        if (!tagStart) break;

        const WCHAR* contentStart = wcschr(tagStart, L'>');
        if (!contentStart) break;
        contentStart++;

        if (*(contentStart - 2) == L'/') {
            pos = contentStart;
            textIndex++;
            continue;
        }

        const WCHAR* contentEnd = wcsstr(contentStart, L"</text>");
        if (!contentEnd) break;

        int contentLen = (int)(contentEnd - contentStart);
        if (contentLen > 0) {
            if (textIndex == 0) {
                int cl = MIN_VAL(contentLen, (int)ARRAYSIZE(g_lastXmlTitle) - 1);
                wcsncpy(g_lastXmlTitle, contentStart, cl);
                g_lastXmlTitle[cl] = L'\0';
                g_haveXmlText = true;
            } else {
                int cur = (int)wcslen(g_lastXmlBody);
                if (cur > 0 && cur < (int)ARRAYSIZE(g_lastXmlBody) - 2) {
                    g_lastXmlBody[cur] = L'\n';
                    cur++;
                }
                int cl = MIN_VAL(contentLen, (int)ARRAYSIZE(g_lastXmlBody) - cur - 1);
                wcsncpy(g_lastXmlBody + cur, contentStart, cl);
                g_lastXmlBody[cur + cl] = L'\0';
            }
        }

        pos = contentEnd + 7;
        textIndex++;
    }

    DecodeXmlEntities(g_lastXmlTitle);
    DecodeXmlEntities(g_lastXmlBody);
    StripInvisibleChars(g_lastXmlTitle);
    StripInvisibleChars(g_lastXmlBody);

    if (g_haveXmlText) {
        Wh_Log(L"Parsed: title=\"%s\" body=\"%s\"", g_lastXmlTitle, g_lastXmlBody);
    }
}

// Vtable patching

static bool PatchVtableEntry(void** vtbl, int index, void* hookFn, void** origFn) {
    void* current = vtbl[index];
    if (current == hookFn) return true;

    *origFn = current;
    DWORD oldProtect;
    if (VirtualProtect(&vtbl[index], sizeof(void*), PAGE_READWRITE, &oldProtect)) {
        vtbl[index] = hookFn;
        VirtualProtect(&vtbl[index], sizeof(void*), oldProtect, &oldProtect);
        return true;
    }
    if (VirtualProtect(&vtbl[index], sizeof(void*), PAGE_EXECUTE_READWRITE, &oldProtect)) {
        vtbl[index] = hookFn;
        VirtualProtect(&vtbl[index], sizeof(void*), oldProtect, &oldProtect);
        return true;
    }
    return false;
}

// LoadXml hooking

typedef HRESULT (STDMETHODCALLTYPE *LoadXml_t)(void* pThis, HSTRING xml);
static LoadXml_t LoadXml_Orig = nullptr;
static bool g_loadXmlHooked = false;

static const GUID IID_IXmlDocumentIO =
    {0x6CD0E74E, 0xEE65, 0x4489, {0x9E,0xBF,0xCA,0x43,0xE8,0x7B,0xA6,0x37}};

static HRESULT STDMETHODCALLTYPE LoadXml_Hook(void* pThis, HSTRING xml) {
    UINT32 len = 0;
    const WCHAR* xmlStr = WindowsGetStringRawBuffer(xml, &len);

    if (xmlStr && len > 0) {
        Wh_Log(L"=== RAW XML (len=%u) ===", len);
        const int CHUNK = 500;
        for (UINT32 i = 0; i < len; i += CHUNK) {
            WCHAR chunk[CHUNK + 1] = {};
            int cl = MIN_VAL((int)(len - i), CHUNK);
            wcsncpy(chunk, xmlStr + i, cl);
            chunk[cl] = L'\0';
            Wh_Log(L"%s", chunk);
        }

        if (wcsstr(xmlStr, L"<toast")) {
            ParseTextFromXmlString(xmlStr);
        }
    }

    return LoadXml_Orig(pThis, xml);
}

static void TryHookLoadXml(void* pXmlDocInstance) {
    if (g_loadXmlHooked || !pXmlDocInstance) return;

    void* pDocIO = nullptr;
    HRESULT hr = ((IUnknown*)pXmlDocInstance)->QueryInterface(IID_IXmlDocumentIO, &pDocIO);
    if (FAILED(hr) || !pDocIO) return;

    void** ioVtbl = *(void***)pDocIO;
    if (PatchVtableEntry(ioVtbl, 6, (void*)LoadXml_Hook, (void**)&LoadXml_Orig)) {
        g_loadXmlHooked = true;
        Wh_Log(L"Hooked LoadXml via vtable patch");
    }

    ((IUnknown*)pDocIO)->Release();
}

// ToastNotifier_Show hooking

typedef HRESULT (STDMETHODCALLTYPE *ToastNotifier_Show_t)(void* pThis, void* pNotification);
static ToastNotifier_Show_t ToastNotifier_Show_Orig = nullptr;
static bool g_showHooked = false;

static HRESULT STDMETHODCALLTYPE ToastNotifier_Show_Hook(void* pThis, void* pNotification) {
    Wh_Log(L"Show: toast suppressed");
    return S_OK;
}

// CreateToastNotification hooking

typedef HRESULT (STDMETHODCALLTYPE *CreateToastNotification_t)(void* pThis, void* pXmlDoc, void** ppNotification);
static CreateToastNotification_t CreateToastNotification_Orig = nullptr;
static bool g_notifFactoryHooked = false;

static HRESULT STDMETHODCALLTYPE CreateToastNotification_Hook(void* pThis, void* pXmlDoc, void** ppNotification) {
    if (g_haveXmlText) {
        ShowBalloonNotification(
            g_lastXmlTitle[0] ? g_lastXmlTitle : L"Discord",
            g_lastXmlBody[0] ? g_lastXmlBody : L"New message"
        );
        g_haveXmlText = false;
    } else {
        ShowBalloonNotification(L"Discord", L"New message");
    }

    return CreateToastNotification_Orig(pThis, pXmlDoc, ppNotification);
}

static void TryHookNotificationFactory(void* pFactory) {
    if (g_notifFactoryHooked || !pFactory) return;

    void** vtbl = *(void***)pFactory;
    if (!vtbl) return;

    if (PatchVtableEntry(vtbl, 6, (void*)CreateToastNotification_Hook, (void**)&CreateToastNotification_Orig)) {
        g_notifFactoryHooked = true;
        Wh_Log(L"Hooked CreateToastNotification via vtable patch");
    }
}

// Proactive Show hooking

static bool g_inProactiveHook = false;

static void ProactivelyHookShow() {
    if (g_showHooked || g_inProactiveHook) return;
    g_inProactiveHook = true;

    static const GUID IID_IToastNotificationManagerStatics =
        {0x50AC103F, 0xD235, 0x4598, {0xBB,0xEF,0x98,0xFE,0x4D,0x1A,0x3A,0xD4}};

    HSTRING hClassName = nullptr;
    WindowsCreateString(
        L"Windows.UI.Notifications.ToastNotificationManager",
        (UINT32)wcslen(L"Windows.UI.Notifications.ToastNotificationManager"),
        &hClassName);

    void* pManagerStatics = nullptr;
    HRESULT hr = RoGetActivationFactory(hClassName, IID_IToastNotificationManagerStatics, &pManagerStatics);
    WindowsDeleteString(hClassName);

    if (FAILED(hr) || !pManagerStatics) {
        g_inProactiveHook = false;
        return;
    }

    void** managerVtbl = *(void***)pManagerStatics;
    typedef HRESULT (STDMETHODCALLTYPE *CreateToastNotifierWithId_t)(void* pThis, HSTRING appId, void** ppNotifier);
    CreateToastNotifierWithId_t createWithId = (CreateToastNotifierWithId_t)managerVtbl[7];

    void* pNotifier = nullptr;
    HSTRING hAppId = nullptr;
    WindowsCreateString(L"Discord", 7, &hAppId);
    hr = createWithId(pManagerStatics, hAppId, &pNotifier);
    WindowsDeleteString(hAppId);

    if (SUCCEEDED(hr) && pNotifier) {
        void** notifierVtbl = *(void***)pNotifier;
        void* showAddr = notifierVtbl[6];
        Wh_SetFunctionHook(showAddr, (void*)ToastNotifier_Show_Hook, (void**)&ToastNotifier_Show_Orig);
        Wh_ApplyHookOperations();
        g_showHooked = true;
        Wh_Log(L"Hooked IToastNotifier::Show at %p", showAddr);
        ((IUnknown*)pNotifier)->Release();
    }

    ((IUnknown*)pManagerStatics)->Release();
    g_inProactiveHook = false;
}

// RoGetActivationFactory hooking

typedef HRESULT (WINAPI *RoGetActivationFactory_t)(HSTRING classId, REFIID iid, void** factory);
static RoGetActivationFactory_t RoGetActivationFactory_Orig = nullptr;

HRESULT WINAPI RoGetActivationFactory_Hook(HSTRING classId, REFIID iid, void** factory) {
    HRESULT hr = RoGetActivationFactory_Orig(classId, iid, factory);

    if (SUCCEEDED(hr) && factory && *factory) {
        UINT32 len = 0;
        const WCHAR* name = WindowsGetStringRawBuffer(classId, &len);
        if (!name) return hr;

        if (wcscmp(name, L"Windows.UI.Notifications.ToastNotification") == 0) {
            TryHookNotificationFactory(*factory);
            if (!g_showHooked) ProactivelyHookShow();
        }
        else if (!g_showHooked && !g_inProactiveHook && wcsstr(name, L"ToastNotificationManager")) {
            ProactivelyHookShow();
        }
    }

    return hr;
}

// RoActivateInstance hooking

typedef HRESULT (WINAPI *RoActivateInstance_t)(HSTRING classId, void** instance);
static RoActivateInstance_t RoActivateInstance_Orig = nullptr;

HRESULT WINAPI RoActivateInstance_Hook(HSTRING classId, void** instance) {
    HRESULT hr = RoActivateInstance_Orig(classId, instance);

    if (SUCCEEDED(hr) && instance && *instance && !g_loadXmlHooked) {
        UINT32 len = 0;
        const WCHAR* name = WindowsGetStringRawBuffer(classId, &len);
        if (name && wcscmp(name, L"Windows.Data.Xml.Dom.XmlDocument") == 0) {
            TryHookLoadXml(*instance);
        }
    }

    return hr;
}

// Windhawk functions

BOOL Wh_ModInit(void) {
    Wh_Log(L"Discord Balloon Notifications mod started");

    LoadSettings();
    
    g_currentPid = GetCurrentProcessId();

    WCHAR exePath[MAX_PATH];
    if (GetModuleFileNameW(nullptr, exePath, MAX_PATH)) {
        g_hAppIcon = ExtractIconW(GetModuleHandleW(nullptr), exePath, 0);
        if (g_hAppIcon == (HICON)1) g_hAppIcon = nullptr;
    }

    InitializeCriticalSection(&g_cs);

    g_hThread = CreateThread(nullptr, 0, BalloonThread, nullptr, 0, nullptr);
    Sleep(200);
    
    // Start Discord process monitor thread
    g_stopMonitor = false;
    g_hMonitorThread = CreateThread(nullptr, 0, DiscordMonitorThread, nullptr, 0, nullptr);

    HMODULE combase = GetModuleHandleW(L"combase.dll");
    if (!combase) combase = LoadLibraryW(L"combase.dll");

    if (combase) {
        void* p1 = (void*)GetProcAddress(combase, "RoGetActivationFactory");
        void* p2 = (void*)GetProcAddress(combase, "RoActivateInstance");

        if (p1) Wh_SetFunctionHook(p1, (void*)RoGetActivationFactory_Hook, (void**)&RoGetActivationFactory_Orig);
        if (p2) Wh_SetFunctionHook(p2, (void*)RoActivateInstance_Hook, (void**)&RoActivateInstance_Orig);
    }

    Wh_Log(L"Mod init complete");
    return TRUE;
}

void Wh_ModSettingsChanged(void) {
    LoadSettings();
}

void Wh_ModUninit(void) {
    // Stop monitor thread
    g_stopMonitor = true;
    if (g_hMonitorThread) {
        WaitForSingleObject(g_hMonitorThread, 3000);
        CloseHandle(g_hMonitorThread);
        g_hMonitorThread = nullptr;
    }

    RemoveTrayIcon();

    if (g_lastNotifIcon) {
        DestroyIcon(g_lastNotifIcon);
        g_lastNotifIcon = nullptr;
    }

    if (g_hBalloonWnd)
        PostMessageW(g_hBalloonWnd, WM_QUIT, 0, 0);

    if (g_hThread) {
        WaitForSingleObject(g_hThread, 3000);
        CloseHandle(g_hThread);
    }

    if (g_hAppIcon)
        DestroyIcon(g_hAppIcon);

    DeleteCriticalSection(&g_cs);
    UnregisterClassW(L"WindhawkDiscordBalloonClass", GetModuleHandleW(nullptr));
}
