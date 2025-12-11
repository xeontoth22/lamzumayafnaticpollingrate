#define UNICODE
#define _UNICODE
#include <windows.h>
#include <shellapi.h>
#include <hidsdi.h>
#include <setupapi.h>
#include <tlhelp32.h>
#include <commctrl.h>
#include <string>
#include <vector>
#include <fstream>
#include <thread>
#include <atomic>
#include <algorithm>

#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "hid.lib")
#pragma comment(lib, "setupapi.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#define WM_TRAYICON (WM_USER + 1)
#define ID_TRAY_EXIT 1001
#define ID_TRAY_1000 1002
#define ID_TRAY_4000 1003
#define ID_TRAY_8000 1004
#define ID_TRAY_AUTO 1005
#define ID_TRAY_NOTIFY 1006
#define ID_TRAY_ADDRULE 1007
#define ID_TRAY_OPEN 1008
#define ID_TRAY_HIGHSPEED 1009
#define ID_TRAY_RULE_BASE 2000

#define ID_BTN_1000 3001
#define ID_BTN_4000 3002
#define ID_BTN_8000 3003
#define ID_CHK_AUTO 3004
#define ID_CHK_NOTIFY 3005
#define ID_CHK_HIGHSPEED 3006
#define ID_LIST_RULES 3007
#define ID_BTN_ADDRULE 3008
#define ID_BTN_EDITRULE 3009
#define ID_BTN_DELRULE 3010

const USHORT VENDOR_ID = 0x3554;
const USHORT PRODUCT_ID = 0xF510;

// Asiimov colors
const COLORREF CLR_BG = RGB(240, 240, 240);
const COLORREF CLR_DARK = RGB(30, 30, 30);
const COLORREF CLR_ORANGE = RGB(255, 140, 26);
const COLORREF CLR_WHITE = RGB(255, 255, 255);
const COLORREF CLR_GRAY = RGB(180, 180, 180);

struct Rule {
    std::wstring process;
    int rate;
};

struct AppState {
    HWND hwnd;
    HWND mainWnd;
    NOTIFYICONDATA nid;
    HICON icon1k, icon4k, icon8k;
    HBRUSH bgBrush, darkBrush, orangeBrush;
    HFONT titleFont, normalFont, boldFont;
    int currentRate = 1000;
    bool autoEnabled = false;
    bool notificationsEnabled = true;
    bool highSpeedEnabled = false;
    std::vector<Rule> rules;
    std::wstring devicePath;
    std::atomic<bool> running{true};
    std::wstring lastMatchedProcess;
    std::wstring configPath;
    bool mainWndVisible = false;
};

AppState g_app;

BYTE POLLING_1000[] = {0x08,0x07,0x00,0x00,0x00,0x02,0x01,0x54,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0xef};
BYTE POLLING_4000[] = {0x08,0x07,0x00,0x00,0x00,0x02,0x20,0x35,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0xef};
BYTE POLLING_8000[] = {0x08,0x07,0x00,0x00,0x00,0x02,0x40,0x15,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0xef};
BYTE HIGHSPEED_ON[] = {0x08,0x07,0x00,0x00,0xb5,0x02,0x01,0x54,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x3a};
BYTE HIGHSPEED_OFF[] = {0x08,0x07,0x00,0x00,0xb5,0x02,0x00,0x55,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x3a};

std::wstring GetConfigPath() {
    wchar_t path[MAX_PATH];
    GetModuleFileNameW(NULL, path, MAX_PATH);
    std::wstring p(path);
    size_t pos = p.find_last_of(L"\\");
    return p.substr(0, pos + 1) + L"polling_rate.json";
}

void SaveConfig() {
    std::wofstream f(g_app.configPath);
    if (!f) return;
    f << L"{\n";
    f << L"  \"rate\": " << g_app.currentRate << L",\n";
    f << L"  \"auto_enabled\": " << (g_app.autoEnabled ? L"true" : L"false") << L",\n";
    f << L"  \"notifications\": " << (g_app.notificationsEnabled ? L"true" : L"false") << L",\n";
    f << L"  \"highspeed\": " << (g_app.highSpeedEnabled ? L"true" : L"false") << L",\n";
    f << L"  \"rules\": [\n";
    for (size_t i = 0; i < g_app.rules.size(); i++) {
        f << L"    {\"process\": \"" << g_app.rules[i].process << L"\", \"rate\": " << g_app.rules[i].rate << L"}";
        if (i < g_app.rules.size() - 1) f << L",";
        f << L"\n";
    }
    f << L"  ]\n}\n";
}

void LoadConfig() {
    std::wifstream f(g_app.configPath);
    if (!f) return;
    std::wstring content((std::istreambuf_iterator<wchar_t>(f)), std::istreambuf_iterator<wchar_t>());
    
    size_t pos;
    if ((pos = content.find(L"\"rate\":")) != std::wstring::npos) {
        g_app.currentRate = _wtoi(content.c_str() + pos + 7);
    }
    if (content.find(L"\"auto_enabled\": true") != std::wstring::npos) {
        g_app.autoEnabled = true;
    }
    if (content.find(L"\"notifications\": false") != std::wstring::npos) {
        g_app.notificationsEnabled = false;
    }
    if (content.find(L"\"highspeed\": true") != std::wstring::npos) {
        g_app.highSpeedEnabled = true;
    }
    
    size_t rulesStart = content.find(L"\"rules\":");
    if (rulesStart != std::wstring::npos) {
        size_t searchPos = rulesStart;
        while ((pos = content.find(L"\"process\":", searchPos)) != std::wstring::npos) {
            size_t nameStart = content.find(L"\"", pos + 10) + 1;
            size_t nameEnd = content.find(L"\"", nameStart);
            size_t ratePos = content.find(L"\"rate\":", pos);
            if (nameEnd != std::wstring::npos && ratePos != std::wstring::npos) {
                Rule r;
                r.process = content.substr(nameStart, nameEnd - nameStart);
                r.rate = _wtoi(content.c_str() + ratePos + 7);
                g_app.rules.push_back(r);
            }
            searchPos = pos + 1;
        }
    }
}

std::wstring FindDevice() {
    GUID hidGuid;
    HidD_GetHidGuid(&hidGuid);
    
    HDEVINFO devInfo = SetupDiGetClassDevsW(&hidGuid, NULL, NULL, DIGCF_PRESENT | DIGCF_DEVICEINTERFACE);
    if (devInfo == INVALID_HANDLE_VALUE) return L"";
    
    SP_DEVICE_INTERFACE_DATA ifData;
    ifData.cbSize = sizeof(SP_DEVICE_INTERFACE_DATA);
    
    for (DWORD i = 0; SetupDiEnumDeviceInterfaces(devInfo, NULL, &hidGuid, i, &ifData); i++) {
        DWORD reqSize;
        SetupDiGetDeviceInterfaceDetailW(devInfo, &ifData, NULL, 0, &reqSize, NULL);
        
        auto detailData = (PSP_DEVICE_INTERFACE_DETAIL_DATA_W)malloc(reqSize);
        detailData->cbSize = sizeof(SP_DEVICE_INTERFACE_DETAIL_DATA_W);
        
        if (SetupDiGetDeviceInterfaceDetailW(devInfo, &ifData, detailData, reqSize, NULL, NULL)) {
            HANDLE h = CreateFileW(detailData->DevicePath, GENERIC_READ | GENERIC_WRITE,
                FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, NULL);
            
            if (h != INVALID_HANDLE_VALUE) {
                HIDD_ATTRIBUTES attr;
                attr.Size = sizeof(HIDD_ATTRIBUTES);
                if (HidD_GetAttributes(h, &attr)) {
                    if (attr.VendorID == VENDOR_ID && attr.ProductID == PRODUCT_ID) {
                        PHIDP_PREPARSED_DATA preparsed;
                        if (HidD_GetPreparsedData(h, &preparsed)) {
                            HIDP_CAPS caps;
                            if (HidP_GetCaps(preparsed, &caps) == HIDP_STATUS_SUCCESS) {
                                if (caps.OutputReportByteLength == 17 && caps.UsagePage == 0xFF02) {
                                    std::wstring path = detailData->DevicePath;
                                    HidD_FreePreparsedData(preparsed);
                                    CloseHandle(h);
                                    free(detailData);
                                    SetupDiDestroyDeviceInfoList(devInfo);
                                    return path;
                                }
                            }
                            HidD_FreePreparsedData(preparsed);
                        }
                    }
                }
                CloseHandle(h);
            }
        }
        free(detailData);
    }
    
    SetupDiDestroyDeviceInfoList(devInfo);
    return L"";
}

bool SendCommand(BYTE* data, size_t len) {
    if (g_app.devicePath.empty()) {
        g_app.devicePath = FindDevice();
    }
    if (g_app.devicePath.empty()) return false;
    
    HANDLE h = CreateFileW(g_app.devicePath.c_str(), GENERIC_READ | GENERIC_WRITE,
        FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, NULL);
    
    if (h == INVALID_HANDLE_VALUE) return false;
    
    BOOL result = HidD_SetOutputReport(h, data, (ULONG)len);
    CloseHandle(h);
    return result != 0;
}

HICON CreateRateIcon(int rate) {
    HDC screenDC = GetDC(NULL);
    HDC memDC = CreateCompatibleDC(screenDC);
    
    BITMAPINFO bmi = {};
    bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth = 32;
    bmi.bmiHeader.biHeight = -32;
    bmi.bmiHeader.biPlanes = 1;
    bmi.bmiHeader.biBitCount = 32;
    
    void* bits;
    HBITMAP hBmp = CreateDIBSection(memDC, &bmi, DIB_RGB_COLORS, &bits, NULL, 0);
    HBITMAP oldBmp = (HBITMAP)SelectObject(memDC, hBmp);
    
    COLORREF bgColor;
    const wchar_t* text;
    if (rate == 8000) { bgColor = RGB(220, 40, 20); text = L"8K"; }
    else if (rate == 4000) { bgColor = RGB(255, 140, 30); text = L"4K"; }
    else { bgColor = RGB(40, 180, 70); text = L"1K"; }
    
    HBRUSH brush = CreateSolidBrush(bgColor);
    HPEN pen = CreatePen(PS_SOLID, 1, bgColor);
    SelectObject(memDC, brush);
    SelectObject(memDC, pen);
    Ellipse(memDC, 0, 0, 32, 32);
    DeleteObject(brush);
    DeleteObject(pen);
    
    SetBkMode(memDC, TRANSPARENT);
    SetTextColor(memDC, RGB(255, 255, 255));
    HFONT font = CreateFontW(14, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
        OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Arial");
    HFONT oldFont = (HFONT)SelectObject(memDC, font);
    
    RECT r = {0, 0, 32, 32};
    DrawTextW(memDC, text, -1, &r, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    
    SelectObject(memDC, oldFont);
    DeleteObject(font);
    
    HBITMAP hMask = CreateCompatibleBitmap(screenDC, 32, 32);
    HDC maskDC = CreateCompatibleDC(screenDC);
    HBITMAP oldMask = (HBITMAP)SelectObject(maskDC, hMask);
    
    HBRUSH whiteBrush = CreateSolidBrush(RGB(255, 255, 255));
    HBRUSH blackBrush = CreateSolidBrush(RGB(0, 0, 0));
    RECT full = {0, 0, 32, 32};
    FillRect(maskDC, &full, whiteBrush);
    SelectObject(maskDC, blackBrush);
    Ellipse(maskDC, 0, 0, 32, 32);
    DeleteObject(whiteBrush);
    DeleteObject(blackBrush);
    
    SelectObject(maskDC, oldMask);
    DeleteDC(maskDC);
    SelectObject(memDC, oldBmp);
    
    ICONINFO ii = {};
    ii.fIcon = TRUE;
    ii.hbmMask = hMask;
    ii.hbmColor = hBmp;
    HICON icon = CreateIconIndirect(&ii);
    
    DeleteObject(hBmp);
    DeleteObject(hMask);
    DeleteDC(memDC);
    ReleaseDC(NULL, screenDC);
    
    return icon;
}

void ShowPopup(const wchar_t* message);
void UpdateMainWindow();

void UpdateTrayIcon() {
    HICON newIcon;
    if (g_app.currentRate == 8000) newIcon = g_app.icon8k;
    else if (g_app.currentRate == 4000) newIcon = g_app.icon4k;
    else newIcon = g_app.icon1k;
    
    g_app.nid.hIcon = newIcon;
    
    std::wstring tip = L"LAMZU: " + std::to_wstring(g_app.currentRate) + L" Hz";
    if (g_app.autoEnabled) tip += L" [AUTO]";
    wcscpy_s(g_app.nid.szTip, tip.c_str());
    
    Shell_NotifyIconW(NIM_MODIFY, &g_app.nid);
    
    if (g_app.mainWnd && g_app.mainWndVisible) {
        UpdateMainWindow();
    }
}

bool SetHighSpeed(bool enabled, bool save = true, bool notify = true) {
    if (enabled == g_app.highSpeedEnabled) return true;
    
    BYTE* data;
    size_t len = 17;
    if (enabled) data = HIGHSPEED_ON;
    else data = HIGHSPEED_OFF;
    
    if (SendCommand(data, len)) {
        g_app.highSpeedEnabled = enabled;
        if (save) SaveConfig();
        UpdateTrayIcon();
        if (notify && g_app.notificationsEnabled) {
            std::wstring msg = L"High Speed: " + std::wstring(enabled ? L"ON" : L"OFF");
            ShowPopup(msg.c_str());
        }
        return true;
    }
    return false;
}

bool SetRate(int rate, bool save = true, bool notify = true) {
    if (rate == g_app.currentRate) return true;
    
    BYTE* data;
    size_t len = 17;
    if (rate == 1000) data = POLLING_1000;
    else if (rate == 4000) data = POLLING_4000;
    else if (rate == 8000) data = POLLING_8000;
    else return false;
    
    if (SendCommand(data, len)) {
        g_app.currentRate = rate;
        if (save) SaveConfig();
        UpdateTrayIcon();
        if (notify && g_app.notificationsEnabled) {
            std::wstring msg = L"Polling rate: " + std::to_wstring(rate) + L" Hz";
            ShowPopup(msg.c_str());
        }
        return true;
    }
    return false;
}

std::vector<std::wstring> GetRunningProcesses() {
    std::vector<std::wstring> procs;
    HANDLE snap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (snap == INVALID_HANDLE_VALUE) return procs;
    
    PROCESSENTRY32W pe;
    pe.dwSize = sizeof(PROCESSENTRY32W);
    
    if (Process32FirstW(snap, &pe)) {
        do {
            std::wstring name = pe.szExeFile;
            std::transform(name.begin(), name.end(), name.begin(), ::towlower);
            procs.push_back(name);
        } while (Process32NextW(snap, &pe));
    }
    
    CloseHandle(snap);
    return procs;
}

void CheckProcesses() {
    if (g_app.rules.empty()) return;
    
    auto running = GetRunningProcesses();
    std::wstring matchedProcess;
    int targetRate = 1000;
    
    for (const auto& rule : g_app.rules) {
        std::wstring procLower = rule.process;
        std::transform(procLower.begin(), procLower.end(), procLower.begin(), ::towlower);
        
        for (const auto& p : running) {
            if (p == procLower) {
                matchedProcess = rule.process;
                targetRate = rule.rate;
                break;
            }
        }
        if (!matchedProcess.empty()) break;
    }
    
    if (matchedProcess != g_app.lastMatchedProcess || 
        (!matchedProcess.empty() && g_app.currentRate != targetRate)) {
        g_app.lastMatchedProcess = matchedProcess;
        SetRate(targetRate, false, true);
    }
}

void MonitorThread() {
    while (g_app.running) {
        if (g_app.autoEnabled) {
            CheckProcesses();
        }
        Sleep(10000);
    }
}

HWND g_popupHwnd = NULL;

void ClosePopupWindow() {
    if (g_popupHwnd && IsWindow(g_popupHwnd)) {
        KillTimer(g_popupHwnd, 1);
        PostMessageW(g_popupHwnd, WM_CLOSE, 0, 0);
    }
    g_popupHwnd = NULL;
}

LRESULT CALLBACK PopupWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    static std::wstring popupText;
    
    switch (msg) {
    case WM_CREATE: {
        CREATESTRUCT* cs = (CREATESTRUCT*)lParam;
        popupText = (const wchar_t*)cs->lpCreateParams;
        SetTimer(hwnd, 1, 2500, NULL);
        return 0;
    }
    case WM_TIMER:
        if (wParam == 1) {
            KillTimer(hwnd, 1);
            PostMessageW(hwnd, WM_CLOSE, 0, 0);
        }
        return 0;
    case WM_LBUTTONDOWN:
    case WM_RBUTTONDOWN:
    case WM_MBUTTONDOWN:
        KillTimer(hwnd, 1);
        PostMessageW(hwnd, WM_CLOSE, 0, 0);
        return 0;
    case WM_NCHITTEST:
        return HTCLIENT;
    case WM_CLOSE:
        KillTimer(hwnd, 1);
        if (g_popupHwnd == hwnd) {
            g_popupHwnd = NULL;
        }
        DestroyWindow(hwnd);
        return 0;
    case WM_DESTROY:
        KillTimer(hwnd, 1);
        if (g_popupHwnd == hwnd) {
            g_popupHwnd = NULL;
        }
        return 0;
    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        
        RECT rc;
        GetClientRect(hwnd, &rc);
        
        HBRUSH borderBrush = CreateSolidBrush(CLR_DARK);
        FillRect(hdc, &rc, borderBrush);
        DeleteObject(borderBrush);
        
        RECT inner = {2, 2, rc.right - 2, rc.bottom - 2};
        HBRUSH bgBrush = CreateSolidBrush(CLR_BG);
        FillRect(hdc, &inner, bgBrush);
        DeleteObject(bgBrush);
        
        RECT accent = {2, 2, rc.right - 2, 6};
        HBRUSH accentBrush = CreateSolidBrush(CLR_ORANGE);
        FillRect(hdc, &accent, accentBrush);
        DeleteObject(accentBrush);
        
        SetBkMode(hdc, TRANSPARENT);
        
        HFONT titleFont = CreateFontW(12, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, 
            CLEARTYPE_QUALITY, FIXED_PITCH, L"Consolas");
        HFONT oldFont = (HFONT)SelectObject(hdc, titleFont);
        SetTextColor(hdc, CLR_DARK);
        RECT titleRect = {15, 12, rc.right - 10, 28};
        DrawTextW(hdc, L"LAMZU", -1, &titleRect, DT_LEFT);
        DeleteObject(titleFont);
        
        HFONT msgFont = CreateFontW(16, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
            CLEARTYPE_QUALITY, FIXED_PITCH, L"Consolas");
        SelectObject(hdc, msgFont);
        SetTextColor(hdc, CLR_ORANGE);
        RECT msgRect = {15, 30, rc.right - 10, rc.bottom - 5};
        DrawTextW(hdc, popupText.c_str(), -1, &msgRect, DT_LEFT);
        
        SelectObject(hdc, oldFont);
        DeleteObject(msgFont);
        
        EndPaint(hwnd, &ps);
        return 0;
    }
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

void ShowPopup(const wchar_t* message) {
    if (g_popupHwnd && IsWindow(g_popupHwnd)) {
        KillTimer(g_popupHwnd, 1);
        DestroyWindow(g_popupHwnd);
        g_popupHwnd = NULL;
        MSG msg;
        while (PeekMessageW(&msg, NULL, 0, 0, PM_REMOVE)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
    
    static bool classRegistered = false;
    if (!classRegistered) {
        WNDCLASSW wc = {};
        wc.lpfnWndProc = PopupWndProc;
        wc.hInstance = GetModuleHandle(NULL);
        wc.lpszClassName = L"LAMZUPopup";
        RegisterClassW(&wc);
        classRegistered = true;
    }
    
    int screenW = GetSystemMetrics(SM_CXSCREEN);
    int screenH = GetSystemMetrics(SM_CYSCREEN);
    int w = 220, h = 60;
    
    g_popupHwnd = CreateWindowExW(
        WS_EX_TOPMOST | WS_EX_TOOLWINDOW,
        L"LAMZUPopup", L"",
        WS_POPUP,
        screenW - w - 20, screenH - h - 50, w, h,
        NULL, NULL, GetModuleHandle(NULL), (LPVOID)message
    );
    
    ShowWindow(g_popupHwnd, SW_SHOWNOACTIVATE);
    UpdateWindow(g_popupHwnd);
}

void ShowContextMenu(HWND hwnd) {
    HMENU menu = CreatePopupMenu();
    HMENU rulesMenu = CreatePopupMenu();
    
    AppendMenuW(menu, 0, ID_TRAY_OPEN, L"Open");
    AppendMenuW(menu, MF_SEPARATOR, 0, NULL);
    AppendMenuW(menu, g_app.currentRate == 1000 ? MF_CHECKED : 0, ID_TRAY_1000, L"1000 Hz");
    AppendMenuW(menu, g_app.currentRate == 4000 ? MF_CHECKED : 0, ID_TRAY_4000, L"4000 Hz");
    AppendMenuW(menu, g_app.currentRate == 8000 ? MF_CHECKED : 0, ID_TRAY_8000, L"8000 Hz");
    AppendMenuW(menu, MF_SEPARATOR, 0, NULL);
    AppendMenuW(menu, g_app.highSpeedEnabled ? MF_CHECKED : 0, ID_TRAY_HIGHSPEED, L"High Speed");
    AppendMenuW(menu, MF_SEPARATOR, 0, NULL);
    AppendMenuW(menu, g_app.autoEnabled ? MF_CHECKED : 0, ID_TRAY_AUTO, L"Auto Mode");
    AppendMenuW(menu, g_app.notificationsEnabled ? MF_CHECKED : 0, ID_TRAY_NOTIFY, L"Notifications");
    
    for (size_t i = 0; i < g_app.rules.size(); i++) {
        std::wstring label = g_app.rules[i].process + L" -> " + std::to_wstring(g_app.rules[i].rate) + L" Hz";
        AppendMenuW(rulesMenu, 0, ID_TRAY_RULE_BASE + i, label.c_str());
    }
    if (!g_app.rules.empty()) {
        AppendMenuW(rulesMenu, MF_SEPARATOR, 0, NULL);
    }
    AppendMenuW(rulesMenu, 0, ID_TRAY_ADDRULE, L"Add Rule...");
    
    AppendMenuW(menu, MF_POPUP, (UINT_PTR)rulesMenu, L"Rules");
    AppendMenuW(menu, MF_SEPARATOR, 0, NULL);
    AppendMenuW(menu, 0, ID_TRAY_EXIT, L"Exit");
    
    POINT pt;
    GetCursorPos(&pt);
    SetForegroundWindow(hwnd);
    TrackPopupMenu(menu, TPM_RIGHTALIGN | TPM_BOTTOMALIGN, pt.x, pt.y, 0, hwnd, NULL);
    
    DestroyMenu(rulesMenu);
    DestroyMenu(menu);
}

void ShowRuleDialog(int editIndex = -1);

void UpdateMainWindow() {
    if (!g_app.mainWnd) return;
    
    // Update rate buttons
    for (int id = ID_BTN_1000; id <= ID_BTN_8000; id++) {
        HWND btn = GetDlgItem(g_app.mainWnd, id);
        int rate = (id == ID_BTN_1000) ? 1000 : (id == ID_BTN_4000 ? 4000 : 8000);
        bool selected = (g_app.currentRate == rate);
        // Will be handled in paint
    }
    
    // Update checkboxes
    CheckDlgButton(g_app.mainWnd, ID_CHK_AUTO, g_app.autoEnabled ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(g_app.mainWnd, ID_CHK_NOTIFY, g_app.notificationsEnabled ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(g_app.mainWnd, ID_CHK_HIGHSPEED, g_app.highSpeedEnabled ? BST_CHECKED : BST_UNCHECKED);
    
    // Update rules list
    HWND list = GetDlgItem(g_app.mainWnd, ID_LIST_RULES);
    SendMessageW(list, LB_RESETCONTENT, 0, 0);
    for (const auto& rule : g_app.rules) {
        std::wstring item = rule.process + L"  ->  " + std::to_wstring(rule.rate) + L" Hz";
        SendMessageW(list, LB_ADDSTRING, 0, (LPARAM)item.c_str());
    }
    
    InvalidateRect(g_app.mainWnd, NULL, TRUE);
}

void DrawRateButton(HDC hdc, RECT rc, const wchar_t* text, bool selected, COLORREF color) {
    HBRUSH brush = CreateSolidBrush(selected ? color : CLR_GRAY);
    HPEN pen = CreatePen(PS_SOLID, 2, CLR_DARK);
    SelectObject(hdc, brush);
    SelectObject(hdc, pen);
    RoundRect(hdc, rc.left, rc.top, rc.right, rc.bottom, 10, 10);
    DeleteObject(brush);
    DeleteObject(pen);
    
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, selected ? CLR_WHITE : CLR_DARK);
    DrawTextW(hdc, text, -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
}

LRESULT CALLBACK MainWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE: {
        // Title
        CreateWindowExW(0, L"STATIC", L"LAMZU POLLING RATE",
            WS_CHILD | WS_VISIBLE | SS_CENTER,
            0, 15, 380, 30, hwnd, NULL, GetModuleHandle(NULL), NULL);
        
        // Rate buttons (custom drawn)
        CreateWindowExW(0, L"BUTTON", L"1000 Hz",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            30, 60, 100, 50, hwnd, (HMENU)ID_BTN_1000, GetModuleHandle(NULL), NULL);
        CreateWindowExW(0, L"BUTTON", L"4000 Hz",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            140, 60, 100, 50, hwnd, (HMENU)ID_BTN_4000, GetModuleHandle(NULL), NULL);
        CreateWindowExW(0, L"BUTTON", L"8000 Hz",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            250, 60, 100, 50, hwnd, (HMENU)ID_BTN_8000, GetModuleHandle(NULL), NULL);
        
        // Separator
        CreateWindowExW(0, L"STATIC", L"",
            WS_CHILD | WS_VISIBLE | SS_ETCHEDHORZ,
            20, 125, 340, 2, hwnd, NULL, GetModuleHandle(NULL), NULL);
        
        // Checkboxes
        CreateWindowExW(0, L"BUTTON", L"Auto Mode",
            WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
            30, 140, 150, 25, hwnd, (HMENU)ID_CHK_AUTO, GetModuleHandle(NULL), NULL);
        CreateWindowExW(0, L"BUTTON", L"Notifications",
            WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
            200, 140, 150, 25, hwnd, (HMENU)ID_CHK_NOTIFY, GetModuleHandle(NULL), NULL);
        CreateWindowExW(0, L"BUTTON", L"High Speed",
            WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
            30, 170, 150, 25, hwnd, (HMENU)ID_CHK_HIGHSPEED, GetModuleHandle(NULL), NULL);
        
        // Rules label
        CreateWindowExW(0, L"STATIC", L"RULES",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            30, 205, 100, 20, hwnd, NULL, GetModuleHandle(NULL), NULL);
        
        // Rules list
        CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", L"",
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | LBS_NOTIFY,
            30, 230, 320, 100, hwnd, (HMENU)ID_LIST_RULES, GetModuleHandle(NULL), NULL);
        
        // Rule buttons
        CreateWindowExW(0, L"BUTTON", L"Add",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            30, 340, 100, 30, hwnd, (HMENU)ID_BTN_ADDRULE, GetModuleHandle(NULL), NULL);
        CreateWindowExW(0, L"BUTTON", L"Edit",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            140, 340, 100, 30, hwnd, (HMENU)ID_BTN_EDITRULE, GetModuleHandle(NULL), NULL);
        CreateWindowExW(0, L"BUTTON", L"Delete",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            250, 340, 100, 30, hwnd, (HMENU)ID_BTN_DELRULE, GetModuleHandle(NULL), NULL);
        
        UpdateMainWindow();
        return 0;
    }
    
    case WM_DRAWITEM: {
        DRAWITEMSTRUCT* dis = (DRAWITEMSTRUCT*)lParam;
        if (dis->CtlID >= ID_BTN_1000 && dis->CtlID <= ID_BTN_8000) {
            int rate = (dis->CtlID == ID_BTN_1000) ? 1000 : (dis->CtlID == ID_BTN_4000 ? 4000 : 8000);
            bool selected = (g_app.currentRate == rate);
            COLORREF color;
            const wchar_t* text;
            if (rate == 1000) { color = RGB(40, 180, 70); text = L"1000 Hz"; }
            else if (rate == 4000) { color = CLR_ORANGE; text = L"4000 Hz"; }
            else { color = RGB(220, 40, 20); text = L"8000 Hz"; }
            
            HFONT font = CreateFontW(18, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
                DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Consolas");
            SelectObject(dis->hDC, font);
            DrawRateButton(dis->hDC, dis->rcItem, text, selected, color);
            DeleteObject(font);
            return TRUE;
        }
        break;
    }
    
    case WM_CTLCOLORSTATIC: {
        HDC hdc = (HDC)wParam;
        SetBkColor(hdc, CLR_BG);
        SetTextColor(hdc, CLR_DARK);
        
        HWND ctrl = (HWND)lParam;
        wchar_t className[32];
        GetClassNameW(ctrl, className, 32);
        
        wchar_t text[64];
        GetWindowTextW(ctrl, text, 64);
        if (wcscmp(text, L"LAMZU POLLING RATE") == 0) {
            SetTextColor(hdc, CLR_ORANGE);
            SelectObject(hdc, g_app.titleFont);
        } else if (wcscmp(text, L"RULES") == 0) {
            SetTextColor(hdc, CLR_ORANGE);
            SelectObject(hdc, g_app.boldFont);
        }
        
        return (LRESULT)g_app.bgBrush;
    }
    
    case WM_CTLCOLORBTN: {
        return (LRESULT)g_app.bgBrush;
    }
    
    case WM_ERASEBKGND: {
        HDC hdc = (HDC)wParam;
        RECT rc;
        GetClientRect(hwnd, &rc);
        FillRect(hdc, &rc, g_app.bgBrush);
        
        // Orange accent line at top
        RECT accent = {0, 0, rc.right, 5};
        FillRect(hdc, &accent, g_app.orangeBrush);
        
        return 1;
    }
    
    case WM_COMMAND: {
        switch (LOWORD(wParam)) {
        case ID_BTN_1000: SetRate(1000); break;
        case ID_BTN_4000: SetRate(4000); break;
        case ID_BTN_8000: SetRate(8000); break;
        case ID_CHK_AUTO:
            g_app.autoEnabled = (IsDlgButtonChecked(hwnd, ID_CHK_AUTO) == BST_CHECKED);
            SaveConfig();
            UpdateTrayIcon();
            break;
        case ID_CHK_NOTIFY:
            g_app.notificationsEnabled = (IsDlgButtonChecked(hwnd, ID_CHK_NOTIFY) == BST_CHECKED);
            SaveConfig();
            break;
        case ID_CHK_HIGHSPEED:
            SetHighSpeed(IsDlgButtonChecked(hwnd, ID_CHK_HIGHSPEED) == BST_CHECKED);
            break;
        case ID_BTN_ADDRULE:
            ShowRuleDialog(-1);
            UpdateMainWindow();
            break;
        case ID_BTN_EDITRULE: {
            HWND list = GetDlgItem(hwnd, ID_LIST_RULES);
            int sel = (int)SendMessageW(list, LB_GETCURSEL, 0, 0);
            if (sel != LB_ERR) {
                ShowRuleDialog(sel);
                UpdateMainWindow();
            }
            break;
        }
        case ID_BTN_DELRULE: {
            HWND list = GetDlgItem(hwnd, ID_LIST_RULES);
            int sel = (int)SendMessageW(list, LB_GETCURSEL, 0, 0);
            if (sel != LB_ERR && sel < (int)g_app.rules.size()) {
                g_app.rules.erase(g_app.rules.begin() + sel);
                SaveConfig();
                UpdateMainWindow();
            }
            break;
        }
        case ID_LIST_RULES:
            if (HIWORD(wParam) == LBN_DBLCLK) {
                HWND list = GetDlgItem(hwnd, ID_LIST_RULES);
                int sel = (int)SendMessageW(list, LB_GETCURSEL, 0, 0);
                if (sel != LB_ERR) {
                    ShowRuleDialog(sel);
                    UpdateMainWindow();
                }
            }
            break;
        }
        return 0;
    }
    
    case WM_CLOSE:
        ShowWindow(hwnd, SW_HIDE);
        g_app.mainWndVisible = false;
        return 0;
    }
    
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

void ShowMainWindow() {
    if (!g_app.mainWnd) {
        WNDCLASSW wc = {};
        wc.lpfnWndProc = MainWndProc;
        wc.hInstance = GetModuleHandle(NULL);
        wc.lpszClassName = L"LAMZUMain";
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground = g_app.bgBrush;
        RegisterClassW(&wc);
        
        g_app.mainWnd = CreateWindowExW(
            WS_EX_APPWINDOW,
            L"LAMZUMain", L"LAMZU Polling Rate",
            WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
            CW_USEDEFAULT, CW_USEDEFAULT, 395, 425,
            NULL, NULL, GetModuleHandle(NULL), NULL
        );
    }
    
    ShowWindow(g_app.mainWnd, SW_SHOW);
    SetForegroundWindow(g_app.mainWnd);
    g_app.mainWndVisible = true;
    UpdateMainWindow();
}

void ShowRuleDialog(int editIndex) {
    HINSTANCE hInst = GetModuleHandle(NULL);
    
    HWND dlg = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_TOPMOST, L"#32770", 
        editIndex >= 0 ? L"Edit Rule" : L"Add Rule",
        WS_VISIBLE | WS_POPUP | WS_CAPTION | WS_SYSMENU,
        CW_USEDEFAULT, CW_USEDEFAULT, 300, 150, 
        g_app.mainWnd ? g_app.mainWnd : g_app.hwnd, NULL, hInst, NULL);
    
    CreateWindowExW(0, L"STATIC", L"Process:", WS_CHILD | WS_VISIBLE,
        15, 18, 60, 20, dlg, NULL, hInst, NULL);
    HWND hProcess = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", 
        WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
        80, 15, 190, 24, dlg, (HMENU)101, hInst, NULL);
    
    CreateWindowExW(0, L"STATIC", L"Rate:", WS_CHILD | WS_VISIBLE,
        15, 50, 60, 20, dlg, NULL, hInst, NULL);
    HWND hRate = CreateWindowExW(0, L"COMBOBOX", L"", 
        WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST,
        80, 47, 190, 100, dlg, (HMENU)102, hInst, NULL);
    
    SendMessageW(hRate, CB_ADDSTRING, 0, (LPARAM)L"1000 Hz");
    SendMessageW(hRate, CB_ADDSTRING, 0, (LPARAM)L"4000 Hz");
    SendMessageW(hRate, CB_ADDSTRING, 0, (LPARAM)L"8000 Hz");
    
    CreateWindowExW(0, L"BUTTON", L"Save", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
        50, 85, 80, 28, dlg, (HMENU)103, hInst, NULL);
    CreateWindowExW(0, L"BUTTON", L"Cancel", WS_CHILD | WS_VISIBLE,
        150, 85, 80, 28, dlg, (HMENU)104, hInst, NULL);
    
    if (editIndex >= 0 && editIndex < (int)g_app.rules.size()) {
        SetWindowTextW(hProcess, g_app.rules[editIndex].process.c_str());
        int sel = g_app.rules[editIndex].rate == 8000 ? 2 : (g_app.rules[editIndex].rate == 4000 ? 1 : 0);
        SendMessageW(hRate, CB_SETCURSEL, sel, 0);
    } else {
        SetWindowTextW(hProcess, L"cs2.exe");
        SendMessageW(hRate, CB_SETCURSEL, 1, 0);
    }
    
    MSG msg;
    while (GetMessageW(&msg, NULL, 0, 0)) {
        if (!IsWindow(dlg)) break;
        
        if (msg.message == WM_KEYDOWN && msg.wParam == VK_ESCAPE) {
            DestroyWindow(dlg);
            break;
        }
        
        if (msg.message == WM_LBUTTONUP) {
            if (msg.hwnd == GetDlgItem(dlg, 103)) {
                wchar_t proc[256];
                GetWindowTextW(hProcess, proc, 256);
                int sel = (int)SendMessageW(hRate, CB_GETCURSEL, 0, 0);
                int rate = sel == 2 ? 8000 : (sel == 1 ? 4000 : 1000);
                
                if (wcslen(proc) > 0) {
                    if (editIndex >= 0 && editIndex < (int)g_app.rules.size()) {
                        g_app.rules[editIndex].process = proc;
                        g_app.rules[editIndex].rate = rate;
                    } else {
                        g_app.rules.push_back({proc, rate});
                    }
                    SaveConfig();
                }
                DestroyWindow(dlg);
                break;
            }
            if (msg.hwnd == GetDlgItem(dlg, 104)) {
                DestroyWindow(dlg);
                break;
            }
        }
        
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_TRAYICON:
        if (lParam == WM_RBUTTONUP) {
            ShowContextMenu(hwnd);
        } else if (lParam == WM_LBUTTONDBLCLK) {
            ShowMainWindow();
        }
        return 0;
        
    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case ID_TRAY_OPEN: ShowMainWindow(); break;
        case ID_TRAY_1000: SetRate(1000); break;
        case ID_TRAY_4000: SetRate(4000); break;
        case ID_TRAY_8000: SetRate(8000); break;
        case ID_TRAY_AUTO:
            g_app.autoEnabled = !g_app.autoEnabled;
            SaveConfig();
            UpdateTrayIcon();
            break;
        case ID_TRAY_NOTIFY:
            g_app.notificationsEnabled = !g_app.notificationsEnabled;
            SaveConfig();
            break;
        case ID_TRAY_HIGHSPEED:
            SetHighSpeed(!g_app.highSpeedEnabled);
            break;
        case ID_TRAY_ADDRULE:
            ShowRuleDialog(-1);
            break;
        case ID_TRAY_EXIT:
            g_app.running = false;
            Shell_NotifyIconW(NIM_DELETE, &g_app.nid);
            PostQuitMessage(0);
            break;
        default:
            if (LOWORD(wParam) >= ID_TRAY_RULE_BASE && 
                LOWORD(wParam) < ID_TRAY_RULE_BASE + g_app.rules.size()) {
                ShowRuleDialog(LOWORD(wParam) - ID_TRAY_RULE_BASE);
            }
            break;
        }
        return 0;
        
    case WM_DESTROY:
        g_app.running = false;
        Shell_NotifyIconW(NIM_DELETE, &g_app.nid);
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, PWSTR, int) {
    // Init resources
    g_app.bgBrush = CreateSolidBrush(CLR_BG);
    g_app.darkBrush = CreateSolidBrush(CLR_DARK);
    g_app.orangeBrush = CreateSolidBrush(CLR_ORANGE);
    g_app.titleFont = CreateFontW(24, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Consolas");
    g_app.normalFont = CreateFontW(14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    g_app.boldFont = CreateFontW(14, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Consolas");
    
    g_app.configPath = GetConfigPath();
    LoadConfig();
    g_app.devicePath = FindDevice();
    
    SetRate(g_app.currentRate, false, false);
    SetHighSpeed(g_app.highSpeedEnabled, false, false);
    
    g_app.icon1k = CreateRateIcon(1000);
    g_app.icon4k = CreateRateIcon(4000);
    g_app.icon8k = CreateRateIcon(8000);
    
    WNDCLASSW wc = {};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"LAMZUPollingTray";
    RegisterClassW(&wc);
    
    g_app.hwnd = CreateWindowExW(0, L"LAMZUPollingTray", L"", 0, 0, 0, 0, 0, 
        HWND_MESSAGE, NULL, hInstance, NULL);
    
    g_app.nid = {};
    g_app.nid.cbSize = sizeof(NOTIFYICONDATA);
    g_app.nid.hWnd = g_app.hwnd;
    g_app.nid.uID = 1;
    g_app.nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    g_app.nid.uCallbackMessage = WM_TRAYICON;
    
    HICON icon = g_app.currentRate == 8000 ? g_app.icon8k : 
                 (g_app.currentRate == 4000 ? g_app.icon4k : g_app.icon1k);
    g_app.nid.hIcon = icon;
    
    std::wstring tip = L"LAMZU: " + std::to_wstring(g_app.currentRate) + L" Hz";
    if (g_app.autoEnabled) tip += L" [AUTO]";
    wcscpy_s(g_app.nid.szTip, tip.c_str());
    
    Shell_NotifyIconW(NIM_ADD, &g_app.nid);
    
    std::thread monitor(MonitorThread);
    monitor.detach();
    
    MSG msg;
    while (GetMessageW(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    
    // Cleanup
    DeleteObject(g_app.bgBrush);
    DeleteObject(g_app.darkBrush);
    DeleteObject(g_app.orangeBrush);
    DeleteObject(g_app.titleFont);
    DeleteObject(g_app.normalFont);
    DeleteObject(g_app.boldFont);
    DestroyIcon(g_app.icon1k);
    DestroyIcon(g_app.icon4k);
    DestroyIcon(g_app.icon8k);
    
    return 0;
}
