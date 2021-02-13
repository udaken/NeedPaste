#include "stdafx.h"
#include "resource.h"

#include <string>
#include <vector>

#include <wil/win32_helpers.h>

#pragma comment(linker, "/manifestdependency:\"type='win32' \
    name='Microsoft.Windows.Common-Controls' \
    version='6.0.0.0' \
    processorArchitecture='*' \
    publicKeyToken='6595b64144ccf1df' \
    language='*'\"")

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "Comctl32.lib")

#define MAX_LOADSTRING 100
#define WM_NOTIFYICON (WM_USER+100)

constexpr UINT NOTIFY_UID = 1;
constexpr INT HOTKEY_ID = 1;
constexpr auto szWindowClass{ L"{B969DAC2-3F28-4398-8D23-C35186DD50B2}" };

static LRESULT CALLBACK    wndProc(HWND, UINT, WPARAM, LPARAM) noexcept;

HINSTANCE g_hInst;
WCHAR g_szTitle[MAX_LOADSTRING] = L"NeedPaste";

struct HOTKEYINFO {
    UINT fsModifiers;
    UINT vk;
} g_hotkey = { MOD_CONTROL | MOD_SHIFT, L'V' };

bool g_clearAfterPaste = false;

// ============================================================================
// ============================================================================
BOOL addNotifyIcon(HWND hWnd, unsigned int uID) noexcept
{
    NOTIFYICONDATA nid{
        .cbSize = sizeof(nid),
        .hWnd = hWnd,
        .uID = uID,
        .uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP,
        .uCallbackMessage = WM_NOTIFYICON,
        .hIcon = LoadIcon(g_hInst, MAKEINTRESOURCE(IDI_SMALL)),
    };

    _tcscpy_s(nid.szTip, g_szTitle);

    return Shell_NotifyIcon(NIM_ADD, &nid);
}

// ============================================================================
// ============================================================================
void deleteNotifyIcon(HWND hWnd, unsigned int uID) noexcept
{
    NOTIFYICONDATA nid{
        .cbSize = sizeof(nid),
        .hWnd = hWnd,
        .uID = uID,
    };

    Shell_NotifyIcon(NIM_DELETE, &nid);
}

// ============================================================================
// ============================================================================
void registerHotKey(HWND hWnd, const HOTKEYINFO& hotkeyInfo)
{
    if (!RegisterHotKey(hWnd, HOTKEY_ID, hotkeyInfo.fsModifiers, hotkeyInfo.vk))
    {
        MessageBox(hWnd, L"Failed To RegisterHotKey.", g_szTitle, MB_ICONEXCLAMATION | MB_OK);
        FAIL_FAST_LAST_ERROR();
    }
}

// ============================================================================
// ============================================================================
void unregisterHotKey(HWND hWnd)
{
    FAIL_FAST_IF(!UnregisterHotKey(hWnd, HOTKEY_ID));
}

// ============================================================================
// ============================================================================
static std::wstring getModuleName()
{
    WCHAR path[MAX_PATH] = {};
    GetModuleFileName(nullptr, path, _countof(path));
    return path;
}

// ============================================================================
// ============================================================================
void inputString(LPCWSTR text) noexcept
{
    std::vector<INPUT> inputs;
    for (size_t i = 0; i < wcslen(text); i++)
    {
        INPUT input = { INPUT_KEYBOARD };
        input.ki.wScan = text[i];
        input.ki.dwFlags = KEYEVENTF_UNICODE; // key down
        inputs.push_back(input);

        input.ki.dwFlags |= KEYEVENTF_KEYUP;
        inputs.push_back(input);
    }

    UINT ret = SendInput(static_cast<UINT>(inputs.size()), inputs.data(), sizeof(INPUT));
    if (ret == 0)
    {
        wprintf(L"SendInput Error: 0x%x\n", ::GetLastError());
    }
}
// ============================================================================
// ============================================================================
int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
    _In_opt_ HINSTANCE hPrevInstance,
    _In_ LPWSTR    lpCmdLine,
    _In_ int       /*nCmdShow*/)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    g_hInst = hInstance;

    FAIL_FAST_IF_FAILED(CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE));

    INITCOMMONCONTROLSEX iccx{
        .dwSize = sizeof(INITCOMMONCONTROLSEX),
        .dwICC = ICC_HOTKEY_CLASS,
    };
    FAIL_FAST_IF(!InitCommonControlsEx(&iccx));

    WNDCLASSEXW wcex{
        .cbSize = sizeof(WNDCLASSEX),
        .style = CS_HREDRAW | CS_VREDRAW,
        .lpfnWndProc = &wndProc,
        .cbClsExtra = 0,
        .cbWndExtra = 0,
        .hInstance = hInstance,
        .hCursor = LoadCursor(nullptr, IDC_ARROW),
        .hbrBackground = (HBRUSH)(COLOR_WINDOW + 1),
        .lpszMenuName = nullptr,
        .lpszClassName = szWindowClass,
        //.hIconSm = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL)),
    };

    FAIL_FAST_IF(!RegisterClassExW(&wcex) != 0);

    HWND hWnd = CreateWindow(szWindowClass, g_szTitle, 0,
        CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

    FAIL_FAST_IF_NULL(hWnd);

    registerHotKey(hWnd, g_hotkey);

    if (!addNotifyIcon(hWnd, NOTIFY_UID))
    {
        return false;
    }

    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return (int)msg.wParam;
}

// ============================================================================
// ============================================================================
constexpr BYTE toHotkeyModifierFlag(DWORD val)
{
    BYTE f = 0;
    if (val & MOD_ALT)
        f |= HOTKEYF_ALT;
    if (val & MOD_SHIFT)
        f |= HOTKEYF_SHIFT;
    if (val & MOD_CONTROL)
        f |= HOTKEYF_CONTROL;
    return f;
}

// ============================================================================
// ============================================================================
constexpr DWORD fromHotkeyModifierFlag(BYTE val)
{
    DWORD f = 0;
    if (val & HOTKEYF_ALT)
        f |= MOD_ALT;
    if (val & HOTKEYF_SHIFT)
        f |= MOD_SHIFT;
    if (val & HOTKEYF_CONTROL)
        f |= MOD_CONTROL;
    return f;
}

// ============================================================================
// ============================================================================
INT_PTR CALLBACK settingDlgProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    switch (message)
    {
    case WM_INITDIALOG:
    {
        SetWindowText(hWnd, g_szTitle);
        SendMessage(hWnd, WM_SETICON, ICON_SMALL, (LPARAM)LoadIcon(g_hInst, MAKEINTRESOURCE(IDI_SMALL)));
        auto fwCombInv = HKCOMB_NONE | HKCOMB_S | HKCOMB_C;
        WORD hotkey = MAKEWORD(g_hotkey.vk, toHotkeyModifierFlag(g_hotkey.fsModifiers));
        SendDlgItemMessage(hWnd, IDC_HOTKEY1, HKM_SETRULES, fwCombInv, MAKELPARAM(hotkey, 0));
        SendDlgItemMessage(hWnd, IDC_HOTKEY1, HKM_SETHOTKEY, hotkey, 0);
        CheckDlgButton(hWnd, IDC_CHECK1, g_clearAfterPaste ? BST_CHECKED : BST_UNCHECKED);
        return (INT_PTR)TRUE;
    }
    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            if (LOWORD(wParam) == IDOK)
            {
                WORD hotkey = SendDlgItemMessage(hWnd, IDC_HOTKEY1, HKM_GETHOTKEY, 0, 0);
                if (hotkey != 0)
                {
                    g_hotkey.fsModifiers = fromHotkeyModifierFlag(HIBYTE(hotkey));
                    g_hotkey.vk = LOBYTE(hotkey);
                }
                auto check = IsDlgButtonChecked(hWnd, IDC_CHECK1);
                g_clearAfterPaste = check;
            }
            EndDialog(hWnd, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

// ============================================================================
// ============================================================================
bool toggleMenuItem(HMENU hMenu, UINT menuItem)
{
    UINT fdwMenu = GetMenuState(hMenu, menuItem, MF_BYCOMMAND);
    auto newState = ((fdwMenu & MF_CHECKED) ? MF_UNCHECKED : MF_CHECKED);
    CheckMenuItem(hMenu, menuItem, MF_BYCOMMAND | newState);
    return newState;
}

void inputClipbordText(HWND hWnd)
{
    if (OpenClipboard(hWnd))
    {
        auto hText = GetClipboardData(CF_UNICODETEXT);
        if (hText) {
            auto pText = (LPCWSTR)GlobalLock(hText);
            inputString(pText);
            GlobalUnlock(hText);

            if (g_clearAfterPaste)
                EmptyClipboard();
            CloseClipboard();
        }
    }
    else
    {
        MessageBox(hWnd, L"Failed To Open Clipboard.", g_szTitle, MB_ICONEXCLAMATION | MB_OK);
    }
}

// ============================================================================
// ============================================================================
LRESULT CALLBACK wndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    static UINT s_uTaskbarRestart;
    static HMENU menu = nullptr;
    switch (message)
    {
    case WM_CREATE:
        s_uTaskbarRestart = RegisterWindowMessage(L"TaskbarCreated");
        FAIL_FAST_IF_NULL(menu = LoadMenu(g_hInst, MAKEINTRESOURCE(IDR_MENU1)));
        break;
    case WM_HOTKEY:
    {
        if (HOTKEY_ID == wParam)
        {
            inputClipbordText(hWnd);
            break;
        }
        else
        {
            return DefWindowProc(hWnd, message, wParam, lParam);
        }
    }
    case WM_DESTROY:
        unregisterHotKey(hWnd);
        deleteNotifyIcon(hWnd, NOTIFY_UID);
        PostQuitMessage(0);
        DestroyMenu(menu);
        return 0;
    case WM_NOTIFYICON:
        switch (lParam)
        {
        case WM_RBUTTONDOWN:
            POINT pt{};
            GetCursorPos(&pt);
            SetForegroundWindow(hWnd);
            FAIL_FAST_IF(!TrackPopupMenu(GetSubMenu(menu, 0), TPM_LEFTALIGN, pt.x, pt.y, 0, hWnd, NULL));
            break;
        }
        return 0;
    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case ID_ROOT_EXIT:
            DestroyWindow(hWnd);
            break;
        case ID_ROOT_SETTING:
            unregisterHotKey(hWnd);
            DialogBox(g_hInst, MAKEINTRESOURCE(IDD_DIALOG1), hWnd, &settingDlgProc);
            registerHotKey(hWnd, g_hotkey);
            break;
        case ID_ROOT_ENABLE:
            toggleMenuItem(menu, ID_ROOT_ENABLE);
            break;
        default:
            break;
        }
        break;
    default:
        if (message == s_uTaskbarRestart)
            addNotifyIcon(hWnd, NOTIFY_UID);
        break;
    }
    return DefWindowProc(hWnd, message, wParam, lParam);
}
