#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif

#include <windows.h>
#include <commctrl.h>
#include <commdlg.h>
#include <shellapi.h>
#include <stdio.h>
#include <stdlib.h>
#include <wchar.h>
#include "excel_cleaner.h"

#pragma comment(lib, "comctl32.lib")

#define ID_BTN_BROWSE       101
#define ID_BTN_START        102
#define ID_BTN_OPEN_FOLDER  103
#define ID_BTN_OPEN_FILE    104
#define ID_RADIO_DEFAULT    105
#define ID_RADIO_ADD        106
#define ID_RADIO_MULTIPLY   107
#define ID_EDIT_FILEPATH    108
#define ID_EDIT_LOG         109
#define ID_PROGRESS         110

// Global Control Handles
HWND g_hEditFilePath = NULL;
HWND g_hRadioDefault = NULL;
HWND g_hRadioAdd = NULL;
HWND g_hRadioMultiply = NULL;
HWND g_hBtnStart = NULL;
HWND g_hBtnOpenFolder = NULL;
HWND g_hBtnOpenFile = NULL;
HWND g_hProgressBar = NULL;
HWND g_hEditLog = NULL;
HFONT g_hFontTitle = NULL;
HFONT g_hFontNormal = NULL;
HFONT g_hFontBold = NULL;

wchar_t g_szCleanedFilePath[MAX_PATH] = {0};
BOOL g_bIsProcessing = FALSE;

// Helper to log text to edit control
void AppendLog(const wchar_t* text) {
    if (!g_hEditLog) return;
    int len = GetWindowTextLengthW(g_hEditLog);
    SendMessageW(g_hEditLog, EM_SETSEL, (WPARAM)len, (LPARAM)len);
    SendMessageW(g_hEditLog, EM_REPLACESEL, FALSE, (LPARAM)text);
    SendMessageW(g_hEditLog, EM_REPLACESEL, FALSE, (LPARAM)L"\r\n");
    SendMessageW(g_hEditLog, WM_VSCROLL, SB_BOTTOM, 0);
}

// Progress callback from pure C excel engine
static void ProgressCallbackFunc(int percentage, const wchar_t* status_msg) {
    SendMessageW(g_hProgressBar, PBM_SETPOS, percentage, 0);
    AppendLog(status_msg);
}

// Background thread structure
typedef struct {
    wchar_t szFilePath[MAX_PATH];
    wchar_t szMode[32];
} ThreadParams;

DWORD WINAPI WorkerThread(LPVOID lpParam) {
    ThreadParams* params = (ThreadParams*)lpParam;
    
    AppendLog(L"=== [纯 C 独立引擎] 开始执行 Excel 数据清洗任务 ===");
    
    CleanMode mode = CLEAN_MODE_DEFAULT;
    if (wcscmp(params->szMode, L"add") == 0) {
        mode = CLEAN_MODE_ADD;
    } else if (wcscmp(params->szMode, L"multiply") == 0) {
        mode = CLEAN_MODE_MULTIPLY;
    }

    wchar_t out_path[MAX_PATH] = {0};
    int res = process_excel_c(params->szFilePath, mode, out_path, ProgressCallbackFunc);

    g_bIsProcessing = FALSE;
    EnableWindow(g_hBtnStart, TRUE);

    if (res == 0) {
        wcscpy(g_szCleanedFilePath, out_path);
        SendMessageW(g_hProgressBar, PBM_SETPOS, 100, 0);
        AppendLog(L"========================================");
        AppendLog(L"🎉 清洗完成！文件已成功生成！");
        wchar_t successMsg[MAX_PATH + 64];
        swprintf(successMsg, MAX_PATH + 64, L"输出文件: %s", out_path);
        AppendLog(successMsg);
        EnableWindow(g_hBtnOpenFolder, TRUE);
        EnableWindow(g_hBtnOpenFile, TRUE);
        MessageBoxW(NULL, L"数据清洗完成！\n已生成带有 _Cleaned 后缀的 Excel 文件。", L"成功", MB_OK | MB_ICONINFORMATION);
    } else {
        AppendLog(L"========================================");
        AppendLog(L"❌ 清洗过程出现错误，请检查日志。");
        MessageBoxW(NULL, L"数据清洗失败，请查看主界面日志输出。", L"错误", MB_OK | MB_ICONERROR);
    }

    free(params);
    return 0;
}

void OnBrowseFile(HWND hwnd) {
    wchar_t szFile[MAX_PATH] = {0};
    OPENFILENAMEW ofn;
    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hwnd;
    ofn.lpstrFilter = L"Excel 文件 (*.xlsx;*.xls)\0*.xlsx;*.xls\0所有文件 (*.*)\0*.*\0";
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = MAX_PATH;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

    if (GetOpenFileNameW(&ofn)) {
        SetWindowTextW(g_hEditFilePath, szFile);
    }
}

void OnStartCleaning(HWND hwnd) {
    if (g_bIsProcessing) return;

    wchar_t szFilePath[MAX_PATH];
    GetWindowTextW(g_hEditFilePath, szFilePath, MAX_PATH);

    if (wcslen(szFilePath) == 0) {
        MessageBoxW(hwnd, L"请先选择要清洗的 Excel 文件！", L"提示", MB_OK | MB_ICONWARNING);
        return;
    }

    if (GetFileAttributesW(szFilePath) == INVALID_FILE_ATTRIBUTES) {
        MessageBoxW(hwnd, L"选定的文件不存在，请检查路径！", L"错误", MB_OK | MB_ICONERROR);
        return;
    }

    wchar_t szMode[32] = L"default";
    if (SendMessageW(g_hRadioAdd, BM_GETCHECK, 0, 0) == BST_CHECKED) {
        wcscpy(szMode, L"add");
    } else if (SendMessageW(g_hRadioMultiply, BM_GETCHECK, 0, 0) == BST_CHECKED) {
        wcscpy(szMode, L"multiply");
    }

    g_bIsProcessing = TRUE;
    EnableWindow(g_hBtnStart, FALSE);
    EnableWindow(g_hBtnOpenFolder, FALSE);
    EnableWindow(g_hBtnOpenFile, FALSE);
    SendMessageW(g_hProgressBar, PBM_SETPOS, 0, 0);
    SetWindowTextW(g_hEditLog, L"");

    ThreadParams* params = (ThreadParams*)malloc(sizeof(ThreadParams));
    wcscpy(params->szFilePath, szFilePath);
    wcscpy(params->szMode, szMode);

    HANDLE hThread = CreateThread(NULL, 0, WorkerThread, params, 0, NULL);
    if (hThread) {
        CloseHandle(hThread);
    }
}

void OnOpenFolder() {
    if (wcslen(g_szCleanedFilePath) > 0) {
        wchar_t szDir[MAX_PATH];
        wcscpy(szDir, g_szCleanedFilePath);
        wchar_t* lastSlash = wcsrchr(szDir, L'\\');
        if (lastSlash) *lastSlash = L'\0';
        ShellExecuteW(NULL, L"open", szDir, NULL, NULL, SW_SHOW);
    }
}

void OnOpenFile() {
    if (wcslen(g_szCleanedFilePath) > 0) {
        ShellExecuteW(NULL, L"open", g_szCleanedFilePath, NULL, NULL, SW_SHOW);
    }
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE: {
        // Create Fonts
        g_hFontTitle = CreateFontW(-18, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                                   OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                                   DEFAULT_PITCH | FF_DONTCARE, L"Microsoft YaHei");
        g_hFontNormal = CreateFontW(-14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                                    OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                                    DEFAULT_PITCH | FF_DONTCARE, L"Microsoft YaHei");
        g_hFontBold = CreateFontW(-14, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                                  OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                                  DEFAULT_PITCH | FF_DONTCARE, L"Microsoft YaHei");

        // Section 1: Title & File Selector
        HWND hLblTitle = CreateWindowW(L"STATIC", L"📊 MSYS2 纯 C 独立数据清洗整理工具",
                                      WS_VISIBLE | WS_CHILD,
                                      20, 15, 600, 30, hwnd, NULL, NULL, NULL);
        SendMessageW(hLblTitle, WM_SETFONT, (WPARAM)g_hFontTitle, TRUE);

        HWND hGroupFile = CreateWindowW(L"BUTTON", L" 1. 数据源文件选择 ",
                                       WS_VISIBLE | WS_CHILD | BS_GROUPBOX,
                                       20, 50, 680, 75, hwnd, NULL, NULL, NULL);
        SendMessageW(hGroupFile, WM_SETFONT, (WPARAM)g_hFontBold, TRUE);

        g_hEditFilePath = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                                          WS_VISIBLE | WS_CHILD | ES_AUTOHSCROLL,
                                          35, 80, 520, 28, hwnd, (HMENU)ID_EDIT_FILEPATH, NULL, NULL);
        SendMessageW(g_hEditFilePath, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);

        HWND hBtnBrowse = CreateWindowW(L"BUTTON", L"浏览...",
                                        WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
                                        565, 80, 115, 28, hwnd, (HMENU)ID_BTN_BROWSE, NULL, NULL);
        SendMessageW(hBtnBrowse, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);

        // Section 2: Cleaning Mode Selection
        HWND hGroupMode = CreateWindowW(L"BUTTON", L" 2. 清理模式选择 ",
                                       WS_VISIBLE | WS_CHILD | BS_GROUPBOX,
                                       20, 135, 680, 130, hwnd, NULL, NULL, NULL);
        SendMessageW(hGroupMode, WM_SETFONT, (WPARAM)g_hFontBold, TRUE);

        g_hRadioDefault = CreateWindowW(L"BUTTON", L"默认清理模式（提取单元格数字，包含负数/小数，多个以逗号分隔；无数字兜底保留）",
                                         WS_VISIBLE | WS_CHILD | BS_AUTORADIOBUTTON | WS_GROUP,
                                         35, 160, 650, 24, hwnd, (HMENU)ID_RADIO_DEFAULT, NULL, NULL);
        SendMessageW(g_hRadioDefault, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);
        SendMessageW(g_hRadioDefault, BM_SETCHECK, BST_CHECKED, 0);

        g_hRadioAdd = CreateWindowW(L"BUTTON", L"相加模式（提取单元格所有数字求和后填入；无数字兜底保留）",
                                     WS_VISIBLE | WS_CHILD | BS_AUTORADIOBUTTON,
                                     35, 190, 650, 24, hwnd, (HMENU)ID_RADIO_ADD, NULL, NULL);
        SendMessageW(g_hRadioAdd, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);

        g_hRadioMultiply = CreateWindowW(L"BUTTON", L"相乘模式（提取单元格所有数字求积后填入；无数字兜底保留）",
                                          WS_VISIBLE | WS_CHILD | BS_AUTORADIOBUTTON,
                                          35, 220, 650, 24, hwnd, (HMENU)ID_RADIO_MULTIPLY, NULL, NULL);
        SendMessageW(g_hRadioMultiply, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);

        // Section 3: Action Buttons
        g_hBtnStart = CreateWindowW(L"BUTTON", L"🚀 开始清理数据",
                                    WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
                                    20, 275, 200, 36, hwnd, (HMENU)ID_BTN_START, NULL, NULL);
        SendMessageW(g_hBtnStart, WM_SETFONT, (WPARAM)g_hFontBold, TRUE);

        g_hBtnOpenFolder = CreateWindowW(L"BUTTON", L"📁 打开输出文件夹",
                                         WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON | WS_DISABLED,
                                         230, 275, 180, 36, hwnd, (HMENU)ID_BTN_OPEN_FOLDER, NULL, NULL);
        SendMessageW(g_hBtnOpenFolder, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);

        g_hBtnOpenFile = CreateWindowW(L"BUTTON", L"📈 查看清洗结果 Excel",
                                       WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON | WS_DISABLED,
                                       420, 275, 180, 36, hwnd, (HMENU)ID_BTN_OPEN_FILE, NULL, NULL);
        SendMessageW(g_hBtnOpenFile, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);

        // Section 4: Progress Bar & Log
        g_hProgressBar = CreateWindowExW(0, PROGRESS_CLASSW, NULL,
                                         WS_VISIBLE | WS_CHILD,
                                         20, 320, 680, 20, hwnd, (HMENU)ID_PROGRESS, NULL, NULL);
        SendMessageW(g_hProgressBar, PBM_SETRANGE, 0, MAKELPARAM(0, 100));

        g_hEditLog = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                                     WS_VISIBLE | WS_CHILD | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY | WS_VSCROLL,
                                     20, 350, 680, 160, hwnd, (HMENU)ID_EDIT_LOG, NULL, NULL);
        SendMessageW(g_hEditLog, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);
        AppendLog(L"[纯 C 原生程序] 就绪。无需 Python 环境即可独立运行。");
        break;
    }
    case WM_COMMAND: {
        int wID = LOWORD(wParam);
        if (wID == ID_BTN_BROWSE) {
            OnBrowseFile(hwnd);
        } else if (wID == ID_BTN_START) {
            OnStartCleaning(hwnd);
        } else if (wID == ID_BTN_OPEN_FOLDER) {
            OnOpenFolder();
        } else if (wID == ID_BTN_OPEN_FILE) {
            OnOpenFile();
        }
        break;
    }
    case WM_DESTROY:
        if (g_hFontTitle) DeleteObject(g_hFontTitle);
        if (g_hFontNormal) DeleteObject(g_hFontNormal);
        if (g_hFontBold) DeleteObject(g_hFontBold);
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    }
    return 0;
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, PWSTR pCmdLine, int nCmdShow) {
    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_PROGRESS_CLASS | ICC_STANDARD_CLASSES;
    InitCommonControlsEx(&icex);

    const wchar_t CLASS_NAME[] = L"DataCleanerWindowClass";

    WNDCLASSW wc = {0};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = CLASS_NAME;
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);

    RegisterClassW(&wc);

    HWND hwnd = CreateWindowExW(
        0, CLASS_NAME, L"纯 C 语言数据清洗整理 GUI 程序 (MSYS2 UCRT64)",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
        CW_USEDEFAULT, CW_USEDEFAULT, 735, 560,
        NULL, NULL, hInstance, NULL
    );

    if (hwnd == NULL) {
        return 0;
    }

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    MSG msg = {0};
    while (GetMessageW(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    return 0;
}
