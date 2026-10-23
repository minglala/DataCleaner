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
#include <shlobj.h>
#include <stdio.h>
#include <stdlib.h>
#include <wchar.h>
#include "excel_cleaner.h"

#pragma comment(lib, "comctl32.lib")

// ---------------- 控件 ID ----------------
#define IDC_LB_FILES        201
#define IDC_BTN_ADD_FILES   202
#define IDC_BTN_ADD_FOLDER  203
#define IDC_BTN_REMOVE_SEL  204
#define IDC_BTN_CLEAR_LIST  205
#define IDC_COMBO_MODE      207
#define IDC_EDIT_SEP        208
#define IDC_CHK_SPECIAL     209
#define IDC_BTN_START       102
#define IDC_BTN_OPEN_FOLDER 103
#define IDC_BTN_OPEN_FILE   104
#define IDC_EDIT_LOG        109
#define IDC_PROGRESS        110

#define REG_KEY         L"Software\\DataCleaner"
#define MAX_FILES_LIMIT 1000
#define OPEN_BUF_LEN    32768

// ---------------- 模式元数据 ----------------
static const wchar_t* const MODE_NAMES[] = {
    L"数字提取（拼接输出）",
    L"数字求和",
    L"数字求乘积",
    L"数字平均值",
    L"数字最大值",
    L"数字最小值",
    L"数字个数统计",
    L"文本清理",
    L"日期标准化",
};

static const wchar_t* const MODE_DESCRIPTIONS[] = {
    L"提取单元格中的数字（含负数/小数），以指定分隔符连接输出；未发现数字时保留原内容。",
    L"提取单元格中的所有数字并求和填入；未发现数字时保留原内容。",
    L"提取单元格中的所有数字并求乘积填入；未发现数字时保留原内容。",
    L"提取单元格中的所有数字并求平均值填入；未发现数字时保留原内容。",
    L"提取单元格中的所有数字并取最大值填入；未发现数字时保留原内容。",
    L"提取单元格中的所有数字并取最小值填入；未发现数字时保留原内容。",
    L"统计单元格中数字的个数填入；未发现数字时保留原内容。",
    L"清理文本：去首尾空格、全角转半角、合并连续空格、删除控制字符；勾选下方选项后仅保留中文/字母/数字/空格。",
    L"日期标准化：识别 2024-1-1、2024.1.1、2024/1/1、20240101、2024年1月1日（可带时间），统一为 YYYY-MM-DD HH:MM:SS；无法识别时保留原内容。",
};

// ---------------- 全局控件句柄 ----------------
static HWND g_hLblTitle       = NULL;
static HWND g_hGroupFiles     = NULL;
static HWND g_hListBox        = NULL;
static HWND g_hBtnAddFiles    = NULL;
static HWND g_hBtnAddFolder   = NULL;
static HWND g_hBtnRemoveSel   = NULL;
static HWND g_hBtnClearList   = NULL;
static HWND g_hLblFileCount   = NULL;
static HWND g_hGroupMode      = NULL;
static HWND g_hLblModeLabel   = NULL;
static HWND g_hComboMode      = NULL;
static HWND g_hLblSep         = NULL;
static HWND g_hEditSep        = NULL;
static HWND g_hChkSpecial     = NULL;
static HWND g_hLblModeDesc    = NULL;
static HWND g_hBtnStart       = NULL;
static HWND g_hBtnOpenFolder  = NULL;
static HWND g_hBtnOpenFile    = NULL;
static HWND g_hProgressBar    = NULL;
static HWND g_hEditLog        = NULL;

static HFONT g_hFontTitle  = NULL;
static HFONT g_hFontNormal = NULL;
static HFONT g_hFontBold   = NULL;

static wchar_t g_szLastOutput[DC_PATH] = {0};
static wchar_t g_szLastDir[DC_PATH]    = {0};
static BOOL    g_bIsProcessing         = FALSE;

// ---------------- 日志 ----------------
static void AppendLog(const wchar_t* text) {
    if (!g_hEditLog) return;
    int len = GetWindowTextLengthW(g_hEditLog);
    SendMessageW(g_hEditLog, EM_SETSEL, (WPARAM)len, (LPARAM)len);
    SendMessageW(g_hEditLog, EM_REPLACESEL, FALSE, (LPARAM)text);
    SendMessageW(g_hEditLog, EM_REPLACESEL, FALSE, (LPARAM)L"\r\n");
    SendMessageW(g_hEditLog, WM_VSCROLL, SB_BOTTOM, 0);
}

// ---------------- 文件列表管理 ----------------
static BOOL IsSupportedFile(const wchar_t* path) {
    const wchar_t* ext = wcsrchr(path, L'.');
    if (!ext) return FALSE;
    return (_wcsicmp(ext, L".xlsx") == 0 || _wcsicmp(ext, L".csv") == 0);
}

static BOOL IsCleanedOutput(const wchar_t* path) {
    const wchar_t* ext = wcsrchr(path, L'.');
    if (!ext) return FALSE;
    size_t blen = (size_t)(ext - path);
    const wchar_t* suffix = L"_Cleaned";
    size_t slen = 8;
    return (blen >= slen && wcsncmp(path + blen - slen, suffix, slen) == 0);
}

static void UpdateFileCountLabel(void) {
    if (!g_hLblFileCount || !g_hListBox) return;
    int n = (int)SendMessageW(g_hListBox, LB_GETCOUNT, 0, 0);
    wchar_t msg[128];
    swprintf(msg, 128, L"待处理文件：%d 个（支持 .xlsx / .csv，也可直接拖放文件到窗口）", n);
    SetWindowTextW(g_hLblFileCount, msg);
}

static void AddFileToList(const wchar_t* path, int* added, int* skipped) {
    if (!IsSupportedFile(path) || IsCleanedOutput(path)) { (*skipped)++; return; }
    if (SendMessageW(g_hListBox, LB_FINDSTRINGEXACT, (WPARAM)-1, (LPARAM)path) != LB_ERR) {
        (*skipped)++;
        return;
    }
    if (SendMessageW(g_hListBox, LB_GETCOUNT, 0, 0) >= MAX_FILES_LIMIT) { (*skipped)++; return; }
    SendMessageW(g_hListBox, LB_ADDSTRING, 0, (LPARAM)path);
    (*added)++;
}

static void LogAddResult(const wchar_t* what, int added, int skipped) {
    wchar_t msg[256];
    swprintf(msg, 256, L"%s：新增 %d 个，跳过 %d 个（重复 / 不支持的类型 / 已是清洗结果）", what, added, skipped);
    AppendLog(msg);
    UpdateFileCountLabel();
}

static void ExtractDir(const wchar_t* path, wchar_t* dir, size_t n) {
    wcsncpy(dir, path, n - 1);
    dir[n - 1] = L'\0';
    wchar_t* slash = wcsrchr(dir, L'\\');
    if (slash) *slash = L'\0';
}

static void ScanFolder(const wchar_t* dir, int* added, int* skipped) {
    if (wcslen(dir) + 4 >= DC_PATH) return;

    wchar_t pattern[DC_PATH];
    swprintf(pattern, DC_PATH, L"%s\\*", dir);

    WIN32_FIND_DATAW fd;
    HANDLE h = FindFirstFileW(pattern, &fd);
    if (h == INVALID_HANDLE_VALUE) return;

    do {
        if (wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0) continue;
        if (wcslen(dir) + wcslen(fd.cFileName) + 2 >= DC_PATH) continue;

        wchar_t full[DC_PATH];
        swprintf(full, DC_PATH, L"%s\\%s", dir, fd.cFileName);

        if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
            ScanFolder(full, added, skipped);
        } else {
            AddFileToList(full, added, skipped);
        }
    } while (FindNextFileW(h, &fd));

    FindClose(h);
}

// ---------------- 浏览操作 ----------------
static void OnAddFiles(HWND hwnd) {
    wchar_t* buf = (wchar_t*)calloc(OPEN_BUF_LEN, sizeof(wchar_t));
    if (!buf) return;

    OPENFILENAMEW ofn;
    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hwnd;
    ofn.lpstrFilter = L"数据文件 (*.xlsx;*.csv)\0*.xlsx;*.csv\0Excel 文件 (*.xlsx)\0*.xlsx\0CSV 文件 (*.csv)\0*.csv\0所有文件 (*.*)\0*.*\0";
    ofn.lpstrFile = buf;
    ofn.nMaxFile = OPEN_BUF_LEN;
    ofn.Flags = OFN_ALLOWMULTISELECT | OFN_EXPLORER | OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;
    ofn.lpstrInitialDir = g_szLastDir[0] ? g_szLastDir : NULL;

    if (GetOpenFileNameW(&ofn)) {
        int added = 0, skipped = 0;
        wchar_t* p = buf + wcslen(buf) + 1;

        if (*p == L'\0') {
            // 单选
            AddFileToList(buf, &added, &skipped);
            ExtractDir(buf, g_szLastDir, DC_PATH);
        } else {
            // 多选：buf 为目录，其后为文件名列表
            wchar_t dir[DC_PATH];
            wcsncpy(dir, buf, DC_PATH - 1);
            dir[DC_PATH - 1] = L'\0';
            while (*p) {
                wchar_t full[DC_PATH];
                swprintf(full, DC_PATH, L"%s\\%s", dir, p);
                AddFileToList(full, &added, &skipped);
                p += wcslen(p) + 1;
            }
            wcsncpy(g_szLastDir, dir, DC_PATH - 1);
            g_szLastDir[DC_PATH - 1] = L'\0';
        }
        LogAddResult(L"添加文件", added, skipped);
    }
    free(buf);
}

static void OnAddFolder(HWND hwnd) {
    BROWSEINFOW bi;
    ZeroMemory(&bi, sizeof(bi));
    bi.hwndOwner = hwnd;
    bi.lpszTitle = L"选择要扫描的文件夹（包含子文件夹中的 .xlsx / .csv）：";
    bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;

    LPITEMIDLIST pidl = SHBrowseForFolderW(&bi);
    if (pidl) {
        wchar_t dir[DC_PATH];
        if (SHGetPathFromIDListW(pidl, dir)) {
            int added = 0, skipped = 0;
            ScanFolder(dir, &added, &skipped);
            wcsncpy(g_szLastDir, dir, DC_PATH - 1);
            g_szLastDir[DC_PATH - 1] = L'\0';
            LogAddResult(L"扫描文件夹", added, skipped);
        }
        CoTaskMemFree(pidl);
    }
}

static void OnRemoveSelected(void) {
    int selCount = (int)SendMessageW(g_hListBox, LB_GETSELCOUNT, 0, 0);
    if (selCount <= 0) return;

    int* items = (int*)malloc(sizeof(int) * selCount);
    if (!items) return;

    int got = (int)SendMessageW(g_hListBox, LB_GETSELITEMS, (WPARAM)selCount, (LPARAM)items);
    // 从高索引向低索引删除，避免索引位移
    for (int i = got - 1; i >= 0; i--) {
        SendMessageW(g_hListBox, LB_DELETESTRING, (WPARAM)items[i], 0);
    }
    free(items);

    wchar_t msg[128];
    swprintf(msg, 128, L"已从列表移除 %d 个文件。", got);
    AppendLog(msg);
    UpdateFileCountLabel();
}

static void OnClearList(void) {
    SendMessageW(g_hListBox, LB_RESETCONTENT, 0, 0);
    AppendLog(L"已清空文件列表。");
    UpdateFileCountLabel();
}

// ---------------- 模式选择联动 ----------------
static void OnModeChanged(void) {
    int sel = (int)SendMessageW(g_hComboMode, CB_GETCURSEL, 0, 0);
    if (sel < 0 || sel >= CLEAN_MODE_TOTAL) return;

    SetWindowTextW(g_hLblModeDesc, MODE_DESCRIPTIONS[sel]);

    BOOL isDefault = (sel == CLEAN_MODE_DEFAULT);
    BOOL isText = (sel == CLEAN_MODE_TEXT);
    if (!g_bIsProcessing) {
        EnableWindow(g_hEditSep, isDefault);
        EnableWindow(g_hLblSep, isDefault);
        EnableWindow(g_hChkSpecial, isText);
    }
}

// ---------------- 打开结果 ----------------
static void OnOpenFolder(void) {
    if (wcslen(g_szLastOutput) > 0) {
        wchar_t dir[DC_PATH];
        wcscpy(dir, g_szLastOutput);
        wchar_t* slash = wcsrchr(dir, L'\\');
        if (slash) *slash = L'\0';
        ShellExecuteW(NULL, L"open", dir, NULL, NULL, SW_SHOW);
    }
}

static void OnOpenFile(void) {
    if (wcslen(g_szLastOutput) > 0) {
        ShellExecuteW(NULL, L"open", g_szLastOutput, NULL, NULL, SW_SHOW);
    }
}

// ---------------- 进度回调（工作线程调用） ----------------
static void ProgressCallbackFunc(void* user, int percentage, const wchar_t* status_msg) {
    (void)user;
    SendMessageW(g_hProgressBar, PBM_SETPOS, percentage, 0);
    if (status_msg) AppendLog(status_msg);
}

// ---------------- 后台线程 ----------------
typedef struct {
    FileJob* jobs;
    int count;
    CleanOptions opts;
} BatchParams;

static void SetProcessingUI(BOOL processing) {
    g_bIsProcessing = processing;
    if (processing) g_cancel_requested = 0;

    EnableWindow(g_hBtnAddFiles, !processing);
    EnableWindow(g_hBtnAddFolder, !processing);
    EnableWindow(g_hBtnRemoveSel, !processing);
    EnableWindow(g_hBtnClearList, !processing);
    EnableWindow(g_hListBox, !processing);
    EnableWindow(g_hComboMode, !processing);
    EnableWindow(g_hLblSep, !processing && SendMessageW(g_hComboMode, CB_GETCURSEL, 0, 0) == CLEAN_MODE_DEFAULT);
    EnableWindow(g_hEditSep, !processing && SendMessageW(g_hComboMode, CB_GETCURSEL, 0, 0) == CLEAN_MODE_DEFAULT);
    EnableWindow(g_hChkSpecial, !processing && SendMessageW(g_hComboMode, CB_GETCURSEL, 0, 0) == CLEAN_MODE_TEXT);

    SetWindowTextW(g_hBtnStart, processing ? L"⏹ 取消任务" : L"🚀 开始清理数据");
    EnableWindow(g_hBtnStart, TRUE);

    if (processing) {
        EnableWindow(g_hBtnOpenFolder, FALSE);
        EnableWindow(g_hBtnOpenFile, FALSE);
    }
}

DWORD WINAPI WorkerThread(LPVOID lpParam) {
    BatchParams* bp = (BatchParams*)lpParam;

    AppendLog(L"=== 开始批量数据清洗任务 ===");
    wchar_t line[256];
    int modeSel = (int)bp->opts.mode;
    if (modeSel >= 0 && modeSel < CLEAN_MODE_TOTAL) {
        swprintf(line, 256, L"清洗模式：%s | 文件数：%d", MODE_NAMES[modeSel], bp->count);
        AppendLog(line);
    }

    CleanStats total;
    memset(&total, 0, sizeof(total));
    run_clean_jobs(bp->jobs, bp->count, &bp->opts, ProgressCallbackFunc, NULL, &total);

    // 汇总结果
    int ok = 0, fail = 0, cancel = 0, skip = 0;
    wchar_t lastOut[DC_PATH] = L"";
    for (int i = 0; i < bp->count; i++) {
        switch (bp->jobs[i].result) {
        case 0: ok++; wcscpy(lastOut, bp->jobs[i].output); break;
        case 1: fail++; break;
        case 2: cancel++; break;
        default: skip++; break;
        }
    }

    AppendLog(L"========================================");
    swprintf(line, 256, L"🎉 任务结束！成功 %d 个，失败 %d 个，取消 %d 个，跳过 %d 个", ok, fail, cancel, skip);
    AppendLog(line);

    swprintf(line, 256, L"累计统计：处理单元格 %lld | 修改 %lld | 保留 %lld",
             total.total_cells, total.modified_cells, total.kept_cells);
    AppendLog(line);
    swprintf(line, 256, L"累计统计：提取数字 %lld 个 | 标准化日期 %lld 个",
             total.numbers_found, total.dates_normalized);
    AppendLog(line);

    if (lastOut[0]) {
        wcscpy(g_szLastOutput, lastOut);
        wchar_t msg[DC_PATH + 64];
        swprintf(msg, DC_PATH + 64, L"最后一个输出文件：%s", lastOut);
        AppendLog(msg);
    } else {
        g_szLastOutput[0] = L'\0';
    }

    // 恢复 UI 状态
    g_cancel_requested = 0;
    SetProcessingUI(FALSE);
    EnableWindow(g_hBtnOpenFolder, lastOut[0] ? TRUE : FALSE);
    EnableWindow(g_hBtnOpenFile, lastOut[0] ? TRUE : FALSE);
    SendMessageW(g_hProgressBar, PBM_SETPOS, ok + fail + cancel + skip == bp->count && bp->count > 0 ? 100 : 0, 0);

    // 结果弹窗
    wchar_t summary[512];
    if (fail == 0 && cancel == 0 && skip == 0) {
        swprintf(summary, 512,
                 L"数据清洗完成！\n\n成功处理 %d 个文件\n修改单元格 %lld 个（共 %lld 个）\n提取数字 %lld 个，标准化日期 %lld 个",
                 ok, total.modified_cells, total.total_cells, total.numbers_found, total.dates_normalized);
        MessageBoxW(NULL, summary, L"成功", MB_OK | MB_ICONINFORMATION);
    } else {
        swprintf(summary, 512,
                 L"任务结束。\n\n成功：%d，失败：%d，取消：%d，跳过：%d\n修改单元格 %lld 个（共 %lld 个）\n详情请查看主界面日志。",
                 ok, fail, cancel, skip, total.modified_cells, total.total_cells);
        MessageBoxW(NULL, summary, L"完成（有异常）", MB_OK | MB_ICONWARNING);
    }

    free(bp->jobs);
    free(bp);
    return 0;
}

// ---------------- 开始 / 取消 ----------------
static void OnStartCleaning(HWND hwnd) {
    if (g_bIsProcessing) {
        g_cancel_requested = 1;
        AppendLog(L"⏹ 已请求取消，正在停止当前任务...");
        return;
    }

    int fileCount = (int)SendMessageW(g_hListBox, LB_GETCOUNT, 0, 0);
    if (fileCount <= 0) {
        MessageBoxW(hwnd, L"请先添加要清洗的 Excel / CSV 文件！", L"提示", MB_OK | MB_ICONWARNING);
        return;
    }

    // 读取清洗选项
    CleanOptions opts;
    int sel = (int)SendMessageW(g_hComboMode, CB_GETCURSEL, 0, 0);
    opts.mode = (sel >= 0 && sel < CLEAN_MODE_TOTAL) ? (CleanMode)sel : CLEAN_MODE_DEFAULT;

    GetWindowTextW(g_hEditSep, opts.separator, 16);
    // 仅空格视为未设置
    wchar_t* sp = opts.separator;
    while (*sp == L' ') sp++;
    if (*sp == L'\0') wcscpy(opts.separator, L", ");

    opts.aggressive_special_removal =
        (SendMessageW(g_hChkSpecial, BM_GETCHECK, 0, 0) == BST_CHECKED) ? 1 : 0;

    // 构建任务列表
    FileJob* jobs = (FileJob*)calloc(fileCount, sizeof(FileJob));
    if (!jobs) {
        MessageBoxW(hwnd, L"内存分配失败！", L"错误", MB_OK | MB_ICONERROR);
        return;
    }
    for (int i = 0; i < fileCount; i++) {
        SendMessageW(g_hListBox, LB_GETTEXT, (WPARAM)i, (LPARAM)jobs[i].input);
    }

    BatchParams* bp = (BatchParams*)malloc(sizeof(BatchParams));
    if (!bp) {
        free(jobs);
        MessageBoxW(hwnd, L"内存分配失败！", L"错误", MB_OK | MB_ICONERROR);
        return;
    }
    bp->jobs = jobs;
    bp->count = fileCount;
    bp->opts = opts;

    // 初始化 UI
    SetProcessingUI(TRUE);
    SendMessageW(g_hProgressBar, PBM_SETPOS, 0, 0);
    SetWindowTextW(g_hEditLog, L"");

    HANDLE hThread = CreateThread(NULL, 0, WorkerThread, bp, 0, NULL);
    if (hThread) {
        CloseHandle(hThread);
    } else {
        free(bp->jobs);
        free(bp);
        SetProcessingUI(FALSE);
        MessageBoxW(hwnd, L"无法创建工作线程！", L"错误", MB_OK | MB_ICONERROR);
    }
}

// ---------------- 拖放支持 ----------------
static void OnDropFiles(HDROP hDrop) {
    UINT count = DragQueryFileW(hDrop, 0xFFFFFFFF, NULL, 0);
    int added = 0, skipped = 0;

    for (UINT i = 0; i < count; i++) {
        wchar_t path[DC_PATH];
        if (DragQueryFileW(hDrop, i, path, DC_PATH) > 0) {
            AddFileToList(path, &added, &skipped);
            if (i == 0 && added == 1) ExtractDir(path, g_szLastDir, DC_PATH);
        }
    }
    DragFinish(hDrop);
    LogAddResult(L"拖放添加", added, skipped);
}

// ---------------- 注册表配置 ----------------
static void SaveConfig(void) {
    HKEY k;
    if (RegCreateKeyExW(HKEY_CURRENT_USER, REG_KEY, 0, NULL, 0, KEY_WRITE, NULL, &k, NULL) != ERROR_SUCCESS)
        return;

    DWORD mode = (DWORD)SendMessageW(g_hComboMode, CB_GETCURSEL, 0, 0);
    DWORD special = (DWORD)(SendMessageW(g_hChkSpecial, BM_GETCHECK, 0, 0) == BST_CHECKED);

    wchar_t sep[32];
    GetWindowTextW(g_hEditSep, sep, 16);

    RegSetValueExW(k, L"Mode", 0, REG_DWORD, (const BYTE*)&mode, sizeof(DWORD));
    RegSetValueExW(k, L"SpecialRemoval", 0, REG_DWORD, (const BYTE*)&special, sizeof(DWORD));
    RegSetValueExW(k, L"Separator", 0, REG_SZ, (const BYTE*)sep, (DWORD)((wcslen(sep) + 1) * sizeof(wchar_t)));
    if (g_szLastDir[0]) {
        RegSetValueExW(k, L"LastDir", 0, REG_SZ, (const BYTE*)g_szLastDir,
                       (DWORD)((wcslen(g_szLastDir) + 1) * sizeof(wchar_t)));
    }

    // 保存文件列表（上限 50 个）
    int cnt = (int)SendMessageW(g_hListBox, LB_GETCOUNT, 0, 0);
    if (cnt > 50) cnt = 50;
    DWORD dwCnt = (DWORD)cnt;
    RegSetValueExW(k, L"FileCount", 0, REG_DWORD, (const BYTE*)&dwCnt, sizeof(DWORD));
    for (int i = 0; i < cnt; i++) {
        wchar_t path[DC_PATH];
        SendMessageW(g_hListBox, LB_GETTEXT, (WPARAM)i, (LPARAM)path);
        wchar_t name[32];
        swprintf(name, 32, L"File%d", i);
        RegSetValueExW(k, name, 0, REG_SZ, (const BYTE*)path,
                       (DWORD)((wcslen(path) + 1) * sizeof(wchar_t)));
    }

    RegCloseKey(k);
}

static void LoadConfig(void) {
    HKEY k;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, REG_KEY, 0, KEY_READ, &k) != ERROR_SUCCESS)
        return;

    DWORD type = 0, size = 0;

    DWORD mode = 0;
    size = sizeof(DWORD);
    if (RegQueryValueExW(k, L"Mode", NULL, &type, (BYTE*)&mode, &size) == ERROR_SUCCESS
        && type == REG_DWORD && mode < CLEAN_MODE_TOTAL) {
        SendMessageW(g_hComboMode, CB_SETCURSEL, (WPARAM)mode, 0);
    } else {
        SendMessageW(g_hComboMode, CB_SETCURSEL, 0, 0);
    }

    DWORD special = 0;
    size = sizeof(DWORD);
    if (RegQueryValueExW(k, L"SpecialRemoval", NULL, &type, (BYTE*)&special, &size) == ERROR_SUCCESS
        && type == REG_DWORD && special) {
        SendMessageW(g_hChkSpecial, BM_SETCHECK, BST_CHECKED, 0);
    }

    wchar_t sep[32];
    size = sizeof(sep);
    if (RegQueryValueExW(k, L"Separator", NULL, &type, (BYTE*)sep, &size) == ERROR_SUCCESS
        && type == REG_SZ && sep[0]) {
        SetWindowTextW(g_hEditSep, sep);
    }

    size = sizeof(g_szLastDir);
    RegQueryValueExW(k, L"LastDir", NULL, &type, (BYTE*)g_szLastDir, &size);

    // 恢复文件列表（仅恢复仍存在的文件）
    DWORD fileCnt = 0;
    size = sizeof(DWORD);
    if (RegQueryValueExW(k, L"FileCount", NULL, &type, (BYTE*)&fileCnt, &size) == ERROR_SUCCESS
        && type == REG_DWORD && fileCnt > 0 && fileCnt <= 50) {
        int added = 0, skipped = 0;
        for (DWORD i = 0; i < fileCnt; i++) {
            wchar_t name[32], path[DC_PATH];
            swprintf(name, 32, L"File%u", i);
            size = sizeof(path);
            if (RegQueryValueExW(k, name, NULL, &type, (BYTE*)path, &size) == ERROR_SUCCESS
                && type == REG_SZ && path[0]
                && GetFileAttributesW(path) != INVALID_FILE_ATTRIBUTES) {
                AddFileToList(path, &added, &skipped);
            }
        }
        if (added > 0) {
            wchar_t msg[128];
            swprintf(msg, 128, L"已恢复上次会话的 %d 个文件。", added);
            AppendLog(msg);
        }
    }

    RegCloseKey(k);
}

// ---------------- 自适应布局 ----------------
static void LayoutControls(HWND hwnd) {
    RECT rc;
    GetClientRect(hwnd, &rc);
    int cw = rc.right;
    int ch = rc.bottom;
    if (cw < 200 || ch < 200) return;

    int gw = cw - 40;      // 分组宽度
    int x = 20;

    // 标题
    MoveWindow(g_hLblTitle, x, 12, gw, 30, TRUE);

    // 分组 1：文件列表
    MoveWindow(g_hGroupFiles, x, 48, gw, 190, TRUE);
    int listW = gw - 170;
    if (listW < 200) listW = 200;
    MoveWindow(g_hListBox, x + 15, 73, listW, 125, TRUE);
    int bx = x + 15 + listW + 12;
    MoveWindow(g_hBtnAddFiles, bx, 73, 118, 26, TRUE);
    MoveWindow(g_hBtnAddFolder, bx, 103, 118, 26, TRUE);
    MoveWindow(g_hBtnRemoveSel, bx, 133, 118, 26, TRUE);
    MoveWindow(g_hBtnClearList, bx, 163, 118, 26, TRUE);
    MoveWindow(g_hLblFileCount, x + 15, 203, listW, 20, TRUE);

    // 分组 2：清洗模式
    MoveWindow(g_hGroupMode, x, 248, gw, 122, TRUE);
    MoveWindow(g_hLblModeLabel, x + 15, 272, 60, 20, TRUE);
    MoveWindow(g_hComboMode, x + 75, 268, 230, 200, TRUE);
    MoveWindow(g_hLblSep, x + 320, 272, 60, 20, TRUE);
    MoveWindow(g_hEditSep, x + 380, 268, 90, 24, TRUE);
    MoveWindow(g_hChkSpecial, x + 485, 268, gw - 470, 22, TRUE);
    MoveWindow(g_hLblModeDesc, x + 15, 305, gw - 30, 52, TRUE);

    // 操作按钮
    MoveWindow(g_hBtnStart, x, 380, 190, 36, TRUE);
    MoveWindow(g_hBtnOpenFolder, x + 205, 380, 200, 36, TRUE);
    MoveWindow(g_hBtnOpenFile, x + 420, 380, 210, 36, TRUE);

    // 进度条与日志
    MoveWindow(g_hProgressBar, x, 428, gw, 18, TRUE);
    int logH = ch - 454 - 15;
    if (logH < 60) logH = 60;
    MoveWindow(g_hEditLog, x, 454, gw, logH, TRUE);
}

// ---------------- 窗口过程 ----------------
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE: {
        // 字体
        g_hFontTitle = CreateFontW(-18, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                                   OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                                   DEFAULT_PITCH | FF_DONTCARE, L"Microsoft YaHei");
        g_hFontNormal = CreateFontW(-14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                                    OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                                    DEFAULT_PITCH | FF_DONTCARE, L"Microsoft YaHei");
        g_hFontBold = CreateFontW(-14, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                                  OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                                  DEFAULT_PITCH | FF_DONTCARE, L"Microsoft YaHei");

        // 标题
        g_hLblTitle = CreateWindowW(L"STATIC", L"📊 数据清洗整理工具 v2.0",
                                    WS_VISIBLE | WS_CHILD,
                                    20, 12, 600, 30, hwnd, NULL, NULL, NULL);
        SendMessageW(g_hLblTitle, WM_SETFONT, (WPARAM)g_hFontTitle, TRUE);

        // 分组 1：文件列表
        g_hGroupFiles = CreateWindowW(L"BUTTON", L" 1. 数据源文件（支持批量） ",
                                      WS_VISIBLE | WS_CHILD | BS_GROUPBOX,
                                      20, 48, 680, 190, hwnd, NULL, NULL, NULL);
        SendMessageW(g_hGroupFiles, WM_SETFONT, (WPARAM)g_hFontBold, TRUE);

        g_hListBox = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", L"",
                                     WS_VISIBLE | WS_CHILD | WS_VSCROLL | LBS_EXTENDEDSEL | LBS_NOTIFY | LBS_NOINTEGRALHEIGHT,
                                     35, 73, 500, 125, hwnd, (HMENU)IDC_LB_FILES, NULL, NULL);
        SendMessageW(g_hListBox, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);

        g_hBtnAddFiles = CreateWindowW(L"BUTTON", L"＋ 添加文件...",
                                       WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
                                       547, 73, 118, 26, hwnd, (HMENU)IDC_BTN_ADD_FILES, NULL, NULL);
        SendMessageW(g_hBtnAddFiles, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);

        g_hBtnAddFolder = CreateWindowW(L"BUTTON", L"📁 添加文件夹...",
                                        WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
                                        547, 103, 118, 26, hwnd, (HMENU)IDC_BTN_ADD_FOLDER, NULL, NULL);
        SendMessageW(g_hBtnAddFolder, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);

        g_hBtnRemoveSel = CreateWindowW(L"BUTTON", L"✖ 移除选中",
                                        WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
                                        547, 133, 118, 26, hwnd, (HMENU)IDC_BTN_REMOVE_SEL, NULL, NULL);
        SendMessageW(g_hBtnRemoveSel, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);

        g_hBtnClearList = CreateWindowW(L"BUTTON", L"🗑 清空列表",
                                        WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
                                        547, 163, 118, 26, hwnd, (HMENU)IDC_BTN_CLEAR_LIST, NULL, NULL);
        SendMessageW(g_hBtnClearList, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);

        g_hLblFileCount = CreateWindowW(L"STATIC", L"待处理文件：0 个",
                                        WS_VISIBLE | WS_CHILD | SS_NOPREFIX,
                                        35, 203, 500, 20, hwnd, NULL, NULL, NULL);
        SendMessageW(g_hLblFileCount, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);

        // 分组 2：清洗模式
        g_hGroupMode = CreateWindowW(L"BUTTON", L" 2. 清洗模式与选项 ",
                                     WS_VISIBLE | WS_CHILD | BS_GROUPBOX,
                                     20, 248, 680, 122, hwnd, NULL, NULL, NULL);
        SendMessageW(g_hGroupMode, WM_SETFONT, (WPARAM)g_hFontBold, TRUE);

        g_hLblModeLabel = CreateWindowW(L"STATIC", L"清洗模式：",
                                        WS_VISIBLE | WS_CHILD,
                                        35, 272, 60, 20, hwnd, NULL, NULL, NULL);
        SendMessageW(g_hLblModeLabel, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);

        g_hComboMode = CreateWindowW(L"COMBOBOX", NULL,
                                     WS_VISIBLE | WS_CHILD | CBS_DROPDOWNLIST | WS_VSCROLL,
                                     95, 268, 230, 200, hwnd, (HMENU)IDC_COMBO_MODE, NULL, NULL);
        SendMessageW(g_hComboMode, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);
        for (int i = 0; i < CLEAN_MODE_TOTAL; i++) {
            SendMessageW(g_hComboMode, CB_ADDSTRING, 0, (LPARAM)MODE_NAMES[i]);
        }
        SendMessageW(g_hComboMode, CB_SETCURSEL, 0, 0);

        g_hLblSep = CreateWindowW(L"STATIC", L"分隔符：",
                                  WS_VISIBLE | WS_CHILD,
                                  340, 272, 60, 20, hwnd, NULL, NULL, NULL);
        SendMessageW(g_hLblSep, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);

        g_hEditSep = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L", ",
                                     WS_VISIBLE | WS_CHILD | ES_AUTOHSCROLL,
                                     400, 268, 90, 24, hwnd, (HMENU)IDC_EDIT_SEP, NULL, NULL);
        SendMessageW(g_hEditSep, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);
        SendMessageW(g_hEditSep, EM_LIMITTEXT, 15, 0);

        g_hChkSpecial = CreateWindowW(L"BUTTON", L"文本清理时仅保留中文/字母/数字/空格",
                                      WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX | WS_DISABLED,
                                      505, 268, 195, 22, hwnd, (HMENU)IDC_CHK_SPECIAL, NULL, NULL);
        SendMessageW(g_hChkSpecial, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);

        g_hLblModeDesc = CreateWindowW(L"STATIC", MODE_DESCRIPTIONS[0],
                                       WS_VISIBLE | WS_CHILD | SS_NOPREFIX,
                                       35, 305, 650, 52, hwnd, NULL, NULL, NULL);
        SendMessageW(g_hLblModeDesc, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);

        // 操作按钮
        g_hBtnStart = CreateWindowW(L"BUTTON", L"🚀 开始清理数据",
                                    WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
                                    20, 380, 190, 36, hwnd, (HMENU)IDC_BTN_START, NULL, NULL);
        SendMessageW(g_hBtnStart, WM_SETFONT, (WPARAM)g_hFontBold, TRUE);

        g_hBtnOpenFolder = CreateWindowW(L"BUTTON", L"📁 打开输出文件夹",
                                         WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON | WS_DISABLED,
                                         225, 380, 200, 36, hwnd, (HMENU)IDC_BTN_OPEN_FOLDER, NULL, NULL);
        SendMessageW(g_hBtnOpenFolder, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);

        g_hBtnOpenFile = CreateWindowW(L"BUTTON", L"📈 查看清洗结果",
                                       WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON | WS_DISABLED,
                                       440, 380, 210, 36, hwnd, (HMENU)IDC_BTN_OPEN_FILE, NULL, NULL);
        SendMessageW(g_hBtnOpenFile, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);

        // 进度条与日志
        g_hProgressBar = CreateWindowExW(0, PROGRESS_CLASSW, NULL,
                                         WS_VISIBLE | WS_CHILD,
                                         20, 428, 680, 18, hwnd, (HMENU)IDC_PROGRESS, NULL, NULL);
        SendMessageW(g_hProgressBar, PBM_SETRANGE, 0, MAKELPARAM(0, 100));

        g_hEditLog = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                                     WS_VISIBLE | WS_CHILD | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY | WS_VSCROLL,
                                     20, 454, 680, 130, hwnd, (HMENU)IDC_EDIT_LOG, NULL, NULL);
        SendMessageW(g_hEditLog, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);

        // 拖放支持
        DragAcceptFiles(hwnd, TRUE);

        // 恢复上次配置
        LoadConfig();
        OnModeChanged();
        UpdateFileCountLabel();

        AppendLog(L"[程序已就绪] 支持 .xlsx / .csv 批量清洗，可将文件直接拖放到窗口。");
        LayoutControls(hwnd);
        break;
    }
    case WM_SIZE:
        LayoutControls(hwnd);
        break;
    case WM_GETMINMAXINFO: {
        MINMAXINFO* mmi = (MINMAXINFO*)lParam;
        mmi->ptMinTrackSize.x = 760;
        mmi->ptMinTrackSize.y = 620;
        return 0;
    }
    case WM_DROPFILES:
        OnDropFiles((HDROP)wParam);
        return 0;
    case WM_COMMAND: {
        int wID = LOWORD(wParam);
        int notif = HIWORD(wParam);
        if (notif == CBN_SELCHANGE && wID == IDC_COMBO_MODE) {
            OnModeChanged();
        } else if (notif == LBN_DBLCLK && wID == IDC_LB_FILES) {
            OnRemoveSelected();
        } else if (wID == IDC_BTN_ADD_FILES) {
            OnAddFiles(hwnd);
        } else if (wID == IDC_BTN_ADD_FOLDER) {
            OnAddFolder(hwnd);
        } else if (wID == IDC_BTN_REMOVE_SEL) {
            OnRemoveSelected();
        } else if (wID == IDC_BTN_CLEAR_LIST) {
            OnClearList();
        } else if (wID == IDC_BTN_START) {
            OnStartCleaning(hwnd);
        } else if (wID == IDC_BTN_OPEN_FOLDER) {
            OnOpenFolder();
        } else if (wID == IDC_BTN_OPEN_FILE) {
            OnOpenFile();
        }
        break;
    }
    case WM_DESTROY:
        SaveConfig();
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
    (void)hPrevInstance;
    (void)pCmdLine;

    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_PROGRESS_CLASS | ICC_STANDARD_CLASSES | ICC_LISTVIEW_CLASSES;
    InitCommonControlsEx(&icex);

    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

    const wchar_t CLASS_NAME[] = L"DataCleanerWindowClass";

    WNDCLASSW wc = {0};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = CLASS_NAME;
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hIcon = LoadIcon(NULL, IDI_APPLICATION);

    RegisterClassW(&wc);

    HWND hwnd = CreateWindowExW(
        0, CLASS_NAME, L"数据清洗整理工具 v2.0",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX | WS_MAXIMIZEBOX | WS_THICKFRAME,
        CW_USEDEFAULT, CW_USEDEFAULT, 900, 660,
        NULL, NULL, hInstance, NULL
    );

    if (hwnd == NULL) {
        CoUninitialize();
        return 0;
    }

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    MSG msg = {0};
    while (GetMessageW(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    CoUninitialize();
    return 0;
}
