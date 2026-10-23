#ifndef EXCEL_CLEANER_H
#define EXCEL_CLEANER_H

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <wchar.h>

// 统一路径缓冲长度（支持较深的文件夹递归）
#define DC_PATH 1024

#ifdef __cplusplus
extern "C" {
#endif

// 清洗模式
typedef enum {
    CLEAN_MODE_DEFAULT  = 0, // 数字提取，以分隔符拼接输出
    CLEAN_MODE_ADD      = 1, // 数字求和
    CLEAN_MODE_MULTIPLY = 2, // 数字求积
    CLEAN_MODE_AVERAGE  = 3, // 数字平均值
    CLEAN_MODE_MAX      = 4, // 数字最大值
    CLEAN_MODE_MIN      = 5, // 数字最小值
    CLEAN_MODE_COUNT    = 6, // 数字个数统计
    CLEAN_MODE_TEXT     = 7, // 文本清理
    CLEAN_MODE_DATE     = 8, // 日期标准化
    CLEAN_MODE_TOTAL    = 9  // 模式总数（内部使用）
} CleanMode;

// 清洗选项（由 UI 层填充）
typedef struct {
    CleanMode mode;
    wchar_t separator[16];           // CLEAN_MODE_DEFAULT 的拼接分隔符
    int aggressive_special_removal;  // CLEAN_MODE_TEXT：仅保留中文/字母/数字/空格
} CleanOptions;

// 清洗统计（单个文件或累计）
typedef struct {
    long long total_cells;       // 处理的非空单元格数
    long long modified_cells;    // 内容发生变化的单元格数
    long long kept_cells;        // 保持原样的单元格数
    long long numbers_found;     // 提取到的数字个数
    long long dates_normalized;  // 标准化的日期个数
} CleanStats;

// 进度回调：user 为透传上下文
typedef void (*ProgressCallback)(void* user, int percentage, const wchar_t* status_msg);

// 单个文件任务（输入由调用方填写，输出由引擎填写）
typedef struct {
    wchar_t input[DC_PATH];
    wchar_t output[DC_PATH];
    int result;        // 0 成功 / 1 失败 / 2 已取消 / 3 已跳过
    CleanStats stats;
} FileJob;

// 取消标志：UI 线程置 1，引擎在文件/工作表/行边界检查
extern volatile long g_cancel_requested;

// 批量执行清洗任务；逐文件回调进度，累计统计写入 total_stats（可为 NULL）
// 返回 0 表示整体流程结束（具体成败见各 FileJob.result）
int run_clean_jobs(FileJob* jobs, int job_count, const CleanOptions* opts,
                   ProgressCallback cb, void* user, CleanStats* total_stats);

#ifdef __cplusplus
}
#endif

#endif // EXCEL_CLEANER_H
