#include "excel_cleaner.h"
#include <xlsxio_read.h>
#include <xlsxwriter.h>
#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <math.h>

#define MAX_NUMS_PER_CELL 128
#define MAX_CELL_STR_LEN  4096
#define MAX_SHEETS        64

// 取消标志定义（UI 线程设置）
volatile long g_cancel_requested = 0;

// ============================================================
// 编码转换工具
// ============================================================

static char* wchar_to_utf8(const wchar_t* wstr) {
    if (!wstr) return NULL;
    int len = WideCharToMultiByte(CP_UTF8, 0, wstr, -1, NULL, 0, NULL, NULL);
    if (len <= 0) return NULL;
    char* str = (char*)malloc(len);
    if (!str) return NULL;
    WideCharToMultiByte(CP_UTF8, 0, wstr, -1, str, len, NULL, NULL);
    return str;
}

static wchar_t* utf8_to_wchar(const char* str) {
    if (!str) return NULL;
    int len = MultiByteToWideChar(CP_UTF8, 0, str, -1, NULL, 0);
    if (len <= 0) return NULL;
    wchar_t* wstr = (wchar_t*)malloc((size_t)len * sizeof(wchar_t));
    if (!wstr) return NULL;
    MultiByteToWideChar(CP_UTF8, 0, str, -1, wstr, len);
    return wstr;
}

static char* wchar_to_utf8_len(const wchar_t* wstr, int wlen) {
    if (!wstr || wlen < 0) return NULL;
    int len = WideCharToMultiByte(CP_UTF8, 0, wstr, wlen, NULL, 0, NULL, NULL);
    if (len <= 0) return NULL;
    char* str = (char*)malloc(len + 1);
    if (!str) return NULL;
    WideCharToMultiByte(CP_UTF8, 0, wstr, wlen, str, len, NULL, NULL);
    str[len] = '\0';
    return str;
}

// ============================================================
// 数字提取与格式化
// ============================================================

// 读取 p 处的 UTF-8 字符，返回码点并通过 adv 返回字节数；非法字节按单字节原样处理
static unsigned int utf8_decode_at(const char* p, int* adv) {
    const unsigned char* s = (const unsigned char*)p;
    if (s[0] < 0x80) { *adv = 1; return s[0]; }
    if ((s[0] & 0xE0) == 0xC0 && (s[1] & 0xC0) == 0x80) {
        *adv = 2;
        return ((unsigned)(s[0] & 0x1F) << 6) | (unsigned)(s[1] & 0x3F);
    }
    if ((s[0] & 0xF0) == 0xE0 && (s[1] & 0xC0) == 0x80 && (s[2] & 0xC0) == 0x80) {
        *adv = 3;
        return ((unsigned)(s[0] & 0x0F) << 12) | ((unsigned)(s[1] & 0x3F) << 6) | (unsigned)(s[2] & 0x3F);
    }
    if ((s[0] & 0xF8) == 0xF0 && (s[1] & 0xC0) == 0x80 && (s[2] & 0xC0) == 0x80 && (s[3] & 0xC0) == 0x80) {
        *adv = 4;
        return ((unsigned)(s[0] & 0x07) << 18) | ((unsigned)(s[1] & 0x3F) << 12) |
               ((unsigned)(s[2] & 0x3F) << 6) | (unsigned)(s[3] & 0x3F);
    }
    *adv = 1;
    return s[0];
}

// 数字识别前的文本规范化（返回 malloc 缓冲，调用方 free；失败返回 NULL）：
// 1) 全角数字/小数点/逗号/负号转半角（０-９ ． ， － → 0-9 . , -），其余字符原样保留
// 2) 去除千分位分隔符：逗号前一位为数字、其后恰好 3 位数字且再后一位不是数字时，
//    视为分组逗号并删除（"1,234,567" → "1234567"；"3, 5"、"3,50" 中的逗号不合并）
static char* normalize_number_text(const char* text) {
    size_t n = strlen(text);
    char* a = (char*)malloc(n + 1);
    if (!a) return NULL;

    size_t w = 0;
    for (size_t i = 0; i < n; ) {
        int adv = 1;
        unsigned int cp = utf8_decode_at(text + i, &adv);
        if (cp >= 0xFF10 && cp <= 0xFF19) a[w++] = (char)('0' + (int)(cp - 0xFF10));
        else if (cp == 0xFF0E) a[w++] = '.';
        else if (cp == 0xFF0C) a[w++] = ',';
        else if (cp == 0xFF0D) a[w++] = '-';
        else { memcpy(a + w, text + i, (size_t)adv); w += (size_t)adv; }
        i += (size_t)adv;
    }
    a[w] = '\0';

    char* b = (char*)malloc(w + 1);
    if (!b) { free(a); return NULL; }
    size_t m = 0;
    for (size_t i = 0; i < w; i++) {
        if (a[i] == ',' && m > 0 && isdigit((unsigned char)b[m - 1])
            && i + 3 < w
            && isdigit((unsigned char)a[i + 1]) && isdigit((unsigned char)a[i + 2]) && isdigit((unsigned char)a[i + 3])
            && (i + 4 >= w || !isdigit((unsigned char)a[i + 4]))) {
            continue; // 千分位分组逗号，跳过
        }
        b[m++] = a[i];
    }
    b[m] = '\0';
    free(a);
    return b;
}

// 提取数字；raws 非 NULL 时同时记录每个数字在规范化文本中的原始字面（供长整数原样输出）
static int extract_numbers(const char* text, double* nums, char (*raws)[64], int max_nums) {
    if (!text || !*text) return 0;
    char* norm = normalize_number_text(text);
    const char* src = norm ? norm : text;
    int count = 0;
    const char* p = src;

    while (*p && count < max_nums) {
        const char* start = NULL;
        if (*p == '-' && (isdigit((unsigned char)*(p + 1)) || (*(p + 1) == '.' && isdigit((unsigned char)*(p + 2))))) {
            start = p;
        } else if (isdigit((unsigned char)*p) || (*p == '.' && isdigit((unsigned char)*(p + 1)))) {
            start = p;
        }
        if (start) {
            char* endptr = NULL;
            double val = strtod(start, &endptr);
            if (endptr > start) {
                nums[count] = val;
                if (raws) {
                    size_t tl = (size_t)(endptr - start);
                    if (tl > 63) tl = 63;
                    memcpy(raws[count], start, tl);
                    raws[count][tl] = '\0';
                }
                count++;
                p = endptr;
                continue;
            }
        }
        p++;
    }
    free(norm);
    return count;
}

// 判断字面是否为纯整数（可选负号 + 数字）且位数 > 15（超出 double 可长期精确表示的范围）
static int is_long_integer_token(const char* s) {
    if (!s || !*s) return 0;
    if (*s == '-') s++;
    int digits = 0;
    for (; *s; s++) {
        if (!isdigit((unsigned char)*s)) return 0;
        digits++;
    }
    return digits > 15;
}

static void format_double_clean(double val, char* buf, size_t buf_size) {
    // 2^53 以内的整数可被 double 精确表示，直接输出整数形式，
    // 避免 %.6g 把长整数写成科学计数法而丢失精度
    if (floor(val) == val && fabs(val) < 9007199254740992.0) {
        snprintf(buf, buf_size, "%.0f", val);
    } else {
        snprintf(buf, buf_size, "%.6g", val);
    }
}

// ============================================================
// 文本清理模式
// ============================================================

// 判断字符是否应保留（激进模式：仅中文/字母/数字/空格）
static int text_keep_char(wchar_t c, int aggressive) {
    if (aggressive) {
        if (c == 0x20) return 1;
        if (c >= L'0' && c <= L'9') return 1;
        if (c >= L'A' && c <= L'Z') return 1;
        if (c >= L'a' && c <= L'z') return 1;
        if (c >= 0x3400 && c <= 0x9FFF) return 1;   // CJK 扩展A + 基本区
        if (c >= 0xF900 && c <= 0xFAFF) return 1;   // CJK 兼容表意
        return 0;
    }
    // 普通模式：仅剔除控制字符（含 \t \r \n）
    return (c >= 0x20) ? 1 : 0;
}

static void text_clean_utf8(const char* in, int aggressive, char* out, size_t out_max) {
    out[0] = '\0';
    if (!in || !*in) return;

    wchar_t* w = utf8_to_wchar(in);
    if (!w) {
        strncpy(out, in, out_max - 1);
        out[out_max - 1] = '\0';
        return;
    }

    int n = (int)wcslen(w);
    wchar_t* tmp = (wchar_t*)malloc(((size_t)n + 1) * sizeof(wchar_t));
    if (!tmp) { free(w); strncpy(out, in, out_max - 1); out[out_max - 1] = '\0'; return; }

    // 1. 全角转半角 + 2. 过滤字符
    int k = 0;
    for (int i = 0; i < n; i++) {
        wchar_t c = w[i];
        if (c == 0x3000) c = 0x20;                       // 全角空格
        else if (c >= 0xFF01 && c <= 0xFF5E) c -= 0xFEE0; // 全角 ASCII
        if (!text_keep_char(c, aggressive)) continue;
        tmp[k++] = c;
    }
    tmp[k] = L'\0';

    // 3. 合并连续空格 + 4. 去首尾空格
    int s = 0, e = k - 1;
    while (s <= e && tmp[s] == 0x20) s++;
    while (e >= s && tmp[e] == 0x20) e--;

    if (e < s) {
        free(tmp);
        free(w);
        return; // 清理后为空
    }

    int m = 0;
    for (int i = s; i <= e; i++) {
        if (tmp[i] == 0x20 && m > 0 && tmp[m - 1] == 0x20) continue;
        tmp[m++] = tmp[i];
    }
    tmp[m] = L'\0';

    char* u = wchar_to_utf8_len(tmp, m);
    if (u) {
        strncpy(out, u, out_max - 1);
        out[out_max - 1] = '\0';
        free(u);
    }
    free(tmp);
    free(w);
}

// ============================================================
// 日期标准化模式
// ============================================================

static int is_leap_year(int y) {
    return (y % 4 == 0 && y % 100 != 0) || (y % 400 == 0);
}

static int days_in_month(int y, int m) {
    static const int d[] = { 31,28,31,30,31,30,31,31,30,31,30,31 };
    if (m < 1 || m > 12) return 0;
    if (m == 2 && is_leap_year(y)) return 29;
    return d[m - 1];
}

// 从 *pp 读取最多 max_digits 位数字，返回实际读取位数
static int read_digits(const wchar_t** pp, int max_digits, int* val) {
    const wchar_t* p = *pp;
    int v = 0, n = 0;
    while (n < max_digits && *p >= L'0' && *p <= L'9') {
        v = v * 10 + (int)(*p - L'0');
        p++;
        n++;
    }
    *pp = p;
    *val = v;
    return n;
}

static int is_date_sep(wchar_t c) {
    return c == L'-' || c == L'.' || c == L'/' || c == L'\x5E74'; // 年
}

// 尝试将输入识别为日期（可带时间）并标准化为 YYYY-MM-DD[ HH:MM[:SS]]
// 返回 1 表示已标准化（out 为结果）；返回 0 表示无法识别（out 复制原文）
static int date_normalize_utf8(const char* in, char* out, size_t out_max) {
    strncpy(out, in, out_max - 1);
    out[out_max - 1] = '\0';
    if (!in || !*in) return 0;

    wchar_t* w = utf8_to_wchar(in);
    if (!w) return 0;

    const wchar_t* p = w;
    while (*p == L' ') p++;

    int y = 0, mo = 0, d = 0, hh = -1, mi = -1, ss = -1;
    int matched = 0;

    // 形态一：YYYY?M?D[日][ H:M[:S]]，分隔符为 - . / 年 月
    // 形态二：YYYYMMDD（恰好 8 位数字）
    long long v = 0;
    int nd = 0;
    const wchar_t* q = p;
    while (nd < 8 && *q >= L'0' && *q <= L'9') { v = v * 10 + (*q - L'0'); q++; nd++; }

    if (nd == 8 && (*q == L'\0')) {
        y = (int)(v / 10000);
        mo = (int)(v / 100 % 100);
        d = (int)(v % 100);
        matched = 1;
    } else if (nd == 4) {
        y = (int)v;
        wchar_t sep1 = *q;
        if (is_date_sep(sep1)) {
            q++;
            int n2 = read_digits(&q, 2, &mo);
            if (n2 >= 1 && (*q == L'-' || *q == L'.' || *q == L'/' || *q == L'\x6708')) { // 月
                q++;
                int n3 = read_digits(&q, 2, &d);
                if (n3 >= 1) {
                    if (*q == L'\x65E5') q++; // 日
                    // 可选时间部分
                    const wchar_t* t = q;
                    while (*t == L' ') t++;
                    if (*t >= L'0' && *t <= L'9') {
                        int nh = read_digits(&t, 2, &hh);
                        int mi2 = -1, ss2 = -1;
                        if (nh >= 1 && *t == L':') {
                            t++;
                            int nm = read_digits(&t, 2, &mi2);
                            if (nm >= 1) {
                                if (*t == L':') {
                                    t++;
                                    int ns = read_digits(&t, 2, &ss2);
                                    if (ns < 1) ss2 = -1;
                                }
                            } else {
                                hh = -1;
                            }
                        } else {
                            hh = -1;
                        }
                        if (hh >= 0 && mi2 >= 0) {
                            mi = mi2;
                            ss = ss2;
                            q = t;
                        } else {
                            hh = -1;
                        }
                    }
                    while (*q == L' ') q++;
                    if (*q == L'\0') matched = 1;
                }
            }
        }
    }

    int result = 0;
    if (matched && mo >= 1 && mo <= 12 && d >= 1 && d <= days_in_month(y, mo)
        && hh >= -1 && (hh < 0 || (hh <= 23 && mi >= 0 && mi <= 59 && (ss < 0 || ss <= 59)))) {

        wchar_t wout[64];
        if (hh >= 0 && ss >= 0)      swprintf(wout, 64, L"%04d-%02d-%02d %02d:%02d:%02d", y, mo, d, hh, mi, ss);
        else if (hh >= 0)            swprintf(wout, 64, L"%04d-%02d-%02d %02d:%02d", y, mo, d, hh, mi);
        else                         swprintf(wout, 64, L"%04d-%02d-%02d", y, mo, d);

        char* u = wchar_to_utf8(wout);
        if (u) {
            strncpy(out, u, out_max - 1);
            out[out_max - 1] = '\0';
            free(u);
            result = 1;
        }
    }

    free(w);
    return result;
}

// ============================================================
// 单元格清洗与统计
// ============================================================

typedef struct {
    const CleanOptions* opts;
    CleanStats* stats;
} CellContext;

static void clean_cell_value_c(const char* in, const CleanOptions* opts, char* out,
                               size_t out_max, int* is_number, double* out_num, int* nums_found) {
    *is_number = 0;
    *out_num = 0.0;
    *nums_found = 0;
    out[0] = '\0';
    if (!in || !*in) return;

    if (opts->mode == CLEAN_MODE_TEXT) {
        text_clean_utf8(in, opts->aggressive_special_removal, out, out_max);
        return;
    }

    if (opts->mode == CLEAN_MODE_DATE) {
        if (!date_normalize_utf8(in, out, out_max)) {
            strncpy(out, in, out_max - 1);
            out[out_max - 1] = '\0';
        }
        return;
    }

    double nums[MAX_NUMS_PER_CELL];
    char raws[MAX_NUMS_PER_CELL][64];
    int count = extract_numbers(in, nums, raws, MAX_NUMS_PER_CELL);
    *nums_found = count;

    if (count == 0) {
        strncpy(out, in, out_max - 1);
        out[out_max - 1] = '\0';
        return;
    }

    switch (opts->mode) {
    case CLEAN_MODE_DEFAULT: {
        char sep[48];
        char* s = wchar_to_utf8(opts->separator);
        if (s && *s) {
            strncpy(sep, s, sizeof(sep) - 1);
            sep[sizeof(sep) - 1] = '\0';
        } else {
            strcpy(sep, ", ");
        }
        if (s) free(s);

        char temp[64];
        out[0] = '\0';
        size_t olen = 0;
        for (int i = 0; i < count; i++) {
            // 超过 15 位的纯整数超出 double 精确表示范围，原样输出字面，避免长 ID 丢精度
            const char* piece;
            if (is_long_integer_token(raws[i])) {
                piece = raws[i];
            } else {
                format_double_clean(nums[i], temp, sizeof(temp));
                piece = temp;
            }
            if (i > 0) strncat(out, sep, out_max - strlen(out) - 1);
            strncat(out, piece, out_max - strlen(out) - 1);
            olen = strlen(out);
        }
        (void)olen;
        break;
    }
    case CLEAN_MODE_ADD: {
        double sum = 0.0;
        for (int i = 0; i < count; i++) sum += nums[i];
        *is_number = 1;
        *out_num = sum;
        break;
    }
    case CLEAN_MODE_MULTIPLY: {
        double prod = 1.0;
        for (int i = 0; i < count; i++) prod *= nums[i];
        *is_number = 1;
        *out_num = prod;
        break;
    }
    case CLEAN_MODE_AVERAGE: {
        double sum = 0.0;
        for (int i = 0; i < count; i++) sum += nums[i];
        *is_number = 1;
        *out_num = sum / count;
        break;
    }
    case CLEAN_MODE_MAX: {
        double m = nums[0];
        for (int i = 1; i < count; i++) if (nums[i] > m) m = nums[i];
        *is_number = 1;
        *out_num = m;
        break;
    }
    case CLEAN_MODE_MIN: {
        double m = nums[0];
        for (int i = 1; i < count; i++) if (nums[i] < m) m = nums[i];
        *is_number = 1;
        *out_num = m;
        break;
    }
    case CLEAN_MODE_COUNT: {
        *is_number = 1;
        *out_num = (double)count;
        break;
    }
    default: {
        strncpy(out, in, out_max - 1);
        out[out_max - 1] = '\0';
        break;
    }
    }
}

// 清洗 + 统计统一入口（xlsx 与 csv 共用）
static void clean_and_stat(const char* in, const CellContext* ctx, char* out,
                           size_t out_max, int* is_number, double* out_num) {
    int nums_found = 0;
    clean_cell_value_c(in, ctx->opts, out, out_max, is_number, out_num, &nums_found);

    ctx->stats->total_cells++;
    ctx->stats->numbers_found += nums_found;

    if (ctx->opts->mode == CLEAN_MODE_DATE && strcmp(in, out) != 0) {
        ctx->stats->dates_normalized++;
    }

    int modified;
    if (*is_number) {
        char numbuf[64];
        format_double_clean(*out_num, numbuf, sizeof(numbuf));
        modified = (strcmp(numbuf, in) != 0);
    } else {
        modified = (strcmp(out, in) != 0);
    }
    if (modified) ctx->stats->modified_cells++;
    else ctx->stats->kept_cells++;
}

// ============================================================
// 进度映射：文件内 0-100 → 批次整体进度
// ============================================================

typedef struct {
    ProgressCallback cb;
    void* user;
    int base;
    int span;
} MappedCb;

static void MappedProgress(void* u, int pct, const wchar_t* msg) {
    MappedCb* m = (MappedCb*)u;
    if (m->cb) m->cb(m->user, m->base + (int)((long long)pct * m->span / 100), msg);
}

// ============================================================
// XLSX 处理
// ============================================================

static int process_xlsx(const wchar_t* in_path, const wchar_t* out_path,
                        const CleanOptions* opts, ProgressCallback cb, void* user,
                        CleanStats* stats) {
    char* in_utf8 = wchar_to_utf8(in_path);
    char* out_utf8 = wchar_to_utf8(out_path);
    if (!in_utf8 || !out_utf8) {
        free(in_utf8);
        free(out_utf8);
        return 1;
    }

    if (cb) cb(user, 5, L"正在打开 Excel 文件...");

    xlsxioreader read_handle = xlsxioread_open(in_utf8);
    if (!read_handle) {
        if (cb) cb(user, 0, L"错误：无法打开输入 Excel 文件！");
        free(in_utf8);
        free(out_utf8);
        return 1;
    }

    lxw_workbook* write_handle = workbook_new(out_utf8);
    if (!write_handle) {
        xlsxioread_close(read_handle);
        if (cb) cb(user, 0, L"错误：无法创建输出 Excel 文件！");
        free(in_utf8);
        free(out_utf8);
        return 1;
    }

    CellContext ctx = { opts, stats };

    char* sheets[MAX_SHEETS];
    int sheet_count = 0;
    int cancelled = 0;

    xlsxioreadersheetlist sheetlist = xlsxioread_sheetlist_open(read_handle);
    const char* name = NULL;
    while (sheetlist && (name = xlsxioread_sheetlist_next(sheetlist)) != NULL && sheet_count < MAX_SHEETS) {
        sheets[sheet_count] = strdup(name);
        sheet_count++;
    }
    if (sheetlist) xlsxioread_sheetlist_close(sheetlist);

    for (int i = 0; i < sheet_count && !cancelled; i++) {
        if (g_cancel_requested) { cancelled = 1; break; }

        wchar_t status_buf[256];
        swprintf(status_buf, 256, L"正在处理工作表 (%d/%d)...", i + 1, sheet_count);
        int pct = 10 + (int)(80.0 * i / (sheet_count > 0 ? sheet_count : 1));
        if (cb) cb(user, pct, status_buf);

        lxw_worksheet* out_ws = workbook_add_worksheet(write_handle, sheets[i]);
        xlsxioreadersheet read_sheet = xlsxioread_sheet_open(read_handle, sheets[i], XLSXIOREAD_SKIP_NONE);

        if (read_sheet) {
            while (!cancelled && xlsxioread_sheet_next_row(read_sheet)) {
                if (g_cancel_requested) { cancelled = 1; break; }
                size_t r = xlsxioread_sheet_last_row_index(read_sheet) - 1;
                char* cell_val = NULL;
                while (xlsxioread_sheet_next_cell_string(read_sheet, &cell_val)) {
                    size_t c = xlsxioread_sheet_last_column_index(read_sheet) - 1;
                    if (cell_val && *cell_val) {
                        char cleaned_str[MAX_CELL_STR_LEN];
                        int is_num = 0;
                        double num_val = 0.0;

                        clean_and_stat(cell_val, &ctx, cleaned_str, sizeof(cleaned_str), &is_num, &num_val);

                        if (is_num) {
                            worksheet_write_number(out_ws, (lxw_row_t)r, (lxw_col_t)c, num_val, NULL);
                        } else if (cleaned_str[0] != '\0') {
                            worksheet_write_string(out_ws, (lxw_row_t)r, (lxw_col_t)c, cleaned_str, NULL);
                        }
                    }
                    if (cell_val) xlsxioread_free(cell_val);
                }
            }
            xlsxioread_sheet_close(read_sheet);
        }
    }

    for (int i = 0; i < sheet_count; i++) free(sheets[i]);

    if (cb) cb(user, 95, L"正在保存并关闭 Excel 文件...");

    xlsxioread_close(read_handle);
    workbook_close(write_handle);

    free(in_utf8);
    free(out_utf8);

    if (cancelled) {
        DeleteFileW(out_path);
        if (cb) cb(user, 0, L"任务已取消，已丢弃该文件的部分输出。");
        return 2;
    }

    if (cb) cb(user, 100, L"文件处理完成。");
    return 0;
}

// ============================================================
// CSV 处理
// ============================================================

// 严格校验 UTF-8
static int is_valid_utf8(const unsigned char* s, size_t len) {
    size_t i = 0;
    while (i < len) {
        unsigned char c = s[i];
        int extra;
        unsigned int min, cp;
        if (c < 0x80) { i++; continue; }
        else if ((c & 0xE0) == 0xC0) { extra = 1; min = 0x80;  cp = c & 0x1F; }
        else if ((c & 0xF0) == 0xE0) { extra = 2; min = 0x800; cp = c & 0x0F; }
        else if ((c & 0xF8) == 0xF0) { extra = 3; min = 0x10000; cp = c & 0x07; }
        else return 0;
        if (i + (size_t)extra >= len) return 0;
        for (int k = 1; k <= extra; k++) {
            if ((s[i + k] & 0xC0) != 0x80) return 0;
            cp = (cp << 6) | (unsigned)(s[i + k] & 0x3F);
        }
        if (cp < min || cp > 0x10FFFF || (cp >= 0xD800 && cp <= 0xDFFF)) return 0;
        i += (size_t)extra + 1;
    }
    return 1;
}

// 读取文本文件并统一转为 UTF-8（识别 BOM / UTF-8 / GBK）
static char* load_text_file(const wchar_t* path, size_t* out_len) {
    HANDLE h = CreateFileW(path, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, 0, NULL);
    if (h == INVALID_HANDLE_VALUE) return NULL;

    LARGE_INTEGER sz;
    if (!GetFileSizeEx(h, &sz) || sz.QuadPart <= 0 || sz.QuadPart > (LONGLONG)(64 * 1024 * 1024)) {
        CloseHandle(h);
        return NULL;
    }

    size_t n = (size_t)sz.QuadPart;
    unsigned char* raw = (unsigned char*)malloc(n + 1);
    if (!raw) { CloseHandle(h); return NULL; }

    DWORD rd = 0;
    if (!ReadFile(h, raw, (DWORD)n, &rd, NULL)) {
        CloseHandle(h);
        free(raw);
        return NULL;
    }
    CloseHandle(h);
    n = rd;

    char* utf8 = NULL;
    size_t ulen = 0;

    if (n >= 3 && raw[0] == 0xEF && raw[1] == 0xBB && raw[2] == 0xBF) {
        // UTF-8 BOM
        ulen = n - 3;
        utf8 = (char*)malloc(ulen + 1);
        if (utf8) { memcpy(utf8, raw + 3, ulen); utf8[ulen] = '\0'; }
    } else if (n >= 2 && raw[0] == 0xFF && raw[1] == 0xFE) {
        // UTF-16 LE BOM
        int wlen = (int)((n - 2) / 2);
        if (wlen > 0) {
            wchar_t* w = (wchar_t*)malloc((size_t)wlen * sizeof(wchar_t));
            if (w) {
                memcpy(w, raw + 2, (size_t)wlen * sizeof(wchar_t));
                int blen = WideCharToMultiByte(CP_UTF8, 0, w, wlen, NULL, 0, NULL, NULL);
                if (blen > 0) {
                    utf8 = (char*)malloc(blen + 1);
                    if (utf8) {
                        WideCharToMultiByte(CP_UTF8, 0, w, wlen, utf8, blen, NULL, NULL);
                        utf8[blen] = '\0';
                        ulen = (size_t)blen;
                    }
                }
                free(w);
            }
        }
        if (!utf8) { utf8 = (char*)calloc(1, 1); ulen = 0; }
    } else if (is_valid_utf8(raw, n)) {
        // UTF-8（或纯 ASCII）
        ulen = n;
        utf8 = (char*)malloc(ulen + 1);
        if (utf8) { memcpy(utf8, raw, ulen); utf8[ulen] = '\0'; }
    } else {
        // 按 GBK（CP936）转 UTF-8
        if (n > 0) {
            int wlen = MultiByteToWideChar(936, 0, (const char*)raw, (int)n, NULL, 0);
            if (wlen > 0) {
                wchar_t* w = (wchar_t*)malloc((size_t)wlen * sizeof(wchar_t));
                if (w) {
                    MultiByteToWideChar(936, 0, (const char*)raw, (int)n, w, wlen);
                    int blen = WideCharToMultiByte(CP_UTF8, 0, w, wlen, NULL, 0, NULL, NULL);
                    if (blen > 0) {
                        utf8 = (char*)malloc(blen + 1);
                        if (utf8) {
                            WideCharToMultiByte(CP_UTF8, 0, w, wlen, utf8, blen, NULL, NULL);
                            utf8[blen] = '\0';
                            ulen = (size_t)blen;
                        }
                    }
                    free(w);
                }
            }
        }
        if (!utf8) { utf8 = (char*)calloc(1, 1); ulen = 0; }
    }

    free(raw);
    if (!utf8) return NULL;
    if (out_len) *out_len = ulen;
    return utf8;
}

// 可增长缓冲区
typedef struct {
    char* data;
    size_t len;
    size_t cap;
} Buf;

static int buf_init(Buf* b) {
    b->cap = 256;
    b->len = 0;
    b->data = (char*)malloc(b->cap);
    return b->data ? 0 : -1;
}

static int buf_append_ch(Buf* b, char c) {
    if (b->len + 1 >= b->cap) {
        size_t nc = b->cap * 2;
        char* nd = (char*)realloc(b->data, nc);
        if (!nd) return -1;
        b->data = nd;
        b->cap = nc;
    }
    b->data[b->len++] = c;
    b->data[b->len] = '\0';
    return 0;
}

static void buf_free(Buf* b) {
    free(b->data);
    b->data = NULL;
    b->len = b->cap = 0;
}

// 写出 CSV 字段（必要时加引号转义）
static void csv_write_field(FILE* f, const char* s) {
    if (!s) return;
    int need_quote = 0;
    for (const char* p = s; *p; p++) {
        if (*p == ',' || *p == '"' || *p == '\n' || *p == '\r') { need_quote = 1; break; }
    }
    if (need_quote) {
        fputc('"', f);
        for (const char* p = s; *p; p++) {
            if (*p == '"') fputc('"', f);
            fputc(*p, f);
        }
        fputc('"', f);
    } else {
        fputs(s, f);
    }
}

static int process_csv(const wchar_t* in_path, const wchar_t* out_path,
                       const CleanOptions* opts, ProgressCallback cb, void* user,
                       CleanStats* stats) {
    if (cb) cb(user, 5, L"正在读取 CSV 文件...");

    size_t len = 0;
    char* content = load_text_file(in_path, &len);
    if (!content) {
        if (cb) cb(user, 0, L"错误：无法读取 CSV 文件！");
        return 1;
    }

    FILE* out = _wfopen(out_path, L"wb");
    if (!out) {
        if (cb) cb(user, 0, L"错误：无法创建输出 CSV 文件！");
        free(content);
        return 1;
    }

    // UTF-8 BOM（Excel 友好）
    fwrite("\xEF\xBB\xBF", 1, 3, out);

    CellContext ctx = { opts, stats };
    Buf field;
    if (buf_init(&field) != 0) {
        fclose(out);
        free(content);
        return 1;
    }

    size_t i = 0;
    int row_fields = 0;
    int in_quotes = 0;
    int done = 0;
    int cancelled = 0;
    int last_pct = 5;

    while (!done && !cancelled) {
        // 读取一个字段
        field.len = 0;
        field.data[0] = '\0';
        for (;;) {
            if (i >= len) { done = 1; break; }
            char ch = content[i];
            if (!in_quotes && (ch == ',' || ch == '\n' || ch == '\r')) break;
            if (ch == '"') {
                if (in_quotes && i + 1 < len && content[i + 1] == '"') {
                    if (buf_append_ch(&field, '"') != 0) { cancelled = 0; done = 1; break; }
                    i += 2;
                    continue;
                }
                in_quotes = !in_quotes;
                i++;
                continue;
            }
            if (buf_append_ch(&field, ch) != 0) { done = 1; break; }
            i++;
        }

        // 文件末尾的空字段不输出（避免多余空行/空列）
        if (!(done && field.len == 0 && row_fields == 0)) {
            if (field.len > 0) {
                char cleaned[MAX_CELL_STR_LEN];
                int is_num = 0;
                double num_val = 0.0;
                clean_and_stat(field.data, &ctx, cleaned, sizeof(cleaned), &is_num, &num_val);
                if (is_num) {
                    char numbuf[64];
                    format_double_clean(num_val, numbuf, sizeof(numbuf));
                    csv_write_field(out, numbuf);
                } else if (cleaned[0] != '\0') {
                    csv_write_field(out, cleaned);
                } else {
                    csv_write_field(out, "");
                }
            } else {
                csv_write_field(out, "");
            }
            row_fields++;
        }

        if (done) {
            if (row_fields > 0) fputc('\n', out);
            break;
        }

        // 处理分隔符 / 行结束
        char ch = content[i];
        if (ch == ',') {
            i++;
        } else {
            if (row_fields > 0) fputc('\n', out);
            row_fields = 0;
            if (ch == '\r' && i + 1 < len && content[i + 1] == '\n') i += 2;
            else i++;

            // 行边界：检查取消与进度
            if (g_cancel_requested) { cancelled = 1; break; }
            int pct = 10 + (int)(85.0 * (double)i / (double)(len > 0 ? len : 1));
            if (pct > 95) pct = 95;
            if (pct != last_pct) {
                last_pct = pct;
                if (cb) cb(user, pct, L"正在处理 CSV 数据行...");
            }
        }
    }

    buf_free(&field);
    free(content);

    if (fclose(out) != 0) {
        if (cancelled) DeleteFileW(out_path);
        return 1;
    }

    if (cancelled) {
        DeleteFileW(out_path);
        if (cb) cb(user, 0, L"任务已取消，已丢弃该文件的部分输出。");
        return 2;
    }

    if (cb) cb(user, 100, L"文件处理完成。");
    return 0;
}

// ============================================================
// 批量调度
// ============================================================

static void make_output_path(const wchar_t* in, wchar_t* out) {
    // out 固定为 DC_PATH 长度：先算总长，超长时截断文件名主体，保证 "_Cleaned" 与扩展名完整，
    // 不再无检查地 wcscat（输入路径接近上限时会写越界）
    const size_t cap = DC_PATH - 1;
    const wchar_t suffix[] = L"_Cleaned";
    const size_t suf_len = wcslen(suffix);
    size_t in_len = wcslen(in);
    const wchar_t* dot = wcsrchr(in, L'.');
    size_t base_len = (dot && dot != in) ? (size_t)(dot - in) : in_len;
    const wchar_t* ext = (dot && dot != in) ? dot : L"";
    size_t ext_len = wcslen(ext);

    if (base_len + suf_len + ext_len > cap) {
        size_t room = (cap > suf_len + ext_len) ? (cap - suf_len - ext_len) : 0;
        if (base_len > room) base_len = room;
    }

    size_t p = 0;
    memcpy(out, in, base_len * sizeof(wchar_t));
    p += base_len;
    memcpy(out + p, suffix, suf_len * sizeof(wchar_t));
    p += suf_len;
    memcpy(out + p, ext, ext_len * sizeof(wchar_t));
    p += ext_len;
    out[p] = L'\0';
}

int run_clean_jobs(FileJob* jobs, int job_count, const CleanOptions* opts,
                   ProgressCallback cb, void* user, CleanStats* total_stats) {
    CleanStats total;
    memset(&total, 0, sizeof(total));

    for (int i = 0; i < job_count; i++) {
        if (g_cancel_requested) {
            jobs[i].result = 3;
            continue;
        }

        const wchar_t* fname = wcsrchr(jobs[i].input, L'\\');
        fname = fname ? fname + 1 : jobs[i].input;

        int base = (int)((long long)i * 100 / (job_count > 0 ? job_count : 1));
        int span = job_count > 0 ? 100 / job_count : 100;

        wchar_t hdr[DC_PATH + 64];
        swprintf(hdr, DC_PATH + 64, L"【%d/%d】%s", i + 1, job_count, fname);
        if (cb) cb(user, base, hdr);

        make_output_path(jobs[i].input, jobs[i].output);

        memset(&jobs[i].stats, 0, sizeof(CleanStats));
        MappedCb mapped = { cb, user, base, span };

        int res;
        const wchar_t* ext = wcsrchr(jobs[i].input, L'.');
        if (ext && _wcsicmp(ext, L".csv") == 0) {
            res = process_csv(jobs[i].input, jobs[i].output, opts, MappedProgress, &mapped, &jobs[i].stats);
        } else {
            res = process_xlsx(jobs[i].input, jobs[i].output, opts, MappedProgress, &mapped, &jobs[i].stats);
        }
        jobs[i].result = res;

        total.total_cells += jobs[i].stats.total_cells;
        total.modified_cells += jobs[i].stats.modified_cells;
        total.kept_cells += jobs[i].stats.kept_cells;
        total.numbers_found += jobs[i].stats.numbers_found;
        total.dates_normalized += jobs[i].stats.dates_normalized;

        wchar_t msg[DC_PATH + 160];
        const wchar_t* fn2 = wcsrchr(jobs[i].input, L'\\');
        fn2 = fn2 ? fn2 + 1 : jobs[i].input;
        if (res == 0) {
            swprintf(msg, DC_PATH + 160, L"✅ 完成：%s（修改 %lld / %lld 个单元格）",
                     fn2, jobs[i].stats.modified_cells, jobs[i].stats.total_cells);
        } else if (res == 2) {
            swprintf(msg, DC_PATH + 160, L"⏹ 已取消：%s", fn2);
        } else {
            swprintf(msg, DC_PATH + 160, L"❌ 失败：%s（请检查日志）", fn2);
        }
        if (cb) cb(user, base + span, msg);

        if (res == 2) {
            for (int j = i + 1; j < job_count; j++) jobs[j].result = 3;
            break;
        }
    }

    if (total_stats) *total_stats = total;
    return 0;
}
