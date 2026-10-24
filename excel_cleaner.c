#include "excel_cleaner.h"
#include <xlsxio_read.h>
#include <xlsxwriter.h>
#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <math.h>
#include <stdint.h>
#include <zlib.h>

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

static int process_xlsx_rebuild(const wchar_t* in_path, const wchar_t* out_path,
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
// XLSX 原位处理：在原文件的副本上只替换被清洗改动的单元格值，
// 其余内容（样式、公式、合并单元格、图表、图片等）逐字节保留，
// 解决重建写入丢失格式/公式/合并单元格的问题。
// zip 读写只依赖 zlib（Makefile 已链接 libz），不引入新的库。
// 返回 DC_ZIP_UNSUPPORTED 表示文件结构超出本路径处理范围，
// 调用方回退到原重建路径（process_xlsx_rebuild）。
// ============================================================

#define DC_ZIP_UNSUPPORTED (-1)
#define DC_XLSX_MAX_FILE (512LL * 1024 * 1024)

#ifdef _WIN32
#define dc_seek_end(f) _fseeki64((f), 0, SEEK_END)
#define dc_tell(f)     _ftelli64(f)
#else
#define dc_seek_end(f) fseeko((f), 0, SEEK_END)
#define dc_tell(f)     ftello(f)
#endif

static uint16_t dc_rd16(const unsigned char* p) {
    return (uint16_t)((uint16_t)p[0] | ((uint16_t)p[1] << 8));
}
static uint32_t dc_rd32(const unsigned char* p) {
    return (uint32_t)p[0] | ((uint32_t)p[1] << 8) | ((uint32_t)p[2] << 16) | ((uint32_t)p[3] << 24);
}
static void dc_wr16(unsigned char* p, uint16_t v) {
    p[0] = (unsigned char)(v & 0xFF);
    p[1] = (unsigned char)((v >> 8) & 0xFF);
}
static void dc_wr32(unsigned char* p, uint32_t v) {
    p[0] = (unsigned char)(v & 0xFF);
    p[1] = (unsigned char)((v >> 8) & 0xFF);
    p[2] = (unsigned char)((v >> 16) & 0xFF);
    p[3] = (unsigned char)((v >> 24) & 0xFF);
}

// 可增长字节缓冲
typedef struct {
    unsigned char* d;
    size_t len, cap;
} DcBuf;

static int dc_buf_put(DcBuf* b, const void* p, size_t n) {
    if (n == 0) return 0;
    if (b->len + n > b->cap) {
        size_t nc = b->cap ? b->cap : 4096;
        while (nc < b->len + n) nc *= 2;
        unsigned char* nd = (unsigned char*)realloc(b->d, nc);
        if (!nd) return -1;
        b->d = nd;
        b->cap = nc;
    }
    memcpy(b->d + b->len, p, n);
    b->len += n;
    return 0;
}
static int dc_buf_puts(DcBuf* b, const char* s) { return dc_buf_put(b, s, strlen(s)); }
static void dc_buf_free(DcBuf* b) { free(b->d); b->d = NULL; b->len = b->cap = 0; }

// ---------------- zip 读取 ----------------

typedef struct {
    char* name;                    // 条目名（malloc）
    uint16_t method, flags;
    uint16_t mod_time, mod_date;
    uint32_t crc, comp_size, size;
    uint32_t local_off;
    const unsigned char* raw;      // 原本地记录起始（指向整文件缓冲）
    size_t raw_len;                // 原本地记录总长（含数据描述符）
    const unsigned char* comp;     // 压缩数据起始
    size_t cd_off, cd_len;         // 原中央目录记录在整文件缓冲中的位置/长度
    unsigned char* new_data;       // 替换后的未压缩内容（NULL 表示原样复制）
    size_t new_len;
    int drop;                      // 输出时丢弃
} DcZipEntry;

typedef struct {
    unsigned char* file;
    size_t file_len;
    DcZipEntry* e;
    int n;
    const unsigned char* comment;
    uint16_t comment_len;
} DcZip;

static void dc_zip_free(DcZip* z) {
    if (z->e) {
        for (int i = 0; i < z->n; i++) {
            free(z->e[i].name);
            free(z->e[i].new_data);
        }
        free(z->e);
    }
    free(z->file);
    memset(z, 0, sizeof(*z));
}

// 0 成功 / DC_ZIP_UNSUPPORTED 结构不支持 / 1 读取失败
static int dc_zip_load(const wchar_t* path, DcZip* z) {
    memset(z, 0, sizeof(*z));
    FILE* f = _wfopen(path, L"rb");
    if (!f) return 1;
    if (dc_seek_end(f) != 0) { fclose(f); return 1; }
    long long fsz = (long long)dc_tell(f);
    if (fsz <= 0) { fclose(f); return 1; }
    if (fsz > DC_XLSX_MAX_FILE) { fclose(f); return DC_ZIP_UNSUPPORTED; }
    rewind(f);
    z->file = (unsigned char*)malloc((size_t)fsz);
    if (!z->file) { fclose(f); return 1; }
    if (fread(z->file, 1, (size_t)fsz, f) != (size_t)fsz) { fclose(f); dc_zip_free(z); return 1; }
    fclose(f);
    z->file_len = (size_t)fsz;

    // 定位 EOCD（末尾 22 字节 + 最长 64KB 注释）
    if (z->file_len < 22) { free(z->file); z->file = NULL; return DC_ZIP_UNSUPPORTED; }
    size_t back = z->file_len - 22;
    size_t minpos = (z->file_len > 65557) ? (z->file_len - 65557) : 0;
    const unsigned char* eocd = NULL;
    for (size_t i = back;; i--) {
        if (dc_rd32(z->file + i) == 0x06054b50 && i + 22 + dc_rd16(z->file + i + 20) == z->file_len) {
            eocd = z->file + i;
            break;
        }
        if (i == minpos) break;
    }
    if (!eocd) { free(z->file); z->file = NULL; return DC_ZIP_UNSUPPORTED; }
    if (dc_rd16(eocd + 4) != 0 || dc_rd16(eocd + 6) != 0) { free(z->file); z->file = NULL; return DC_ZIP_UNSUPPORTED; } // 多磁盘
    int total = dc_rd16(eocd + 10);
    uint32_t cd_size = dc_rd32(eocd + 12);
    uint32_t cd_off = dc_rd32(eocd + 16);
    if (total == 0xFFFF || cd_size == 0xFFFFFFFFu || cd_off == 0xFFFFFFFFu) { free(z->file); z->file = NULL; return DC_ZIP_UNSUPPORTED; } // zip64
    z->comment = eocd + 22;
    z->comment_len = dc_rd16(eocd + 20);
    if ((size_t)cd_off + cd_size > z->file_len) { free(z->file); z->file = NULL; return DC_ZIP_UNSUPPORTED; }

    z->e = (DcZipEntry*)calloc((size_t)(total > 0 ? total : 1), sizeof(DcZipEntry));
    if (!z->e) return 1;
    z->n = total;

    const unsigned char* p = z->file + cd_off;
    const unsigned char* cd_end = p + cd_size;
    for (int i = 0; i < total; i++) {
        if ((size_t)(cd_end - p) < 46 || dc_rd32(p) != 0x02014b50) { dc_zip_free(z); return DC_ZIP_UNSUPPORTED; }
        DcZipEntry* e = &z->e[i];
        e->flags = dc_rd16(p + 8);
        e->method = dc_rd16(p + 10);
        e->mod_time = dc_rd16(p + 12);
        e->mod_date = dc_rd16(p + 14);
        e->crc = dc_rd32(p + 16);
        e->comp_size = dc_rd32(p + 20);
        e->size = dc_rd32(p + 24);
        uint16_t nl = dc_rd16(p + 28);
        uint16_t xl = dc_rd16(p + 30);
        uint16_t cl = dc_rd16(p + 32);
        e->local_off = dc_rd32(p + 42);
        if (e->crc == 0xFFFFFFFFu || e->comp_size == 0xFFFFFFFFu || e->size == 0xFFFFFFFFu || e->local_off == 0xFFFFFFFFu) {
            dc_zip_free(z);
            return DC_ZIP_UNSUPPORTED; // zip64
        }
        if ((size_t)(cd_end - p) < (size_t)46 + nl + xl + cl) { dc_zip_free(z); return DC_ZIP_UNSUPPORTED; }
        e->name = (char*)malloc((size_t)nl + 1);
        if (!e->name) { dc_zip_free(z); return 1; }
        memcpy(e->name, p + 46, nl);
        e->name[nl] = '\0';
        e->cd_off = (size_t)(p - z->file);
        e->cd_len = (size_t)46 + nl + xl + cl;
        p += e->cd_len;

        if (e->flags & 0x0001) { dc_zip_free(z); return DC_ZIP_UNSUPPORTED; } // 加密

        // 解析本地头，定位压缩数据
        if ((size_t)e->local_off + 30 > z->file_len) { dc_zip_free(z); return DC_ZIP_UNSUPPORTED; }
        const unsigned char* lh = z->file + e->local_off;
        if (dc_rd32(lh) != 0x04034b50) { dc_zip_free(z); return DC_ZIP_UNSUPPORTED; }
        uint16_t lnl = dc_rd16(lh + 26);
        uint16_t lxl = dc_rd16(lh + 28);
        const unsigned char* data = lh + 30 + lnl + lxl;
        if ((size_t)(data - z->file) + e->comp_size > z->file_len) { dc_zip_free(z); return DC_ZIP_UNSUPPORTED; }
        e->comp = data;
        e->raw = lh;
        size_t raw_len = (size_t)(data - lh) + e->comp_size;
        if (e->flags & 0x0008) { // 数据描述符
            const unsigned char* after = data + e->comp_size;
            size_t remain = (size_t)(z->file + z->file_len - after);
            if (remain >= 4 && dc_rd32(after) == 0x08074b50) raw_len += 16;
            else if (remain >= 12 && dc_rd32(after) == e->crc && dc_rd32(after + 4) == e->comp_size && dc_rd32(after + 8) == e->size) raw_len += 12;
            else { dc_zip_free(z); return DC_ZIP_UNSUPPORTED; }
        }
        e->raw_len = raw_len;
    }
    return 0;
}

static DcZipEntry* dc_zip_find(DcZip* z, const char* name) {
    for (int i = 0; i < z->n; i++) {
        if (strcmp(z->e[i].name, name) == 0) return &z->e[i];
    }
    return NULL;
}

// 解压一个条目（malloc，末尾补 '\0'；out_len 可为 NULL）。方法不支持或数据异常返回 NULL
static unsigned char* dc_zip_read(DcZip* z, DcZipEntry* e, size_t* out_len) {
    (void)z;
    unsigned char* buf = (unsigned char*)malloc((size_t)e->size + 1);
    if (!buf) return NULL;
    if (e->method == 0) {
        if (e->comp_size != e->size) { free(buf); return NULL; }
        memcpy(buf, e->comp, e->size);
    } else if (e->method == 8) {
        if (e->size == 0) {
            // 空内容
        } else {
            z_stream zs;
            memset(&zs, 0, sizeof(zs));
            if (inflateInit2(&zs, -MAX_WBITS) != Z_OK) { free(buf); return NULL; }
            zs.next_in = (Bytef*)e->comp;
            zs.avail_in = (uInt)e->comp_size;
            zs.next_out = buf;
            zs.avail_out = (uInt)e->size;
            int rc = inflate(&zs, Z_FINISH);
            uLong produced = zs.total_out;
            inflateEnd(&zs);
            if (rc != Z_STREAM_END || produced != e->size) { free(buf); return NULL; }
        }
    } else {
        free(buf);
        return NULL;
    }
    buf[e->size] = '\0';
    if (out_len) *out_len = e->size;
    return buf;
}

// deflate 压缩（malloc 结果）
static unsigned char* dc_deflate(const unsigned char* data, size_t len, size_t* out_len) {
    z_stream zs;
    memset(&zs, 0, sizeof(zs));
    if (deflateInit2(&zs, Z_DEFAULT_COMPRESSION, Z_DEFLATED, -MAX_WBITS, 8, Z_DEFAULT_STRATEGY) != Z_OK) {
        return NULL;
    }
    DcBuf b = { 0, 0, 0 };
    zs.next_in = (Bytef*)data;
    zs.avail_in = (uInt)len;
    int rc;
    do {
        unsigned char chunk[65536];
        zs.next_out = chunk;
        zs.avail_out = sizeof(chunk);
        rc = deflate(&zs, (zs.avail_in == 0) ? Z_FINISH : Z_NO_FLUSH);
        size_t have = sizeof(chunk) - zs.avail_out;
        if (have > 0 && dc_buf_put(&b, chunk, have) != 0) {
            deflateEnd(&zs);
            dc_buf_free(&b);
            return NULL;
        }
    } while (rc != Z_STREAM_END);
    deflateEnd(&zs);
    if (out_len) *out_len = b.len;
    return b.d;
}

// 写出修改后的 zip：替换条目重新压缩，其余条目原样复制本地记录
static int dc_zip_write(DcZip* z, const wchar_t* out_path) {
    FILE* f = _wfopen(out_path, L"wb");
    if (!f) return 1;
    DcBuf cd = { 0, 0, 0 };
    long long pos = 0;
    int failed = 0;
    uint16_t count = 0;

    for (int i = 0; i < z->n && !failed; i++) {
        DcZipEntry* e = &z->e[i];
        if (e->drop) continue;

        // 准备该条目的中央目录记录（克隆原记录后改写偏移/变化字段）
        unsigned char cdrec[4096];
        if (e->cd_len > sizeof(cdrec)) { failed = 1; break; }
        memcpy(cdrec, z->file + e->cd_off, e->cd_len);
        dc_wr32(cdrec + 42, (uint32_t)pos);

        if (e->new_data) {
            unsigned char* comp = NULL;
            size_t comp_len = 0;
            uint16_t method = 8;
            uint32_t crc = (uint32_t)crc32(crc32(0L, Z_NULL, 0), (const Bytef*)e->new_data, (uInt)e->new_len);
            if (e->new_len == 0) {
                method = 0;
            } else {
                comp = dc_deflate(e->new_data, e->new_len, &comp_len);
                if (!comp) { failed = 1; break; }
            }
            unsigned char lh[30];
            dc_wr32(lh + 0, 0x04034b50);
            dc_wr16(lh + 4, 20);
            dc_wr16(lh + 6, 0);
            dc_wr16(lh + 8, method);
            dc_wr16(lh + 10, e->mod_time);
            dc_wr16(lh + 12, e->mod_date);
            dc_wr32(lh + 14, crc);
            dc_wr32(lh + 18, (uint32_t)comp_len);
            dc_wr32(lh + 22, (uint32_t)e->new_len);
            dc_wr16(lh + 26, (uint16_t)strlen(e->name));
            dc_wr16(lh + 28, 0);
            if (fwrite(lh, 1, 30, f) != 30 ||
                fwrite(e->name, 1, strlen(e->name), f) != strlen(e->name) ||
                (comp_len > 0 && fwrite(comp, 1, comp_len, f) != comp_len)) {
                free(comp);
                failed = 1;
                break;
            }
            free(comp);
            pos += 30 + (long long)strlen(e->name) + (long long)comp_len;

            dc_wr16(cdrec + 8, 0);
            dc_wr16(cdrec + 10, method);
            dc_wr32(cdrec + 16, crc);
            dc_wr32(cdrec + 20, (uint32_t)comp_len);
            dc_wr32(cdrec + 24, (uint32_t)e->new_len);
        } else {
            if (fwrite(e->raw, 1, e->raw_len, f) != e->raw_len) { failed = 1; break; }
            pos += (long long)e->raw_len;
        }

        if (dc_buf_put(&cd, cdrec, e->cd_len) != 0) { failed = 1; break; }
        count++;
    }

    if (!failed) {
        long long cd_off = pos;
        if (cd.len > 0 && fwrite(cd.d, 1, cd.len, f) != cd.len) {
            failed = 1;
        } else {
            pos += (long long)cd.len;
            unsigned char eo[22];
            dc_wr32(eo + 0, 0x06054b50);
            dc_wr16(eo + 4, 0);
            dc_wr16(eo + 6, 0);
            dc_wr16(eo + 8, count);
            dc_wr16(eo + 10, count);
            dc_wr32(eo + 12, (uint32_t)cd.len);
            dc_wr32(eo + 16, (uint32_t)cd_off);
            dc_wr16(eo + 20, z->comment_len);
            if (fwrite(eo, 1, 22, f) != 22) failed = 1;
            else if (z->comment_len > 0 && fwrite(z->comment, 1, z->comment_len, f) != z->comment_len) failed = 1;
        }
    }

    dc_buf_free(&cd);
    if (fclose(f) != 0) failed = 1;
    return failed ? 1 : 0;
}


// ---------------- XML 小工具（字符串扫描，无 DOM） ----------------

static int dc_buf_put_cp(DcBuf* b, unsigned int cp) {
    char tmp[4];
    int n;
    if (cp < 0x80) { tmp[0] = (char)cp; n = 1; }
    else if (cp < 0x800) { tmp[0] = (char)(0xC0 | (cp >> 6)); tmp[1] = (char)(0x80 | (cp & 0x3F)); n = 2; }
    else if (cp < 0x10000) { tmp[0] = (char)(0xE0 | (cp >> 12)); tmp[1] = (char)(0x80 | ((cp >> 6) & 0x3F)); tmp[2] = (char)(0x80 | (cp & 0x3F)); n = 3; }
    else { tmp[0] = (char)(0xF0 | (cp >> 18)); tmp[1] = (char)(0x80 | ((cp >> 12) & 0x3F)); tmp[2] = (char)(0x80 | ((cp >> 6) & 0x3F)); tmp[3] = (char)(0x80 | (cp & 0x3F)); n = 4; }
    return dc_buf_put(b, tmp, (size_t)n);
}

// XML 实体反转义（malloc 结果）
static char* dc_xml_unescape(const char* s, size_t n) {
    DcBuf b = { 0, 0, 0 };
    size_t i = 0;
    while (i < n) {
        if (s[i] == '&') {
            size_t j = i + 1;
            while (j < n && j < i + 12 && s[j] != ';') j++;
            if (j < n && s[j] == ';' && j > i + 1) {
                size_t el = j - i - 1;
                const char* e = s + i + 1;
                int done = 1;
                if (el == 3 && strncmp(e, "amp", 3) == 0) dc_buf_put(&b, "&", 1);
                else if (el == 2 && strncmp(e, "lt", 2) == 0) dc_buf_put(&b, "<", 1);
                else if (el == 2 && strncmp(e, "gt", 2) == 0) dc_buf_put(&b, ">", 1);
                else if (el == 4 && strncmp(e, "quot", 4) == 0) dc_buf_put(&b, "\"", 1);
                else if (el == 4 && strncmp(e, "apos", 4) == 0) dc_buf_put(&b, "'", 1);
                else if (el >= 2 && e[0] == '#') {
                    char numbuf[16];
                    size_t nl = (el - 1 < 15) ? el - 1 : 15;
                    memcpy(numbuf, e + 1, nl);
                    numbuf[nl] = '\0';
                    long cp;
                    if (numbuf[0] == 'x' || numbuf[0] == 'X') cp = strtol(numbuf + 1, NULL, 16);
                    else cp = strtol(numbuf, NULL, 10);
                    if (cp > 0 && cp <= 0x10FFFF) dc_buf_put_cp(&b, (unsigned int)cp);
                } else done = 0;
                if (done) { i = j + 1; continue; }
            }
            dc_buf_put(&b, "&", 1);
            i++;
            continue;
        }
        dc_buf_put(&b, s + i, 1);
        i++;
    }
    dc_buf_put(&b, "", 1); // 末尾 '\0'
    return (char*)b.d;
}

// XML 文本转义（malloc 结果；写回单元格文本时使用）
static char* dc_xml_escape(const char* s) {
    DcBuf b = { 0, 0, 0 };
    for (const char* p = s; *p; p++) {
        if (*p == '&') dc_buf_puts(&b, "&amp;");
        else if (*p == '<') dc_buf_puts(&b, "&lt;");
        else if (*p == '>') dc_buf_puts(&b, "&gt;");
        else dc_buf_put(&b, p, 1);
    }
    dc_buf_put(&b, "", 1);
    return (char*)b.d;
}

// 找到标签结束位置（p 指向 '<'），返回 '>' 之后的位置；引号内的 '>' 不算结束
static const char* dc_find_tag_end(const char* p, const char* limit, int* self_closing) {
    char quote = 0;
    *self_closing = 0;
    for (const char* q = p + 1; q < limit; q++) {
        if (quote) {
            if (*q == quote) quote = 0;
        } else if (*q == '"' || *q == '\'') {
            quote = *q;
        } else if (*q == '>') {
            if (q > p && *(q - 1) == '/') *self_closing = 1;
            return q + 1;
        }
    }
    return NULL;
}

// 标签属性解析：name/value 的原文区间
typedef struct {
    const char* ns, * ne;  // 属性名
    const char* vs, * ve;  // 属性值（引号内）
    const char* ae;        // 整个赋值结束位置（闭引号之后）
} DcAttr;

// 解析 [p, end) 处起始标签的属性（p 指向 '<'），返回属性个数（<0 表示结构异常）
static int dc_tag_attrs(const char* p, const char* end, DcAttr* attrs, int max_attrs) {
    const char* q = p + 1;
    // 跳过元素名
    while (q < end && *q != ' ' && *q != '\t' && *q != '\r' && *q != '\n' && *q != '>' && *q != '/') q++;
    int n = 0;
    while (q < end) {
        while (q < end && (*q == ' ' || *q == '\t' || *q == '\r' || *q == '\n')) q++;
        if (q >= end || *q == '>' || *q == '/') break;
        if (n >= max_attrs) return -1;
        DcAttr* a = &attrs[n];
        a->ns = q;
        while (q < end && *q != '=' && *q != ' ' && *q != '\t' && *q != '\r' && *q != '\n' && *q != '>' && *q != '/') q++;
        a->ne = q;
        a->vs = a->ve = q;
        a->ae = q;
        while (q < end && (*q == ' ' || *q == '\t')) q++;
        if (q < end && *q == '=') {
            q++;
            while (q < end && (*q == ' ' || *q == '\t')) q++;
            if (q < end && (*q == '"' || *q == '\'')) {
                char quote = *q++;
                a->vs = q;
                while (q < end && *q != quote) q++;
                if (q >= end) return -1;
                a->ve = q;
                a->ae = q + 1;
                q++;
            } else {
                return -1;
            }
        }
        n++;
    }
    return n;
}

static const DcAttr* dc_attr_find(const DcAttr* attrs, int n, const char* name) {
    size_t nl = strlen(name);
    for (int i = 0; i < n; i++) {
        if ((size_t)(attrs[i].ne - attrs[i].ns) == nl && strncmp(attrs[i].ns, name, nl) == 0) {
            return &attrs[i];
        }
    }
    return NULL;
}

// 判断 p 处是否为 "<name" 且后跟标签边界字符
static int dc_tag_at(const char* p, const char* name) {
    size_t nl = strlen(name);
    if (strncmp(p, name, nl) != 0) return 0;
    char c = p[nl];
    return c == '>' || c == ' ' || c == '\t' || c == '\r' || c == '\n' || c == '/';
}

// 拼接 [s, end) 内所有 <t> 的文本（跳过 <rPh>…</rPh> 拼音段），malloc 结果
static char* dc_collect_t_text(const char* s, const char* end) {
    DcBuf b = { 0, 0, 0 };
    const char* p = s;
    while (p < end) {
        if (dc_tag_at(p, "<rPh")) {
            const char* close = strstr(p, "</rPh>");
            if (!close || close > end) break;
            p = close + 6;
            continue;
        }
        if (dc_tag_at(p, "<t")) {
            int selfclose = 0;
            const char* te = dc_find_tag_end(p, end, &selfclose);
            if (!te) break;
            if (!selfclose) {
                const char* close = strstr(te, "</t>");
                if (!close || close > end) break;
                char* txt = dc_xml_unescape(te, (size_t)(close - te));
                if (txt) {
                    dc_buf_puts(&b, txt);
                    free(txt);
                }
                p = close + 4;
                continue;
            }
            p = te;
            continue;
        }
        p++;
    }
    dc_buf_put(&b, "", 1);
    return (char*)b.d;
}

// 解析 sharedStrings.xml，返回字符串数组（每个元素 malloc），失败返回 NULL
static char** dc_parse_shared_strings(const char* xml, int* out_count) {
    int cap = 256, n = 0;
    char** arr = (char**)malloc(sizeof(char*) * (size_t)cap);
    if (!arr) return NULL;
    const char* end = xml + strlen(xml);
    const char* p = xml;
    while (p < end) {
        const char* si = strstr(p, "<si");
        if (!si || si >= end) break;
        if (!dc_tag_at(si, "<si")) { p = si + 3; continue; }
        int selfclose = 0;
        const char* te = dc_find_tag_end(si, end, &selfclose);
        if (!te) goto fail;
        char* txt;
        const char* next;
        if (selfclose) {
            txt = dc_xml_unescape("", 0);
            next = te;
        } else {
            const char* close = strstr(te, "</si>");
            if (!close) goto fail;
            txt = dc_collect_t_text(te, close);
            next = close + 5;
        }
        if (!txt) goto fail;
        if (n == cap) {
            cap *= 2;
            char** na = (char**)realloc(arr, sizeof(char*) * (size_t)cap);
            if (!na) { free(txt); goto fail; }
            arr = na;
        }
        arr[n++] = txt;
        p = next;
    }
    *out_count = n;
    return arr;
fail:
    for (int i = 0; i < n; i++) free(arr[i]);
    free(arr);
    return NULL;
}

static void dc_free_shared(char** arr, int n) {
    if (!arr) return;
    for (int i = 0; i < n; i++) free(arr[i]);
    free(arr);
}

// 路径归一化：去掉前导 '/' 与 './'，相对路径补 xl/ 前缀
static void dc_normalize_part_path(const char* target, char* out, size_t out_size) {
    const char* t = target;
    if (t[0] == '/') t++;
    else if (strncmp(t, "./", 2) == 0) t += 2;
    if (strncmp(t, "xl/", 3) == 0) {
        strncpy(out, t, out_size - 1);
    } else {
        snprintf(out, out_size, "xl/%s", t);
    }
    out[out_size - 1] = '\0';
}

// 从 workbook.xml + rels 解析工作表部件路径（按工作簿顺序）
// 0 成功 / DC_ZIP_UNSUPPORTED 不支持
static int dc_parse_sheets(DcZip* z, char (*paths)[256], int max_sheets, int* out_count) {
    *out_count = 0;
    DcZipEntry* wb_e = dc_zip_find(z, "xl/workbook.xml");
    DcZipEntry* rel_e = dc_zip_find(z, "xl/_rels/workbook.xml.rels");
    if (!wb_e || !rel_e) return DC_ZIP_UNSUPPORTED;
    unsigned char* wb = dc_zip_read(z, wb_e, NULL);
    unsigned char* rel = dc_zip_read(z, rel_e, NULL);
    if (!wb || !rel) { free(wb); free(rel); return DC_ZIP_UNSUPPORTED; }

    int rc = DC_ZIP_UNSUPPORTED;

    // 收集 worksheet 关系：rId -> Target
    typedef struct { char id[64]; char target[256]; } RelMap;
    RelMap maps[256];
    int map_count = 0;
    {
        const char* end = (const char*)rel + strlen((const char*)rel);
        const char* p = (const char*)rel;
        while (p < end) {
            const char* tag = strstr(p, "<Relationship");
            if (!tag || tag >= end) break;
            if (!dc_tag_at(tag, "<Relationship")) { p = tag + 13; continue; }
            int selfclose = 0;
            const char* te = dc_find_tag_end(tag, end, &selfclose);
            if (!te) goto done;
            DcAttr attrs[12];
            int na = dc_tag_attrs(tag, te, attrs, 12);
            if (na < 0) goto done;
            const DcAttr* a_id = dc_attr_find(attrs, na, "Id");
            const DcAttr* a_type = dc_attr_find(attrs, na, "Type");
            const DcAttr* a_tg = dc_attr_find(attrs, na, "Target");
            if (a_id && a_type && a_tg) {
                const char* tv = a_type->vs;
                size_t tl = (size_t)(a_type->ve - a_type->vs);
                if (tl >= 10 && strncmp(a_type->ve - 10, "/worksheet", 10) == 0) {
                    if (map_count < 256) {
                        RelMap* m = &maps[map_count++];
                        size_t il = (size_t)(a_id->ve - a_id->vs);
                        if (il > 63) il = 63;
                        memcpy(m->id, a_id->vs, il);
                        m->id[il] = '\0';
                        size_t gl = (size_t)(a_tg->ve - a_tg->vs);
                        if (gl > 255) gl = 255;
                        memcpy(m->target, a_tg->vs, gl);
                        m->target[gl] = '\0';
                        (void)tv;
                    }
                }
            }
            p = te;
        }
    }

    // 按 workbook 中 <sheet> 顺序取部件路径
    {
        const char* end = (const char*)wb + strlen((const char*)wb);
        const char* p = (const char*)wb;
        int count = 0;
        while (p < end) {
            const char* tag = strstr(p, "<sheet");
            if (!tag || tag >= end) break;
            if (!dc_tag_at(tag, "<sheet")) { p = tag + 6; continue; }
            int selfclose = 0;
            const char* te = dc_find_tag_end(tag, end, &selfclose);
            if (!te) goto done;
            DcAttr attrs[12];
            int na = dc_tag_attrs(tag, te, attrs, 12);
            if (na < 0) goto done;
            const DcAttr* a_rid = dc_attr_find(attrs, na, "r:id");
            if (!a_rid) goto done;
            char rid[64];
            size_t rl = (size_t)(a_rid->ve - a_rid->vs);
            if (rl > 63) goto done;
            memcpy(rid, a_rid->vs, rl);
            rid[rl] = '\0';
            const char* target = NULL;
            for (int i = 0; i < map_count; i++) {
                if (strcmp(maps[i].id, rid) == 0) { target = maps[i].target; break; }
            }
            if (!target) goto done;
            if (count >= max_sheets) goto done;
            dc_normalize_part_path(target, paths[count], 256);
            if (!dc_zip_find(z, paths[count])) goto done;
            count++;
            p = te;
        }
        if (count == 0) goto done;
        *out_count = count;
        rc = 0;
    }

done:
    free(wb);
    free(rel);
    return rc;
}

// 数值单元格写入 <v> 的文本形式
static void dc_format_sheet_number(double val, char* buf, size_t buf_size) {
    if (floor(val) == val && fabs(val) < 9007199254740992.0) {
        snprintf(buf, buf_size, "%.0f", val);
    } else {
        snprintf(buf, buf_size, "%.15g", val);
    }
}

// 处理一个工作表 XML：逐单元格清洗，仅替换发生变化的单元格
// 0 成功（*out_xml 为 malloc 的新 XML）/ DC_ZIP_UNSUPPORTED 不支持
static int dc_process_sheet_xml(const char* xml, char** shared, int shared_count,
                                const CellContext* ctx, int* formulas_changed,
                                char** out_xml) {
    DcBuf out = { 0, 0, 0 };
    const char* xml_end = xml + strlen(xml);
    const char* cur = xml;   // 已输出到的位置

    while (cur < xml_end) {
        // 找下一个单元格起始标签
        const char* cand = NULL;
        const char* scan = cur;
        while (scan < xml_end) {
            const char* f = strstr(scan, "<c");
            if (!f || f >= xml_end) { scan = xml_end; break; }
            if (dc_tag_at(f, "<c")) { cand = f; break; }
            scan = f + 2;
        }
        if (!cand) {
            if (dc_buf_put(&out, cur, (size_t)(xml_end - cur)) != 0) goto unsupported;
            break;
        }
        int selfclose = 0;
        const char* tag_end = dc_find_tag_end(cand, xml_end, &selfclose);
        if (!tag_end) goto unsupported;
        if (dc_buf_put(&out, cur, (size_t)(cand - cur)) != 0) goto unsupported;

        if (selfclose) { // 空单元格，原样保留
            if (dc_buf_put(&out, cand, (size_t)(tag_end - cand)) != 0) goto unsupported;
            cur = tag_end;
            continue;
        }

        const char* close = strstr(tag_end, "</c>");
        if (!close) goto unsupported;
        const char* cell_end = close + 4;
        const char* inner = tag_end;
        const char* inner_end = close;

        // 属性
        DcAttr attrs[12];
        int na = dc_tag_attrs(cand, tag_end, attrs, 12);
        if (na < 0) goto unsupported;
        const DcAttr* a_r = dc_attr_find(attrs, na, "r");
        const DcAttr* a_s = dc_attr_find(attrs, na, "s");
        const DcAttr* a_t = dc_attr_find(attrs, na, "t");
        if (!a_r) goto unsupported; // 无坐标的单元格无法安全重写
        char tbuf[16];
        tbuf[0] = 'n';
        tbuf[1] = '\0';
        if (a_t) {
            size_t tl = (size_t)(a_t->ve - a_t->vs);
            if (tl > 15) goto unsupported;
            memcpy(tbuf, a_t->vs, tl);
            tbuf[tl] = '\0';
        }

        // 公式？
        int has_f = 0;
        for (const char* q = inner; q < inner_end; q++) {
            if (dc_tag_at(q, "<f")) { has_f = 1; break; }
        }

        // 取当前值（字符串形式，与重建路径经 xlsxio 读到的口径一致）
        const char* borrowed = NULL;
        char* owned = NULL;
        if (strcmp(tbuf, "s") == 0) {
            const char* v = NULL;
            for (const char* q = inner; q < inner_end; q++) {
                if (dc_tag_at(q, "<v")) {
                    int sc2 = 0;
                    const char* vte = dc_find_tag_end(q, inner_end, &sc2);
                    if (!vte) goto unsupported;
                    const char* vc = strstr(vte, "</v>");
                    if (!vc || vc > inner_end) goto unsupported;
                    owned = dc_xml_unescape(vte, (size_t)(vc - vte));
                    v = owned;
                    break;
                }
            }
            if (!v) { borrowed = ""; }
            else {
                char* ep = NULL;
                long idx = strtol(v, &ep, 10);
                free(owned);
                owned = NULL;
                if (ep == v || idx < 0 || idx >= shared_count) goto unsupported;
                borrowed = shared[idx];
            }
        } else if (strcmp(tbuf, "inlineStr") == 0) {
            owned = dc_collect_t_text(inner, inner_end);
            borrowed = owned;
        } else {
            const char* v = NULL;
            for (const char* q = inner; q < inner_end; q++) {
                if (dc_tag_at(q, "<v")) {
                    int sc2 = 0;
                    const char* vte = dc_find_tag_end(q, inner_end, &sc2);
                    if (!vte) goto unsupported;
                    if (sc2) { v = ""; break; }
                    const char* vc = strstr(vte, "</v>");
                    if (!vc || vc > inner_end) goto unsupported;
                    owned = dc_xml_unescape(vte, (size_t)(vc - vte));
                    v = owned;
                    break;
                }
            }
            borrowed = v ? v : "";
        }

        const char* value = borrowed ? borrowed : "";

        if (value[0] == '\0') { // 空值单元格：原样保留，不计统计（与重建路径一致）
            if (dc_buf_put(&out, cand, (size_t)(cell_end - cand)) != 0) { free(owned); goto unsupported; }
            free(owned);
            cur = cell_end;
            continue;
        }

        char cleaned[MAX_CELL_STR_LEN];
        int is_num = 0;
        double num_val = 0.0;
        clean_and_stat(value, ctx, cleaned, sizeof(cleaned), &is_num, &num_val);

        if (is_num) {
            char numbuf[64];
            dc_format_sheet_number(num_val, numbuf, sizeof(numbuf));
            // <c r=".." [s=".."]><v>num</v></c>
            if (dc_buf_puts(&out, "<c ") != 0 ||
                dc_buf_put(&out, a_r->ns, (size_t)(a_r->ae - a_r->ns)) != 0) { free(owned); goto unsupported; }
            if (a_s) {
                if (dc_buf_puts(&out, " ") != 0 ||
                    dc_buf_put(&out, a_s->ns, (size_t)(a_s->ae - a_s->ns)) != 0) { free(owned); goto unsupported; }
            }
            if (dc_buf_puts(&out, "><v>") != 0 || dc_buf_puts(&out, numbuf) != 0 ||
                dc_buf_puts(&out, "</v></c>") != 0) { free(owned); goto unsupported; }
            if (has_f) *formulas_changed = 1;
        } else if (strcmp(cleaned, value) == 0) {
            // 未变化：整格原样保留（公式、样式、合并关系都不动）
            if (dc_buf_put(&out, cand, (size_t)(cell_end - cand)) != 0) { free(owned); goto unsupported; }
        } else {
            if (has_f) *formulas_changed = 1;
            if (cleaned[0] == '\0') {
                // 清洗后为空：保留坐标与样式的空单元格
                if (dc_buf_puts(&out, "<c ") != 0 ||
                    dc_buf_put(&out, a_r->ns, (size_t)(a_r->ae - a_r->ns)) != 0) { free(owned); goto unsupported; }
                if (a_s) {
                    if (dc_buf_puts(&out, " ") != 0 ||
                        dc_buf_put(&out, a_s->ns, (size_t)(a_s->ae - a_s->ns)) != 0) { free(owned); goto unsupported; }
                }
                if (dc_buf_puts(&out, "/>") != 0) { free(owned); goto unsupported; }
            } else {
                char* esc = dc_xml_escape(cleaned);
                if (!esc) { free(owned); goto unsupported; }
                size_t cl = strlen(cleaned);
                int preserve = (cleaned[0] == ' ' || cleaned[cl - 1] == ' ');
                int bad = 0;
                if (dc_buf_puts(&out, "<c ") != 0 ||
                    dc_buf_put(&out, a_r->ns, (size_t)(a_r->ae - a_r->ns)) != 0) bad = 1;
                if (!bad && a_s) {
                    if (dc_buf_puts(&out, " ") != 0 ||
                        dc_buf_put(&out, a_s->ns, (size_t)(a_s->ae - a_s->ns)) != 0) bad = 1;
                }
                if (!bad && (dc_buf_puts(&out, " t=\"inlineStr\"><is><t") != 0 ||
                             (preserve && dc_buf_puts(&out, " xml:space=\"preserve\"") != 0) ||
                             dc_buf_puts(&out, ">") != 0 ||
                             dc_buf_puts(&out, esc) != 0 ||
                             dc_buf_puts(&out, "</t></is></c>") != 0)) bad = 1;
                free(esc);
                if (bad) { free(owned); goto unsupported; }
            }
        }

        free(owned);
        cur = cell_end;
    }

    if (dc_buf_put(&out, "", 1) != 0) goto unsupported;
    *out_xml = (char*)out.d;
    return 0;

unsupported:
    dc_buf_free(&out);
    return DC_ZIP_UNSUPPORTED;
}


// 删除 XML 中提到 needle 的元素（如 calcChain 的 Override / Relationship 条目），malloc 结果
static char* dc_remove_elem_mentioning(const char* xml, const char* elem, const char* needle) {
    const char* np = strstr(xml, needle);
    if (!np) return NULL;
    const char* start = NULL;
    for (const char* q = xml; q < np; q++) {
        if (*q == '<' && dc_tag_at(q, elem)) start = q;
    }
    if (!start) return NULL;
    const char* close = strstr(np, "/>");
    if (!close) return NULL;
    const char* after = close + 2;
    DcBuf b = { 0, 0, 0 };
    if (dc_buf_put(&b, xml, (size_t)(start - xml)) != 0 ||
        dc_buf_put(&b, after, strlen(after)) != 0 ||
        dc_buf_put(&b, "", 1) != 0) {
        dc_buf_free(&b);
        return NULL;
    }
    return (char*)b.d;
}

static void dc_merge_stats(CleanStats* dst, const CleanStats* src) {
    dst->total_cells += src->total_cells;
    dst->modified_cells += src->modified_cells;
    dst->kept_cells += src->kept_cells;
    dst->numbers_found += src->numbers_found;
    dst->dates_normalized += src->dates_normalized;
}

// XLSX 原位处理主流程：统计先记本地账，成功/取消时才并入，
// 保证回退到重建路径时不会重复计数。
// 返回 0 成功 / 1 打开失败 / 2 已取消 / DC_ZIP_UNSUPPORTED 结构不支持（回退）
static int process_xlsx_inplace(const wchar_t* in_path, const wchar_t* out_path,
                                const CleanOptions* opts, ProgressCallback cb, void* user,
                                CleanStats* stats) {
    DcZip z;
    int lrc = dc_zip_load(in_path, &z);
    if (lrc == 1) return 1;
    if (lrc != 0) return DC_ZIP_UNSUPPORTED;

    if (cb) cb(user, 5, L"正在打开 Excel 文件...");

    int rc = DC_ZIP_UNSUPPORTED;
    char** shared = NULL;
    int shared_count = 0;
    char sheet_paths[MAX_SHEETS][256];
    int sheet_count = 0;
    CleanStats local;
    memset(&local, 0, sizeof(local));
    CellContext ctx = { opts, &local };
    int formulas_changed = 0;
    int cancelled = 0;

    // sharedStrings.xml 是可选部件
    DcZipEntry* ss = dc_zip_find(&z, "xl/sharedStrings.xml");
    if (ss) {
        unsigned char* ss_xml = dc_zip_read(&z, ss, NULL);
        if (!ss_xml) goto done;
        shared = dc_parse_shared_strings((const char*)ss_xml, &shared_count);
        free(ss_xml);
        if (!shared) goto done;
    }

    if (dc_parse_sheets(&z, sheet_paths, MAX_SHEETS, &sheet_count) != 0) goto done;

    for (int i = 0; i < sheet_count; i++) {
        if (g_cancel_requested) { cancelled = 1; break; }

        wchar_t status_buf[256];
        swprintf(status_buf, 256, L"正在处理工作表 (%d/%d)...", i + 1, sheet_count);
        int pct = 10 + (int)(80.0 * i / sheet_count);
        if (cb) cb(user, pct, status_buf);

        DcZipEntry* se = dc_zip_find(&z, sheet_paths[i]);
        if (!se) goto done;
        unsigned char* sheet_raw = dc_zip_read(&z, se, NULL);
        if (!sheet_raw) goto done;
        char* new_xml = NULL;
        int prc = dc_process_sheet_xml((const char*)sheet_raw, shared, shared_count,
                                       &ctx, &formulas_changed, &new_xml);
        if (prc != 0) { free(sheet_raw); goto done; }
        if (strcmp(new_xml, (const char*)sheet_raw) == 0) {
            free(new_xml); // 整表无变化：条目原样复制，字节级保留
        } else {
            se->new_data = (unsigned char*)new_xml;
            se->new_len = strlen(new_xml);
        }
        free(sheet_raw);
    }

    if (cancelled) {
        dc_merge_stats(stats, &local);
        if (cb) cb(user, 0, L"任务已取消，已丢弃该文件的部分输出。");
        rc = 2;
        goto done;
    }

    // 有公式单元格被改写：其缓存值已失效，摘除 calcChain 让 Excel 打开时重算
    if (formulas_changed) {
        DcZipEntry* cc = dc_zip_find(&z, "xl/calcChain.xml");
        if (cc) {
            cc->drop = 1;
            DcZipEntry* ct = dc_zip_find(&z, "[Content_Types].xml");
            if (ct && !ct->new_data) {
                unsigned char* ct_xml = dc_zip_read(&z, ct, NULL);
                if (ct_xml) {
                    char* patched = dc_remove_elem_mentioning((const char*)ct_xml, "<Override", "calcChain.xml");
                    free(ct_xml);
                    if (patched) {
                        ct->new_data = (unsigned char*)patched;
                        ct->new_len = strlen(patched);
                    }
                }
            }
            DcZipEntry* wr = dc_zip_find(&z, "xl/_rels/workbook.xml.rels");
            if (wr && !wr->new_data) {
                unsigned char* wr_xml = dc_zip_read(&z, wr, NULL);
                if (wr_xml) {
                    char* patched = dc_remove_elem_mentioning((const char*)wr_xml, "<Relationship", "calcChain.xml");
                    free(wr_xml);
                    if (patched) {
                        wr->new_data = (unsigned char*)patched;
                        wr->new_len = strlen(patched);
                    }
                }
            }
        }
    }

    if (cb) cb(user, 95, L"正在保存并关闭 Excel 文件...");
    if (dc_zip_write(&z, out_path) != 0) {
        DeleteFileW(out_path);
        rc = DC_ZIP_UNSUPPORTED; // 写出异常时回退重建路径重试
        goto done;
    }

    dc_merge_stats(stats, &local);
    if (cb) cb(user, 100, L"文件处理完成。");
    rc = 0;

done:
    dc_free_shared(shared, shared_count);
    dc_zip_free(&z);
    return rc;
}

// XLSX 入口：优先原位改值（保留样式/公式/合并单元格），结构不支持时回退原重建路径
static int process_xlsx(const wchar_t* in_path, const wchar_t* out_path,
                        const CleanOptions* opts, ProgressCallback cb, void* user,
                        CleanStats* stats) {
    int rc = process_xlsx_inplace(in_path, out_path, opts, cb, user, stats);
    if (rc == DC_ZIP_UNSUPPORTED) {
        rc = process_xlsx_rebuild(in_path, out_path, opts, cb, user, stats);
    }
    return rc;
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
