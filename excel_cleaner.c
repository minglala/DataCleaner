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

static char* wchar_to_utf8(const wchar_t* wstr) {
    if (!wstr) return NULL;
    int len = WideCharToMultiByte(CP_UTF8, 0, wstr, -1, NULL, 0, NULL, NULL);
    if (len <= 0) return NULL;
    char* str = (char*)malloc(len);
    WideCharToMultiByte(CP_UTF8, 0, wstr, -1, str, len, NULL, NULL);
    return str;
}

static int extract_numbers(const char* text, double* nums, int max_nums) {
    if (!text || !*text) return 0;
    int count = 0;
    const char* p = text;

    while (*p && count < max_nums) {
        if (*p == '-' && (isdigit((unsigned char)*(p + 1)) || (*(p + 1) == '.' && isdigit((unsigned char)*(p + 2))))) {
            char* endptr = NULL;
            double val = strtod(p, &endptr);
            if (endptr > p) {
                nums[count++] = val;
                p = endptr;
                continue;
            }
        } else if (isdigit((unsigned char)*p) || (*p == '.' && isdigit((unsigned char)*(p + 1)))) {
            char* endptr = NULL;
            double val = strtod(p, &endptr);
            if (endptr > p) {
                nums[count++] = val;
                p = endptr;
                continue;
            }
        }
        p++;
    }
    return count;
}

static void format_double_clean(double val, char* buf, size_t buf_size) {
    if (floor(val) == val && fabs(val) < 1e15) {
        snprintf(buf, buf_size, "%.0f", val);
    } else {
        snprintf(buf, buf_size, "%.6g", val);
    }
}

static void clean_cell_value_c(const char* in_str, CleanMode mode, char* out_str, size_t out_max, int* is_number, double* out_num) {
    *is_number = 0;
    *out_num = 0.0;
    out_str[0] = '\0';

    if (!in_str || !*in_str) return;

    double nums[MAX_NUMS_PER_CELL];
    int count = extract_numbers(in_str, nums, MAX_NUMS_PER_CELL);

    if (count == 0) {
        strncpy(out_str, in_str, out_max - 1);
        out_str[out_max - 1] = '\0';
        return;
    }

    if (mode == CLEAN_MODE_DEFAULT) {
        char temp[64];
        out_str[0] = '\0';
        for (int i = 0; i < count; i++) {
            format_double_clean(nums[i], temp, sizeof(temp));
            if (i > 0) strncat(out_str, ", ", out_max - strlen(out_str) - 1);
            strncat(out_str, temp, out_max - strlen(out_str) - 1);
        }
    } else if (mode == CLEAN_MODE_ADD) {
        double sum = 0.0;
        for (int i = 0; i < count; i++) sum += nums[i];
        *is_number = 1;
        *out_num = sum;
    } else if (mode == CLEAN_MODE_MULTIPLY) {
        double prod = 1.0;
        for (int i = 0; i < count; i++) prod *= nums[i];
        *is_number = 1;
        *out_num = prod;
    }
}

int process_excel_c(const wchar_t* input_path_w, CleanMode mode, wchar_t* output_path_w, ProgressCallback cb) {
    if (!input_path_w || !*input_path_w) return 1;

    wchar_t out_path_w[MAX_PATH];
    if (!output_path_w || !*output_path_w) {
        wcscpy(out_path_w, input_path_w);
        wchar_t* dot = wcsrchr(out_path_w, L'.');
        if (dot) {
            wchar_t ext[32];
            wcscpy(ext, dot);
            wcscpy(dot, L"_Cleaned");
            wcscat(out_path_w, ext);
        } else {
            wcscat(out_path_w, L"_Cleaned.xlsx");
        }
    } else {
        wcscpy(out_path_w, output_path_w);
    }

    char* in_utf8 = wchar_to_utf8(input_path_w);
    char* out_utf8 = wchar_to_utf8(out_path_w);

    if (cb) cb(10, L"正在打开 Excel 文件...");

    xlsxioreader read_handle = xlsxioread_open(in_utf8);
    if (!read_handle) {
        if (cb) cb(0, L"错误：无法打开输入 Excel 文件！");
        free(in_utf8);
        free(out_utf8);
        return 1;
    }

    lxw_workbook* write_handle = workbook_new(out_utf8);
    if (!write_handle) {
        xlsxioread_close(read_handle);
        if (cb) cb(0, L"错误：无法创建输出 Excel 文件！");
        free(in_utf8);
        free(out_utf8);
        return 1;
    }

    char* sheets[MAX_SHEETS];
    int sheet_count = 0;

    xlsxioreadersheetlist sheetlist = xlsxioread_sheetlist_open(read_handle);
    const char* name = NULL;
    while (sheetlist && (name = xlsxioread_sheetlist_next(sheetlist)) != NULL && sheet_count < MAX_SHEETS) {
        sheets[sheet_count] = strdup(name);
        sheet_count++;
    }
    if (sheetlist) xlsxioread_sheetlist_close(sheetlist);

    for (int i = 0; i < sheet_count; i++) {
        wchar_t status_buf[256];
        swprintf(status_buf, 256, L"正在处理工作表 (%d/%d)...", i + 1, sheet_count);
        int pct = 10 + (int)(80.0 * i / sheet_count);
        if (cb) cb(pct, status_buf);

        lxw_worksheet* out_ws = workbook_add_worksheet(write_handle, sheets[i]);
        xlsxioreadersheet read_sheet = xlsxioread_sheet_open(read_handle, sheets[i], XLSXIOREAD_SKIP_NONE);

        if (read_sheet) {
            while (xlsxioread_sheet_next_row(read_sheet)) {
                size_t r = xlsxioread_sheet_last_row_index(read_sheet) - 1;
                char* cell_val = NULL;
                while (xlsxioread_sheet_next_cell_string(read_sheet, &cell_val)) {
                    size_t c = xlsxioread_sheet_last_column_index(read_sheet) - 1;
                    if (cell_val && *cell_val) {
                        char cleaned_str[MAX_CELL_STR_LEN];
                        int is_num = 0;
                        double num_val = 0.0;

                        clean_cell_value_c(cell_val, mode, cleaned_str, sizeof(cleaned_str), &is_num, &num_val);

                        if (is_num) {
                            worksheet_write_number(out_ws, (lxw_row_t)r, (lxw_col_t)c, num_val, NULL);
                        } else if (cleaned_str[0] != '\0') {
                            worksheet_write_string(out_ws, (lxw_row_t)r, (lxw_col_t)c, cleaned_str, NULL);
                        }
                        xlsxioread_free(cell_val);
                    }
                }
            }
            xlsxioread_sheet_close(read_sheet);
        }
        free(sheets[i]);
    }

    if (cb) cb(95, L"正在保存并关闭 Excel 文件...");

    printf("Closing xlsxioreader...\n"); fflush(stdout);
    xlsxioread_close(read_handle);

    printf("Closing workbook...\n"); fflush(stdout);
    lxw_error err = workbook_close(write_handle);
    printf("Workbook close returned: %d\n", err); fflush(stdout);

    if (output_path_w) {
        wcscpy(output_path_w, out_path_w);
    }

    free(in_utf8);
    free(out_utf8);

    if (cb) cb(100, L"清洗完毕！已成功导出 Excel 文件。");

    return 0;
}
