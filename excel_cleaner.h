#ifndef EXCEL_CLEANER_H
#define EXCEL_CLEANER_H

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>

#ifdef __cplusplus
extern "C" {
#endif

// Cleaning modes
typedef enum {
    CLEAN_MODE_DEFAULT = 0, // Extract numbers, join with comma
    CLEAN_MODE_ADD = 1,     // Sum of extracted numbers
    CLEAN_MODE_MULTIPLY = 2 // Product of extracted numbers
} CleanMode;

// Progress callback function type
typedef void (*ProgressCallback)(int percentage, const wchar_t* status_msg);

// Main Excel processing function written in pure C
// Returns 0 on success, non-zero on failure
int process_excel_c(const wchar_t* input_path_w, CleanMode mode, wchar_t* output_path_w, ProgressCallback cb);

#ifdef __cplusplus
}
#endif

#endif // EXCEL_CLEANER_H
