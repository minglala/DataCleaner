// P0 回归测试：直接包含引擎源码以访问 static 函数，用 libxlsxwriter 现场生成样本。
// 构建运行：make test
#include "../excel_cleaner.c"

static int g_fail = 0;
static int g_pass = 0;
#define CHECK(cond, what) do { \
    if (cond) { g_pass++; printf("  [PASS] %s\n", what); } \
    else { g_fail++; printf("  [FAIL] %s  (%s:%d)\n", what, __FILE__, __LINE__); } \
} while (0)

// ---------- 进度回调：收集日志用于断言 ----------
static wchar_t g_log[1 << 16];
static void test_cb(void* user, int pct, const wchar_t* msg) {
    (void)user; (void)pct;
    if (!msg) return;
    size_t n = wcslen(g_log);
    if (n + wcslen(msg) + 2 < sizeof(g_log) / sizeof(g_log[0])) {
        wcscat(g_log, msg);
        wcscat(g_log, L"\n");
    }
}

static wchar_t g_dir[MAX_PATH];

static void path_in_dir(const wchar_t* name, wchar_t* out) {
    swprintf(out, DC_PATH, L"%ls\\%ls", g_dir, name);
}

static int run_one(const wchar_t* in, CleanMode mode, FileJob* job) {
    memset(job, 0, sizeof(*job));
    wcscpy(job->input, in);
    CleanOptions o;
    memset(&o, 0, sizeof(o));
    o.mode = mode;
    wcscpy(o.separator, L", ");
    g_log[0] = 0;
    g_cancel_requested = 0;
    run_clean_jobs(job, 1, &o, test_cb, NULL, NULL);
    return job->result;
}

// 读取 xlsx 中一个部件（malloc，调用方 free）；失败返回 NULL
static char* read_part(const wchar_t* xlsx, const char* part) {
    DcZip z;
    if (dc_zip_load(xlsx, &z, NULL) != 0) return NULL;
    DcZipEntry* e = dc_zip_find(&z, part);
    char* s = e ? (char*)dc_zip_read(&z, e, NULL) : NULL;
    dc_zip_free(&z);
    return s;
}

static int count_occ(const char* s, const char* needle) {
    int n = 0;
    for (const char* p = strstr(s, needle); p; p = strstr(p + 1, needle)) n++;
    return n;
}

static int file_exists(const wchar_t* p) {
    DWORD a = GetFileAttributesW(p);
    return a != INVALID_FILE_ATTRIBUTES;
}

static char* utf8(const wchar_t* w) { return wchar_to_utf8(w); }

// ============================================================
// P0-1：公式单元格保留；数值未变化不重写；依赖单元格改动后设置打开重算
// ============================================================
static void test_formula_preserved(void) {
    printf("P0-1 formula cells preserved\n");
    wchar_t in[DC_PATH], out[DC_PATH];
    path_in_dir(L"formula.xlsx", in);
    path_in_dir(L"formula_Cleaned.xlsx", out);

    char* u = utf8(in);
    lxw_workbook* wb = workbook_new(u);
    lxw_worksheet* ws = workbook_add_worksheet(wb, NULL);
    worksheet_write_number(ws, 0, 0, 5, NULL);                       // A1 = 5（数值，结果不变）
    worksheet_write_string(ws, 1, 0, "3 apples 4", NULL);            // A2 → 求和 7
    worksheet_write_formula_num(ws, 2, 0, "=A1*2", NULL, 10);        // A3 公式，缓存数值
    worksheet_write_formula_num(ws, 3, 0, "=SUM(A1:A2)", NULL, 5);   // A4 公式
    workbook_close(wb);
    free(u);

    FileJob job;
    int rc = run_one(in, CLEAN_MODE_ADD, &job);
    CHECK(rc == 0, "job succeeds");
    CHECK(job.degraded == 0, "in-place mode used (not degraded)");

    char* sheet = read_part(out, "xl/worksheets/sheet1.xml");
    CHECK(sheet != NULL, "output sheet readable");
    if (sheet) {
        CHECK(strstr(sheet, "<f>A1*2</f>") != NULL, "formula A3 kept");
        CHECK(strstr(sheet, "<f>SUM(A1:A2)</f>") != NULL, "formula A4 kept");
        CHECK(strstr(sheet, "<c r=\"A1\"><v>5</v></c>") != NULL, "unchanged number A1 kept verbatim");
        CHECK(strstr(sheet, "<c r=\"A2\"><v>7</v></c>") != NULL, "A2 cleaned to 7");
        free(sheet);
    }
    char* wbx = read_part(out, "xl/workbook.xml");
    CHECK(wbx && strstr(wbx, "fullCalcOnLoad=\"1\"") != NULL, "workbook marked fullCalcOnLoad");
    free(wbx);
    CHECK(job.stats.formula_cells == 2, "formula_cells stat == 2");
}

// 无任何改动时：整表字节保留，不触碰 workbook.xml
static void test_no_change_bytes(void) {
    printf("P0-1 no-op file keeps parts byte-identical\n");
    wchar_t in[DC_PATH], out[DC_PATH];
    path_in_dir(L"noop.xlsx", in);
    path_in_dir(L"noop_Cleaned.xlsx", out);
    char* u = utf8(in);
    lxw_workbook* wb = workbook_new(u);
    lxw_worksheet* ws = workbook_add_worksheet(wb, NULL);
    worksheet_write_number(ws, 0, 0, 42, NULL);
    worksheet_write_formula_num(ws, 1, 0, "=A1+1", NULL, 43);
    workbook_close(wb);
    free(u);

    FileJob job;
    CHECK(run_one(in, CLEAN_MODE_ADD, &job) == 0, "job succeeds");
    char* a = read_part(in, "xl/worksheets/sheet1.xml");
    char* b = read_part(out, "xl/worksheets/sheet1.xml");
    CHECK(a && b && strcmp(a, b) == 0, "sheet1.xml identical");
    free(a); free(b);
    a = read_part(in, "xl/workbook.xml");
    b = read_part(out, "xl/workbook.xml");
    CHECK(a && b && strcmp(a, b) == 0, "workbook.xml untouched when nothing changed");
    free(a); free(b);
}

// ============================================================
// P0-2：超过 64 个工作表不丢表
// ============================================================
static void test_many_sheets(void) {
    printf("P0-2 >64 sheets not truncated\n");
    wchar_t in[DC_PATH], out[DC_PATH];
    path_in_dir(L"many.xlsx", in);
    path_in_dir(L"many_Cleaned.xlsx", out);
    const int N = 70;
    char* u = utf8(in);
    lxw_workbook* wb = workbook_new(u);
    for (int i = 0; i < N; i++) {
        lxw_worksheet* ws = workbook_add_worksheet(wb, NULL);
        worksheet_write_string(ws, 0, 0, "a1b2", NULL);
    }
    workbook_close(wb);
    free(u);

    FileJob job;
    CHECK(run_one(in, CLEAN_MODE_ADD, &job) == 0, "job succeeds");
    CHECK(job.degraded == 0, "in-place mode used");
    CHECK(job.stats.modified_cells == N, "all 70 sheets cleaned (modified_cells == 70)");
    char* last = read_part(out, "xl/worksheets/sheet70.xml");
    CHECK(last && strstr(last, "<v>3</v>") != NULL, "sheet70 cleaned to 3");
    free(last);
    char* wbx = read_part(out, "xl/workbook.xml");
    CHECK(wbx && count_occ(wbx, "<sheet ") == N, "workbook still lists 70 sheets");
    free(wbx);
}

// ============================================================
// P0-3：降级透明 —— 原位不支持时明确告知并标记 degraded
// ============================================================
static void test_degraded_reported(void) {
    printf("P0-3 degraded fallback is reported\n");
    wchar_t in[DC_PATH], out[DC_PATH];
    path_in_dir(L"degraded.xlsx", in);
    path_in_dir(L"degraded_Cleaned.xlsx", out);
    char* u = utf8(in);
    lxw_workbook* wb = workbook_new(u);
    lxw_worksheet* ws = workbook_add_worksheet(wb, NULL);
    worksheet_write_string(ws, 0, 0, "x1y2", NULL);
    workbook_close(wb);
    free(u);

    // 改造样本：单元格去掉 r 属性（原位路径无法安全重写 → 必须走兼容模式）
    DcZip z;
    int ok = dc_zip_load(in, &z, NULL) == 0;
    DcZipEntry* e = ok ? dc_zip_find(&z, "xl/worksheets/sheet1.xml") : NULL;
    char* xml = e ? (char*)dc_zip_read(&z, e, NULL) : NULL;
    if (xml) {
        char* r = strstr(xml, " r=\"A1\"");
        if (r) memmove(r, r + 7, strlen(r + 7) + 1);
        e->new_data = (unsigned char*)xml;
        e->new_len = strlen(xml);
    }
    wchar_t mod[DC_PATH];
    path_in_dir(L"degraded2.xlsx", mod);
    ok = ok && xml && dc_zip_write(&z, mod) == 0;
    dc_zip_free(&z);
    CHECK(ok, "crafted sample written");
    MoveFileExW(mod, in, MOVEFILE_REPLACE_EXISTING);

    FileJob job;
    int rc = run_one(in, CLEAN_MODE_ADD, &job);
    CHECK(rc == 0, "job succeeds via rebuild");
    CHECK(job.degraded == 1, "job.degraded flagged");
    CHECK(wcsstr(g_log, L"兼容模式") != NULL, "log announces compatibility mode");
    CHECK(wcsstr(g_log, L"无法安全改写") != NULL, "log gives the reason");
    CHECK(file_exists(out), "output produced");
}

// ============================================================
// P0-3 / P0-4：写出失败报失败、不回退；原子替换
// ============================================================
static void make_simple_xlsx(const wchar_t* path) {
    char* u = utf8(path);
    lxw_workbook* wb = workbook_new(u);
    lxw_worksheet* ws = workbook_add_worksheet(wb, NULL);
    worksheet_write_string(ws, 0, 0, "1 and 2", NULL);
    workbook_close(wb);
    free(u);
}

static void write_text(const wchar_t* p, const char* s) {
    FILE* f = _wfopen(p, L"wb");
    if (f) { fputs(s, f); fclose(f); }
}
static int text_equals(const wchar_t* p, const char* s) {
    char buf[64] = {0};
    FILE* f = _wfopen(p, L"rb");
    if (!f) return 0;
    size_t n = fread(buf, 1, sizeof(buf) - 1, f);
    fclose(f);
    buf[n] = 0;
    return strcmp(buf, s) == 0;
}

static void test_write_failure_not_degraded(void) {
    printf("P0-3 write failure => failed, no lossy retry\n");
    wchar_t in[DC_PATH], tmpdir[DC_PATH], out[DC_PATH];
    path_in_dir(L"wfail.xlsx", in);
    path_in_dir(L"wfail_Cleaned.xlsx", out);
    path_in_dir(L"wfail_Cleaned.xlsx.dctmp", tmpdir);
    make_simple_xlsx(in);
    CreateDirectoryW(tmpdir, NULL); // 临时文件路径被目录占用 → 写出必然失败

    FileJob job;
    int rc = run_one(in, CLEAN_MODE_ADD, &job);
    CHECK(rc == 1, "job reported as failed");
    CHECK(job.degraded == 0, "did NOT fall back to rebuild");
    CHECK(wcsstr(g_log, L"兼容模式") == NULL, "no compatibility-mode message");
    CHECK(!file_exists(out), "no output file created");
    RemoveDirectoryW(tmpdir);
}

static void test_atomic_replace(void) {
    printf("P0-4 atomic output\n");
    wchar_t in[DC_PATH], out[DC_PATH], tmp[DC_PATH];
    path_in_dir(L"atomic.xlsx", in);
    path_in_dir(L"atomic_Cleaned.xlsx", out);
    path_in_dir(L"atomic_Cleaned.xlsx.dctmp", tmp);
    make_simple_xlsx(in);

    // 1) 目标被独占打开 → 替换失败：旧文件保持完好、临时文件被清理
    write_text(out, "OLD");
    HANDLE h = CreateFileW(out, GENERIC_READ, 0, NULL, OPEN_EXISTING, 0, NULL);
    FileJob job;
    int rc = run_one(in, CLEAN_MODE_ADD, &job);
    CloseHandle(h);
    CHECK(rc == 1, "locked target => failed");
    CHECK(text_equals(out, "OLD"), "existing output untouched");
    CHECK(!file_exists(tmp), "temp file cleaned up");

    // 2) 正常：替换旧文件，结果完整，无临时文件残留
    rc = run_one(in, CLEAN_MODE_ADD, &job);
    CHECK(rc == 0, "normal run succeeds");
    char* s = read_part(out, "xl/worksheets/sheet1.xml");
    CHECK(s && strstr(s, "<v>3</v>") != NULL, "output replaced with valid xlsx");
    free(s);
    CHECK(!file_exists(tmp), "no temp file left");

    // 3) 取消：不产生输出、不留临时文件
    DeleteFileW(out);
    memset(&job, 0, sizeof(job));
    wcscpy(job.input, in);
    CleanOptions o; memset(&o, 0, sizeof(o)); o.mode = CLEAN_MODE_ADD;
    g_cancel_requested = 1;
    run_clean_jobs(&job, 1, &o, test_cb, NULL, NULL);
    g_cancel_requested = 0;
    CHECK(job.result == 3, "cancelled before start => skipped");
    CHECK(!file_exists(out) && !file_exists(tmp), "no output / temp after cancel");
}

static void test_csv_roundtrip(void) {
    printf("CSV path through temp + rename\n");
    wchar_t in[DC_PATH], out[DC_PATH], tmp[DC_PATH];
    path_in_dir(L"data.csv", in);
    path_in_dir(L"data_Cleaned.csv", out);
    path_in_dir(L"data_Cleaned.csv.dctmp", tmp);
    write_text(in, "a,b\n1 and 2,\"x 3, y 4\"\n");
    FileJob job;
    CHECK(run_one(in, CLEAN_MODE_ADD, &job) == 0, "csv job succeeds");
    FILE* f = _wfopen(out, L"rb");
    char buf[256] = {0};
    if (f) { fread(buf, 1, sizeof(buf) - 1, f); fclose(f); }
    CHECK(strcmp(buf, "\xEF\xBB\xBF" "a,b\n3,7\n") == 0, "csv content correct");
    CHECK(!file_exists(tmp), "no temp file left");
}

int main(void) {
    wchar_t base[MAX_PATH];
    GetTempPathW(MAX_PATH, base);
    swprintf(g_dir, MAX_PATH, L"%lsdc_p0_test_%lu", base, (unsigned long)GetCurrentProcessId());
    CreateDirectoryW(g_dir, NULL);

    test_formula_preserved();
    test_no_change_bytes();
    test_many_sheets();
    test_degraded_reported();
    test_write_failure_not_degraded();
    test_atomic_replace();
    test_csv_roundtrip();

    printf("\n%d passed, %d failed\n", g_pass, g_fail);
    return g_fail ? 1 : 0;
}
