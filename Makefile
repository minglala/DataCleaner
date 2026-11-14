# 自动识别 MSYS2 UCRT64 / MinGW GCC 编译器与资源编译器
CC ?= gcc
WINDRES ?= windres

CFLAGS = -O2 -Wall -DBUILD_XLSXIO_STATIC -municode -mwindows

# 优先使用 MSYS2 UCRT64 系统安装的 C 语言库 (pacman 安装)
STATIC_LIBS = -lxlsxwriter \
              -l:libxlsxio_read.a \
              -l:libzip.a \
              -l:libexpat.a \
              -l:libz.a \
              -l:libbz2.a \
              -l:liblzma.a \
              -l:libzstd.a

SYS_LIBS = -lcomctl32 -lcomdlg32 -lgdi32 -lshell32 -luser32 -lbcrypt -lole32 -ladvapi32

TARGET = DataCleaner.exe
OBJS = main.o excel_cleaner.o res.o

all: $(TARGET)

res.o: app.rc app.manifest
	$(WINDRES) -i app.rc -o res.o

main.o: main.c excel_cleaner.h
	$(CC) $(CFLAGS) -c main.c -o main.o

excel_cleaner.o: excel_cleaner.c excel_cleaner.h
	$(CC) $(CFLAGS) -c excel_cleaner.c -o excel_cleaner.o

$(TARGET): $(OBJS)
	$(CC) $(CFLAGS) $(OBJS) -o $(TARGET) -static-libgcc $(STATIC_LIBS) $(SYS_LIBS)

# ---- 测试 ----
# 控制台程序（不带 -mwindows），直接包含引擎源码以覆盖 static 函数
TEST_CFLAGS = -O1 -g -Wall -DBUILD_XLSXIO_STATIC
TEST_TARGET = tests/test_p0.exe

$(TEST_TARGET): tests/test_p0.c excel_cleaner.c excel_cleaner.h
	$(CC) $(TEST_CFLAGS) tests/test_p0.c -o $(TEST_TARGET) $(STATIC_LIBS) \
	      -lcomctl32 -lshlwapi -lshell32 -lole32 -ladvapi32 -lbcrypt

test: $(TEST_TARGET)
	./$(TEST_TARGET)

clean:
	rm -f *.o $(TARGET) $(TEST_TARGET)

.PHONY: all clean test
