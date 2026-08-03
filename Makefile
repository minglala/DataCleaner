CC = C:/msys64/ucrt64/bin/gcc.exe
WINDRES = C:/msys64/ucrt64/bin/windres.exe
CFLAGS = -O2 -Wall -municode -mwindows
LIBS = -lxlsxio_read -lxlsxwriter -lexpat -lz -lcomctl32 -lcomdlg32 -lgdi32 -lshell32 -luser32

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
	$(CC) $(CFLAGS) $(OBJS) -o $(TARGET) $(LIBS)

clean:
	rm -f *.o $(TARGET)

.PHONY: all clean
