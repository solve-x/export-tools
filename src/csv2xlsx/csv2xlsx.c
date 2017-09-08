#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <errno.h>
#include "csv.h"
#include "xlsxwriter.h"

// To build this program on a Ubuntu machine:
// $ apt-get install libcsv-dev
// $ git clone https://github.com/jmcnamara/libxlsxwriter.git
// $ cs libxlsxwriter/
// $ make
// $ sudo make install
//
// Compile and link:
// cc csv2xlsx.c -O2 -o csv2xlsx -lcsv -lxlsxwriter
//
// Compile and link statically + strip:
// cc csv2xlsx.c -O2 -s -o csv2xlsx -l:libcsv.a -l:libxlsxwriter.a -l:libz.a

lxw_row_t currentRow = 0;
lxw_col_t currentColumn = 0;
lxw_format *formatBold = NULL;

static int is_space(unsigned char c)
{
    return (c == CSV_SPACE);
}

/**
 * Returns true when a string is considered a "number":
 * - must not be empty
 * - can start with '-'
 * - the rest must be only digits (0-9) and optionally one decimal-point ('.')
 */
int isNumeric(char *s)
{
    int dotCount = 0;
    int digitCount = 0;

    if (*s == 0) {
        return 0;
    }

    while (*s) {
        if (*s == '.') {
            dotCount++;
            if (dotCount > 1) {
                return 0;
            }
        } else if (*s == '-') {
            if (digitCount > 0) {
                return 0;
            }
        } else if (isdigit(*s) == 0) {
            return 0;
        }

        digitCount++;
        s++;
    }

    return 1;
}

/**
 * This callback is called for each cell in a CSV.
 */
void cb1(void *s, size_t len, void *data)
{
    lxw_worksheet *worksheet = (lxw_worksheet *)data;
    char *content = (char *)s;

    if (isNumeric(content)) {
        double numberContent = strtod(content, NULL);
        worksheet_write_number(worksheet, currentRow, currentColumn, numberContent, (currentRow == 0) ? formatBold : NULL);
    } else {
        worksheet_write_string(worksheet, currentRow, currentColumn, content, (currentRow == 0) ? formatBold : NULL);
    }

    currentColumn++;
}

void cb2(int c, void *data)
{
    currentColumn = 0;
    currentRow++;
}

int main(int argc, char *argv[])
{
    if (argc < 3) {
        fprintf(stderr, "Usage: csv2xlsx in.csv out.xlsx\n");
        exit(EXIT_FAILURE);
    }

    struct csv_parser parser;
    unsigned char parser_options = CSV_APPEND_NULL;
    if (csv_init(&parser, parser_options) != 0) {
        fprintf(stderr, "Failed to initialize csv parser\n");
        exit(EXIT_FAILURE);
    }

    csv_set_delim(&parser, CSV_TAB);
    csv_set_space_func(&parser, is_space);

    FILE *fp = fopen(argv[1], "rb");
    if (!fp) {
        fprintf(stderr, "Failed to open %s: %s\n", argv[1], strerror(errno));
        exit(EXIT_FAILURE);
    }

    lxw_workbook_options xlsxwriter_options = {
        .constant_memory = LXW_TRUE,
        .tmpdir = NULL
    };

    lxw_workbook  *workbook  = workbook_new_opt(argv[2], &xlsxwriter_options);
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);

    formatBold = workbook_add_format(workbook);
    format_set_bold(formatBold);

    char buffer[1024];
    size_t bytes_read;
    while ((bytes_read = fread(buffer, 1, 1024, fp)) > 0) {
        if (csv_parse(&parser, buffer, bytes_read, cb1, cb2, worksheet) != bytes_read) {
            fprintf(stderr, "Error while parsing file: %s\n", csv_strerror(csv_error(&parser)));
        }
    }

    csv_fini(&parser, cb1, cb2, worksheet);

    if (ferror(fp)) {
        fprintf(stderr, "Error while reading file %s\n", argv[1]);
    }

    fclose(fp);

    workbook_close(workbook);
    csv_free(&parser);
    exit(EXIT_SUCCESS);
}
