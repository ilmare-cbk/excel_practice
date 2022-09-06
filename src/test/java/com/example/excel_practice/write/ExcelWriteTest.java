package com.example.excel_practice.write;

import org.junit.Test;

import java.io.IOException;

public class ExcelWriteTest {

    @Test
    public void generateExcelFile() throws IOException {
        ExcelGenerator excelGenerator = new ExcelGenerator();
        excelGenerator.writeExcel();
    }
}
