package com.example.excel_practice.write;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelGenerator {

    private static String FILE_NAME = "temp.xlsx";
    private String fileLocation;

    public void writeExcel() throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = createSheet(workbook);
        createHeader(workbook, sheet);
        createContents(workbook, sheet);
        createExcelFile(workbook);
        workbook.close();
    }

    private XSSFSheet createSheet(XSSFWorkbook workbook) {
        XSSFSheet sheet = workbook.createSheet("Persons");
        sheet.setColumnWidth(0, 6000);
        sheet.setColumnWidth(1, 4000);
        return sheet;
    }

    private void createHeader(XSSFWorkbook workbook, XSSFSheet sheet) {
        XSSFRow header = sheet.createRow(0);

        XSSFCellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFFont font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 16);
        font.setBold(true);
        headerStyle.setFont(font);

        XSSFCell headerCell = header.createCell(0);
        headerCell.setCellValue("Name");
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(1);
        headerCell.setCellValue("Age");
        headerCell.setCellStyle(headerStyle);
    }

    private void createContents(XSSFWorkbook workbook, XSSFSheet sheet) {
        XSSFCellStyle style = workbook.createCellStyle();
        style.setWrapText(true);

        XSSFRow row = sheet.createRow(2);
        XSSFCell cell = row.createCell(0);
        cell.setCellValue("John Smith");
        cell.setCellStyle(style);

        cell = row.createCell(1);
        cell.setCellValue(20);
        cell.setCellStyle(style);
    }

    private void createExcelFile(XSSFWorkbook workbook) throws IOException {
        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        fileLocation = path.substring(0, path.length() - 1) + FILE_NAME;

        FileOutputStream outputStream = new FileOutputStream(fileLocation);
        workbook.write(outputStream);

        outputStream.close();
    }
}
