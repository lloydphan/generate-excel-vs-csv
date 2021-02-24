package com.tuantu.excel;

import org.apache.commons.lang3.RandomStringUtils;
import org.apache.commons.lang3.time.StopWatch;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

public class Write2003ExcelHSSF {

//    1) Extract in both excel and csv format
//    2) make it 1k, 10k, 50k, 100k, 250k, 500k records
//    3) Write down for each; loading time, file size

    public static final int COLUMN_INDEX_ID = 0;
    public static final int COLUMN_INDEX_NAME = 1;
    public static final int COLUMN_INDEX_CONTACT = 2;
    public static final int COLUMN_INDEX_EMAIL = 3;
    public static final String EXCEL_EXTENSION = ".xls";

    private static CellStyle cellStyleFormatNumber = null;

    public static void main(String[] args) throws IOException {
        Long[] limitRows = {Long.valueOf(1000), Long.valueOf(10000), Long.valueOf(50000)};
        for (int i = 1; i <= 10; i++) {
            for (Long l : limitRows) {
                List<Eform> listEforms = getEforms(l, i * 10);
                extractFile(i * 10, l, listEforms);
            }
        }

    }

    public static void extractFile(long limitCols, long limit, List<Eform> listEforms) throws IOException {
        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        String fileName = "/Users/lloyd/Desktop/excel_" + limitCols + "-cols_" + limit;
        String excelPath = fileName + EXCEL_EXTENSION;
        writeExcel(listEforms, excelPath, limitCols);
        new File(excelPath).renameTo(new File(fileName + "-" + stopWatch.getTime() + "ms" + EXCEL_EXTENSION));
        System.out.println(stopWatch.getTime());
    }

    public static void writeExcel(List<Eform> listEforms, String excelPath, long limitCols) throws IOException {
        // Create Workbook
        HSSFWorkbook workbook = new HSSFWorkbook();

        // Create sheet
        HSSFSheet sheet = workbook.createSheet("Eform"); // Create sheet with sheet name

        // register the columns you wish to track and compute the column width
        //sheet.trackAllColumnsForAutoSizing();
        //sheet.setRandomAccessWindowSize(1000);

        int rowIndex = 0;
        // Write header
        writeHeader(sheet, rowIndex, limitCols);

        rowIndex++;
        for (Eform eform : listEforms) {
            // Create row
            HSSFRow row = sheet.createRow(rowIndex);
            // Write data on row
            writeEform(eform, row, limitCols);
            rowIndex++;
        }

        // Auto resize column witdth
        int numberOfColumn = sheet.getRow(rowIndex - 1).getPhysicalNumberOfCells();
        autosizeColumn(sheet, numberOfColumn);

        createOutputFile(workbook, excelPath);
        System.out.println(excelPath + " was generated completely!!!");
    }

    // Write header with format
    private static void writeHeader(HSSFSheet sheet, int rowIndex, long limitCols) {
        // create CellStyle
        CellStyle cellStyle = createStyleForHeader(sheet);

        // Create row
        HSSFRow row = sheet.createRow(rowIndex);

        // Create cells
        HSSFCell cell = row.createCell(COLUMN_INDEX_ID);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Id");

        cell = row.createCell(COLUMN_INDEX_ID);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("ID");

        cell = row.createCell(COLUMN_INDEX_NAME);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Name");

        cell = row.createCell(COLUMN_INDEX_CONTACT);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Contact");

        cell = row.createCell(COLUMN_INDEX_EMAIL);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Email");

        for (int i = 1; i <= limitCols; i++) {
            cell = row.createCell(i + 3);
            cell.setCellStyle(cellStyle);
            cell.setCellValue("Field_" + i);
        }
    }

    // Auto resize column width
    private static void autosizeColumn(HSSFSheet sheet, int lastColumn) {
        for (int columnIndex = 0; columnIndex < lastColumn; columnIndex++) {
            sheet.autoSizeColumn((short)columnIndex);
        }
    }

    // Write data
    private static void writeEform(Eform eform, HSSFRow row, long limitCols) {
        if (cellStyleFormatNumber == null) {
            // Format number
            short format = (short) BuiltinFormats.getBuiltinFormat("#,##0");
            // DataFormat df = workbook.createDataFormat();
            // short format = df.getFormat("#,##0");

            // Create CellStyle
            HSSFWorkbook workbook = row.getSheet().getWorkbook();
            cellStyleFormatNumber = workbook.createCellStyle();
            cellStyleFormatNumber.setDataFormat(format);
        }

        HSSFCell cell = row.createCell(COLUMN_INDEX_ID);
        cell.setCellValue(eform.getId() + "");

        cell = row.createCell(COLUMN_INDEX_NAME);
        cell.setCellValue(eform.getName());

        cell = row.createCell(COLUMN_INDEX_CONTACT);
        cell.setCellValue(eform.getContact());

        cell = row.createCell(COLUMN_INDEX_EMAIL);
        cell.setCellValue(eform.getEmail());

        for (int i = 1; i <= limitCols; i++) {
            cell = row.createCell(3 + i);
            cell.setCellValue(eform.getListCols().get(i - 1));
        }
    }

    // Create CellStyle for header
    private static CellStyle createStyleForHeader(Sheet sheet) {
        // Create font
        Font font = sheet.getWorkbook().createFont();
        font.setFontName("Times New Roman");
        font.setBold(true);
        font.setFontHeightInPoints((short) 14); // font size
        font.setColor(IndexedColors.WHITE.getIndex()); // text color

        // Create CellStyle
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setFont(font);
        cellStyle.setFillForegroundColor(IndexedColors.BROWN.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        return cellStyle;
    }

    private static List<Eform> getEforms(long limits, long limitCols) {
        List<Eform> listEforms = new ArrayList<>();
        Eform eform;
        for (int i = 1; i <= limits; i++) {
            String generatedString = RandomStringUtils.randomAlphabetic(10);
            eform = new Eform(BigInteger.valueOf(i), "John Doe " + i, "01112" + i, "johndoe_" + i + "@gmail.com", generatedString + "_", limitCols);
            listEforms.add(eform);
        }
        return listEforms;
    }

    // Create output file
    private static void createOutputFile(HSSFWorkbook workbook, String excelFilePath) throws IOException {
        try (OutputStream os = new FileOutputStream(excelFilePath)) {
            workbook.write(os);
        }
    }
}
