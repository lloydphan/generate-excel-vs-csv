package com.tuantu.excel;


import com.opencsv.CSVWriter;
import org.apache.commons.lang3.RandomStringUtils;
import org.apache.commons.lang3.time.StopWatch;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.Writer;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

public class WriteCSV {

    //    1) Extract in both excel and csv format
    //    2) make it 1k, 10k, 50k, 100k, 250k, 500k records
    //    3) Write down for each; loading time, file size

    public static final String CSV_EXTENSION = ".csv";

    public static void main(String[] args) throws IOException {
        Long[] limitRows = {Long.valueOf(1000), Long.valueOf(10000), Long.valueOf(50000),
                Long.valueOf(100000), Long.valueOf(200000), Long.valueOf(500000)};
        for (int i = 1; i <= 10; i++) {
            for (Long l : limitRows) {
                List<Eform> listEforms = getEforms(l, i * 10);
                String[] headerRecords = headerRecords(i * 10);
                extractCsv(i * 10, l, listEforms);
            }
        }
    }

    public static void extractCsv(long limitCols, long limit, List<Eform> listEforms) throws IOException {
        String fileName = "/Users/lloyd/Desktop/csv_" + limitCols + "-cols_" + limit;
        String csvPath = fileName + CSV_EXTENSION;
        try (Writer writer = new FileWriter(csvPath);
             CSVWriter csvWriter = new CSVWriter(writer,
                     CSVWriter.DEFAULT_SEPARATOR,
                     CSVWriter.NO_QUOTE_CHARACTER,
                     CSVWriter.DEFAULT_ESCAPE_CHARACTER,
                     CSVWriter.DEFAULT_LINE_END);) {
            StopWatch stopWatch = new StopWatch();
            stopWatch.start();
            csvWriter.writeNext(headerRecords(limitCols));
            for (Eform eform : listEforms) {
                csvWriter.writeNext(dataRecords(eform));
            }
            new File(csvPath).renameTo(new File(fileName + "-" + stopWatch.getTime() + "ms" + CSV_EXTENSION));
            System.out.println("Generated Time:" + stopWatch.getTime());
        }
    }

    public static String[] headerRecords(long limitCols) {
        List<String> listHeader = new ArrayList<>();
        listHeader.add("ID");
        listHeader.add("Name");
        listHeader.add("Contact");
        listHeader.add("Email");
        for (int i = 1; i <= limitCols; i++) {
            listHeader.add("Field_" + i);
        }
        return listHeader.toArray(new String[listHeader.size()]);
    }

    public static String[] dataRecords(Eform eform) {
        List<String> listData = new ArrayList<>();
        listData.add(eform.getId() + "");
        listData.add(eform.getName());
        listData.add(eform.getContact());
        listData.add(eform.getEmail());
        for (String col : eform.getListCols()) {
            listData.add(col);
        }
        return listData.toArray(new String[listData.size()]);
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

}
