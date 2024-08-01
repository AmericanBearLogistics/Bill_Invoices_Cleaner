package com.example;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Customers {

    public static void main(String[] args) {
        String excelFilePath = "C:\\Users\\CalvinYuen\\Downloads\\NUMBER_FIXER_7_18_2024.xlsx";
        String outputFilePath = "C:\\Users\\CalvinYuen\\Downloads\\NUMBER_FIXER_7_18_2024_UPDATED.xlsx";
        
        Map<String, String> dataMap = readExcelFile(excelFilePath);

        updateExcelFile(excelFilePath, outputFilePath, dataMap);
    }

    public static Map<String, String> readExcelFile(String filePath) {
        Map<String, String> map = new HashMap<>();

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell keyCell = row.getCell(0);
                Cell valueCell = row.getCell(1);

                if (keyCell != null && valueCell != null) {
                    String key = keyCell.toString();
                    String value = valueCell.toString();
                    map.put(key, value);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return map;
    }

    public static void updateExcelFile(String inputFilePath, String outputFilePath, Map<String, String> dataMap) {
        try (FileInputStream fis = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell lookupCell = row.getCell(2); // Third column
                Cell outputCell = row.createCell(3); // Fourth column

                if (lookupCell != null) {
                    String lookupKey = lookupCell.toString();
                    if (dataMap.containsKey(lookupKey)) {
                        String value = dataMap.get(lookupKey);
                        outputCell.setCellValue(value);
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                workbook.write(fos);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
