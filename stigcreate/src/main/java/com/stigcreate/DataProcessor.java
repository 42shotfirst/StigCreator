package com.stigcreate;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataProcessor {

    public void processData() {
        String preloadedPoliciesFile = "/Users/morganreed/Documents/Windows11andWindowsServer2019PolicySettings--23H2.xlsx";
        String uploadedStandardsFile = "/Users/morganreed/Documents/sp800-53r5-control-catalog.xlsx";
        String outputFile = "/Users/morganreed/Documents/updated_standards.xlsx";

        try {
            List<Map<String, String>> policiesData = readExcelFile(preloadedPoliciesFile);
            System.out.println("Number of records read from policies file: " + policiesData.size()); // Debug statement 1

            List<Map<String, String>> standardsData = readExcelFile(uploadedStandardsFile);
            System.out.println("Number of records read from standards file: " + standardsData.size()); // Debug statement 2

            if (policiesData.isEmpty() || standardsData.isEmpty()) {
                System.out.println("One or both input files are empty or could not be read.");
                return;
            }

            List<Map<String, String>> updatedStandards = compareAndAppend(standardsData, policiesData);
            System.out.println("Number of records after compare and append: " + updatedStandards.size()); // Debug statement 3

            if (updatedStandards.isEmpty()) {
                System.out.println("No data to write after comparing and appending.");
                return;
            }

            writeXlsxFile(updatedStandards, outputFile);
            System.out.println("Updated standards saved to " + outputFile);
        } catch (IOException e) {
            System.out.println("An error occurred: " + e.getMessage());
            e.printStackTrace();
        }
    }
    private List<Map<String, String>> readExcelFile(String filePath) throws IOException {
        List<Map<String, String>> resultList = new ArrayList<>();
        File file = new File(filePath);
        if (!file.exists() || !file.canRead()) {
            System.out.println("File does not exist or cannot be read: " + filePath);
            return resultList;
        }

        try (FileInputStream fileInputStream = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {
            Sheet sheet = workbook.getSheetAt(0);
            boolean isFirstRow = true;
            List<String> headers = new ArrayList<>();

            for (Row row : sheet) {
                if (isFirstRow) {
                    for (Cell cell : row) {
                        headers.add(cell.getStringCellValue());
                    }
                    isFirstRow = false;
                } else {
                    Map<String, String> dataMap = new HashMap<>();
                    for (int i = 0; i < headers.size(); i++) {
                        Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        dataMap.put(headers.get(i), cell.toString());
                    }
                    resultList.add(dataMap);
                }
            }
        } catch (FileNotFoundException e) {
            System.out.println("File not found: " + filePath);
            e.printStackTrace();
        } catch (IOException e) {
            System.out.println("Error reading file: " + filePath);
            e.printStackTrace();
        }
        return resultList;
    }
}