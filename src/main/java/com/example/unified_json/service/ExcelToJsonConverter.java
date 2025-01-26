package com.example.unified_json.service;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.*;
import java.util.logging.Logger;



@Service
public class ExcelToJsonConverter {

    private static final Logger LOGGER = Logger.getLogger(ExcelToJsonConverter.class.getName());

    public void convertExcelToJson(InputStream inputStream, String outputFilePath) throws Exception {
        Workbook workbook = new XSSFWorkbook(inputStream);
        Map<String, Object> unifiedJson = new HashMap<>();

        // Process each sheet
        List<Map<String, Object>> sheetsData = new ArrayList<>();

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            String sheetName = sheet.getSheetName();

            if (sheet.getLastRowNum() == 0) {
                LOGGER.warning("Sheet '" + sheetName + "' is empty and will be skipped.");
                continue;
            }

            Map<String, Object> sheetJson = new LinkedHashMap<>();
            List<Map<String, Object>> sheetData = new ArrayList<>();
            Row headerRow = sheet.getRow(0);

            if (headerRow == null) {
                LOGGER.warning("Sheet '" + sheetName + "' does not have a header row and will be skipped.");
                continue;
            }

            // Get headers
            List<String> headers = new ArrayList<>();
            for (Cell cell : headerRow) {
                headers.add(cell.getStringCellValue());
            }

            // Process rows
            for (int j = 1; j <= sheet.getLastRowNum(); j++) {
                Row row = sheet.getRow(j);
                if (row == null) continue;

                Map<String, Object> rowData = new LinkedHashMap<>();
                for (int k = 0; k < headers.size(); k++) {
                    Cell cell = row.getCell(k);
                    String header = headers.get(k);

                    if (cell != null) {
                        switch (cell.getCellType()) {
                            case STRING:
                                rowData.put(header, cell.getStringCellValue());
                                break;
                            case NUMERIC:
                                if (DateUtil.isCellDateFormatted(cell)) {
                                // Convert numeric value to LocalDate
                                LocalDate date = cell.getDateCellValue().toInstant()
                                        .atZone(ZoneId.systemDefault())
                                        .toLocalDate();
                                rowData.put(header, date.toString());
                            }
                                else{
                                    rowData.put(header, cell.getNumericCellValue());
                                }
                                break;
                            case BOOLEAN:
                                rowData.put(header, cell.getBooleanCellValue());
                                break;
                            default:
                                rowData.put(header, null);
                                LOGGER.warning("Invalid cell type in row " + (j + 1) + " under header '" + header + "'. Setting to null.");
                        }
                    }
                }
                sheetData.add(rowData);
            }

            sheetJson.put("sheetName", sheetName);
            sheetJson.put("rowCount", sheet.getLastRowNum());
            sheetJson.put("data", sheetData);

            sheetsData.add(sheetJson);
        }

        workbook.close();
        unifiedJson.put("sheets", sheetsData);

        ObjectMapper objectMapper = new ObjectMapper();
        objectMapper.writerWithDefaultPrettyPrinter().writeValue(new File(outputFilePath), unifiedJson);
    }
}
