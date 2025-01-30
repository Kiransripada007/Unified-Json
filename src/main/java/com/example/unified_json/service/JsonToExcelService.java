package com.example.unified_json.service;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

@Service
public class JsonToExcelService {

    public void convertJsonToExcel(String inputJsonFilePath, String outputExcelFilePath) throws Exception {
        // Read the unified JSON file
        ObjectMapper objectMapper = new ObjectMapper();
        Map<String, Object> unifiedJson = objectMapper.readValue(new File(inputJsonFilePath), Map.class);

        // Get the data from the "sheets" field
        List<Map<String, Object>> sheetsData = (List<Map<String, Object>>) unifiedJson.get("sheets");

        // Create a new Workbook for the Excel file
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("UnifiedData");

        // Step 1: Collect all unique headers from all sheets
        Set<String> allHeaders = new LinkedHashSet<>();
        for (Map<String, Object> sheetData : sheetsData) {
            List<Map<String, Object>> rowsData = (List<Map<String, Object>>) sheetData.get("data");
            if (!rowsData.isEmpty()) {
                // Add all headers from the first row of this sheet to the global headers set
                allHeaders.addAll(rowsData.get(0).keySet());
            }
        }

        // Step 2: Create the header row in the Excel sheet
        Row headerRow = sheet.createRow(0);
        int colIndex = 0;
        Map<String, Integer> headerColumnIndexMap = new HashMap<>();
        for (String header : allHeaders) {
            headerColumnIndexMap.put(header, colIndex);
            Cell cell = headerRow.createCell(colIndex++);
            cell.setCellValue(header);
        }

        // Step 3: Write data rows below the header
        int rowIndex = 1;  // Start inserting data from the second row
        for (Map<String, Object> sheetData : sheetsData) {
            List<Map<String, Object>> rowsData = (List<Map<String, Object>>) sheetData.get("data");
            for (Map<String, Object> rowData : rowsData) {
                Row row = sheet.createRow(rowIndex++);
                for (Map.Entry<String, Object> entry : rowData.entrySet()) {
                    String header = entry.getKey();
                    Object value = entry.getValue();

                    // Get the column index based on the header
                    Integer col = headerColumnIndexMap.get(header);
                    if (col != null) {
                        Cell cell = row.createCell(col);

                        // Set the appropriate value depending on the type
                        if (value instanceof String) {
                            cell.setCellValue((String) value);
                        } else if (value instanceof Double) {
                            cell.setCellValue((Double) value);
                        } else if (value instanceof Boolean) {
                            cell.setCellValue((Boolean) value);
                        } else if (value instanceof String && ((String) value).matches("\\d{4}-\\d{2}-\\d{2}")) {
                            cell.setCellValue((String) value);
                            CellStyle dateCellStyle = workbook.createCellStyle();
                            short dateFormat = workbook.createDataFormat().getFormat("yyyy-mm-dd");
                            dateCellStyle.setDataFormat(dateFormat);
                            cell.setCellStyle(dateCellStyle);
                        } else {
                            cell.setCellValue(value != null ? value.toString() : "");
                        }
                    }
                }
            }
        }

        // Step 4: Write the Excel file to the specified output path
        try (FileOutputStream fileOut = new FileOutputStream(new File(outputExcelFilePath))) {
            workbook.write(fileOut);
        }

        workbook.close();
    }
}
