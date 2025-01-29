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

        // Data structures to hold parsed data
        Map<String, Map<String, Object>> referenceDataAssets = new HashMap<>();
        Map<String, Map<String, Object>> codeValues = new HashMap<>();
        Map<String, List<Map<String, String>>> hierarchies = new HashMap<>();
        List<Map<String, Object>> mappings = new ArrayList<>();

        // Process each sheet
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

            // Process rows in the sheet
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
                                } else {
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

                // Map data based on sheet type
                if (sheetName.equals("1. Reference Data Assets")) {
                    String referenceDataName = (String) rowData.get("Reference Data Name*");
                    referenceDataAssets.put(referenceDataName, rowData);
                } else if (sheetName.equals("2. Code Values")) {
                    String referenceDataName = (String) rowData.get("Reference Data Name*");
                    String codeValue = (String) rowData.get("Reference Data Name*");
                    codeValues.put(codeValue, rowData);
                } else if (sheetName.equals("3. Hierarchy")) {
                    String parent = (String) rowData.get("Parent");
                    Map<String, String> children = new HashMap<>();
                    for (int c = 1; c <= 8; c++) {
                        String child = (String) rowData.get("Child " + c);
                        if (child != null && !child.isEmpty()) {
                            children.put("Child " + c, child);
                        }
                    }
                    hierarchies.put(parent, Collections.singletonList(children));
                } else if (sheetName.equals("4. Mapping")) {
                    mappings.add(rowData);
                }
            }
        }

        workbook.close();

        // Build the unified JSON structure by integrating the logic

        // Add Reference Data Assets with corresponding Code Values
        Map<String, Object> referenceDataWithCodeValues = new HashMap<>();
        for (String referenceDataName : referenceDataAssets.keySet()) {
            Map<String, Object> referenceDataAsset = referenceDataAssets.get(referenceDataName);
            List<Map<String, Object>> relatedCodeValues = new ArrayList<>();

            // Add the code values for the given Reference Data Name
            for (String codeValue : codeValues.keySet()) {
                if (codeValues.get(codeValue).get("Reference Data Name*").equals(referenceDataName)) {
                    relatedCodeValues.add(codeValues.get(codeValue));
                }
            }

            referenceDataAsset.put("relatedCodeValues", relatedCodeValues);
            referenceDataWithCodeValues.put(referenceDataName, referenceDataAsset);
        }

        unifiedJson.put("referenceDataAssets", referenceDataWithCodeValues);

        // Add Hierarchy Relationships to Code Values
        Map<String, Object> codeValuesWithHierarchy = new HashMap<>();
        for (String codeValue : codeValues.keySet()) {
            Map<String, Object> codeValueData = codeValues.get(codeValue);
            List<Map<String, String>> hierarchyData = new ArrayList<>();

            // Check if codeValue has a parent-child hierarchy
            for (String parent : hierarchies.keySet()) {
                if (parent.equals(codeValue)) {
                    hierarchyData.addAll(hierarchies.get(parent));
                }
            }

            codeValueData.put("hierarchy", hierarchyData);
            codeValuesWithHierarchy.put(codeValue, codeValueData);
        }

        unifiedJson.put("codeValues", codeValuesWithHierarchy);

        // Add Mappings to Code Values
        for (Map<String, Object> mapping : mappings) {
            String sourceCodeValue = (String) mapping.get("Source Code Value*");
            String targetCodeValue = (String) mapping.get("Target Code Value*");

            // Add mapping information to the source code value
            if (codeValuesWithHierarchy.containsKey(sourceCodeValue)) {
                Map<String, Object> sourceCodeValueData = (Map<String, Object>) codeValuesWithHierarchy.get(sourceCodeValue);
                List<Map<String, String>> mappingsList = (List<Map<String, String>>) sourceCodeValueData.getOrDefault("mappings", new ArrayList<>());
                mappingsList.add(Collections.singletonMap("targetCodeValue", targetCodeValue));
                sourceCodeValueData.put("mappings", mappingsList);
            }
        }

        // Write the unified JSON to a file
        ObjectMapper objectMapper = new ObjectMapper();
        objectMapper.writerWithDefaultPrettyPrinter().writeValue(new File(outputFilePath), unifiedJson);
    }
}