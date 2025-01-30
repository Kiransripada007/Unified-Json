package com.example.unified_json.controller;

import com.example.unified_json.service.ExcelToJsonConverter;
import com.example.unified_json.service.JsonToExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

@RestController
@RequestMapping("/api/excel")
public class ExcelController {

    @Autowired
    private ExcelToJsonConverter excelToJsonConverter;

    @Autowired
    private JsonToExcelService jsonToExcelService;

    @PostMapping("/convert")
    public ResponseEntity<String> convertExcelToJsonAndExcel(@RequestParam("file") MultipartFile file) {
        try (InputStream inputStream = file.getInputStream()) {
            // Define output file paths
            String jsonFilePath = "src/main/resources/output.json";
            String excelFilePath = "src/main/resources/output.xlsx";

            // Convert Excel to JSON
            excelToJsonConverter.convertExcelToJson(inputStream, jsonFilePath);

            // Convert JSON back to Excel
            jsonToExcelService.convertJsonToExcel(jsonFilePath, excelFilePath);

            // Prepare response with file paths
            return ResponseEntity.ok("Files generated successfully: \n" +
                    "JSON File: " + jsonFilePath + "\n" +
                    "Excel File: " + excelFilePath);
        } catch (Exception e) {
            return ResponseEntity.status(500).body("Error processing file: " + e.getMessage());
        }
    }

    // Optional: Method for downloading the generated files
    @GetMapping("/get/{fileType}")
    public ResponseEntity<byte[]> downloadFile(@PathVariable String fileType) {
        try {
            Path path;
            if ("json".equalsIgnoreCase(fileType)) {
                path = Paths.get("src/main/resources/output.json");
            } else if ("excel".equalsIgnoreCase(fileType)) {
                path = Paths.get("src/main/resources/output.xlsx");
            } else {
                return ResponseEntity.badRequest().body(null);
            }

            File file = path.toFile();
            byte[] fileContent = Files.readAllBytes(path);

            return ResponseEntity.ok()
                    .header("Content-Disposition", "attachment; filename=" + file.getName())
                    .body(fileContent);

        } catch (Exception e) {
            return ResponseEntity.status(500).body(null);
        }
    }
}
