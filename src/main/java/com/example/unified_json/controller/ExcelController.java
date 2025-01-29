package com.example.unified_json.controller;

import com.example.unified_json.service.ExcelToJsonConverter;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;

@RestController
@RequestMapping("/api/excel")
public class ExcelController {

    @Autowired
    private ExcelToJsonConverter converter;

    @PostMapping("/convert")
    public ResponseEntity<String> convertExcelToJson(@RequestParam("file") MultipartFile file) {
        try (InputStream inputStream = file.getInputStream()) {
            String outputFilePath = "src/main/resources/output.json";
            converter.convertExcelToJson(inputStream, outputFilePath);
            return ResponseEntity.ok("JSON file generated successfully at: " + outputFilePath);
        } catch (Exception e) {
            return ResponseEntity.status(500).body("Error processing file: " + e.getMessage());
        }
    }
}