package com.example.controller;

import com.example.model.AgeingByMonth;
import com.example.service.AgeingDataService;
import com.example.service.ExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.util.List;

@RestController
public class ExcelGeneratorController {

    private final AgeingDataService ageingDataService;
    private final ExcelService excelService;

    @Autowired
    public ExcelGeneratorController(AgeingDataService ageingDataService, ExcelService excelService) {
        this.ageingDataService = ageingDataService;
        this.excelService = excelService;
    }

    /**
     * Endpoint to generate an Excel file with ageing data
     * @param companyId the company ID to generate the report for
     * @return Excel file as a download
     */
    @GetMapping("/generate/{companyId}")
    public ResponseEntity<byte[]> generateExcel(@PathVariable String companyId) {
        try {
            // Get data from service (which would normally call external API)
            List<AgeingByMonth> ageingData = ageingDataService.getAgeingByMonth(companyId);
            
            // Generate Excel file
            byte[] excelContent = excelService.generateAgeingReport(ageingData);
            
            // Set up response headers for file download
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            headers.setContentDispositionFormData("attachment", "AgeingReport_" + companyId + ".xlsx");
            headers.setCacheControl("must-revalidate, post-check=0, pre-check=0");
            
            return ResponseEntity.ok()
                    .headers(headers)
                    .body(excelContent);
                    
        } catch (IOException e) {
            return ResponseEntity.internalServerError().build();
        }
    }
}
