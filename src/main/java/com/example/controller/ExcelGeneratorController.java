package com.example.controller;

import com.example.model.AgeingByMonth;
import com.example.model.CompanySummary;
import com.example.model.Customer;
import com.example.model.Item;
import com.example.service.AgeingDataService;
import com.example.service.CompanySummaryService;
import com.example.service.CustomerService;
import com.example.service.ExcelService;
import com.example.service.ItemService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.nio.file.Paths;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;

@RestController
public class ExcelGeneratorController {

    private static final Logger logger = LoggerFactory.getLogger(ExcelGeneratorController.class);
    private final AgeingDataService ageingDataService;
    private final CustomerService customerService;
    private final ItemService itemService;
    private final CompanySummaryService companySummaryService;
    private final ExcelService excelService;
    
    @Value("${excel.default.output.path:/users/mathewbroughton}")
    private String defaultOutputPath;

    @Autowired
    public ExcelGeneratorController(
            AgeingDataService ageingDataService,
            CustomerService customerService,
            ItemService itemService,
            CompanySummaryService companySummaryService,
            ExcelService excelService) {
        this.ageingDataService = ageingDataService;
        this.customerService = customerService;
        this.itemService = itemService;
        this.companySummaryService = companySummaryService;
        this.excelService = excelService;
    }

    /**
     * Endpoint to generate an Excel file with ageing data, customer list and open items
     * @param companyId the company ID to generate the report for
     * @param outputPath optional path to save the Excel file (defaults to configured path)
     * @return Excel file as a download
     */
    @GetMapping("/generate/{companyId}")
    public ResponseEntity<byte[]> generateExcel(
            @PathVariable String companyId,
            @RequestParam(required = false) String outputPath) {
        
        try {
            // Get ageing data from service
            List<AgeingByMonth> ageingData = ageingDataService.getAgeingByMonth(companyId);
            
            // Get customer data from service
            List<Customer> customerData = customerService.getCustomersWithOutstandingBalance(companyId);
            logger.info("Retrieved {} customers for company ID: {}", customerData.size(), companyId);

            // Get open items data from service
            List<Item> itemData = itemService.getOpenItems(companyId);
            logger.info("Retrieved {} open items for company ID: {}", itemData.size(), companyId);

            // Validate that all items have valid customer IDs
            Set<String> customerIds = customerData.stream()
                    .map(Customer::getCustomerId)
                    .collect(Collectors.toSet());

            long invalidItems = itemData.stream()
                    .filter(item -> !customerIds.contains(item.getCustomerId()))
                    .count();

            if (invalidItems > 0) {
                logger.warn("{} items have customer IDs that don't match any customer in the customer data", invalidItems);
            }
            
            // Get company summary data
            CompanySummary companySummary = companySummaryService.getCompanySummary(companyId);
            logger.info("Retrieved company summary for company ID: {}", companyId);

            // Generate Excel file with all sheets including the new summary sheet
            byte[] excelContent = excelService.generateAgeingReport(ageingData, customerData, itemData, companySummary);
            
            // Generate filename for the report
            String fileName = "SummaryReport_" + companyId + ".xlsx";
            
            // If outputPath is provided, save the file to disk
            String filePath = outputPath != null ? outputPath : defaultOutputPath;
            if (filePath != null && !filePath.isEmpty()) {
                try {
                    excelService.saveExcelToFile(excelContent, filePath, fileName);
                    logger.info("Excel file saved to: {}", Paths.get(filePath, fileName));
                } catch (IOException e) {
                    logger.error("Failed to save Excel file to disk", e);
                    // Continue to return the file even if saving to disk fails
                }
            }
            
            // Set up response headers for file download
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            headers.setContentDispositionFormData("attachment", fileName);
            headers.setCacheControl("must-revalidate, post-check=0, pre-check=0");
            
            return ResponseEntity.ok()
                    .headers(headers)
                    .body(excelContent);
                    
        } catch (IOException e) {
            logger.error("Error generating Excel report", e);
            return ResponseEntity.internalServerError().build();
        }
    }
}
