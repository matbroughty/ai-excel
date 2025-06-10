package com.example.service;

import com.example.model.AgeingByMonth;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.format.DateTimeFormatter;
import java.util.List;

@Service
public class ExcelService {

    private static final String[] HEADERS = {
            "Month", "Sales Ledger Balance", "Amount Not Due", "Over 30 Days", 
            "Over 60 Days", "Over 90 Days", "Over Threshold", "Total Credits", 
            "% Over 90 Days"
    };

    /**
     * Generates an Excel report from the ageing data
     * @param ageingData List of ageing data by month
     * @return byte array containing the Excel file
     * @throws IOException if there's an error generating the Excel file
     */
    public byte[] generateAgeingReport(List<AgeingByMonth> ageingData) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Ageing Report");

            // Create header row
            Row headerRow = sheet.createRow(0);
            CellStyle headerStyle = createHeaderStyle(workbook);
            
            for (int i = 0; i < HEADERS.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(HEADERS[i]);
                cell.setCellStyle(headerStyle);
            }

            // Create data rows
            CellStyle dateCellStyle = createDateCellStyle(workbook);
            CellStyle currencyCellStyle = createCurrencyStyle(workbook);
            CellStyle percentageCellStyle = createPercentageStyle(workbook);
            
            DateTimeFormatter monthFormatter = DateTimeFormatter.ofPattern("MMM yyyy");
            
            int rowNum = 1;
            for (AgeingByMonth data : ageingData) {
                Row row = sheet.createRow(rowNum++);
                
                // Month
                Cell monthCell = row.createCell(0);
                monthCell.setCellValue(data.getMonth().format(monthFormatter));
                monthCell.setCellStyle(dateCellStyle);
                
                // Sales Ledger Balance
                createCurrencyCell(row, 1, data.getSalesLedgerBalance(), currencyCellStyle);
                
                // Amount Not Due
                createCurrencyCell(row, 2, data.getAmountNotDue(), currencyCellStyle);
                
                // Over 30 Days
                createCurrencyCell(row, 3, data.getAmountOver30Days(), currencyCellStyle);
                
                // Over 60 Days
                createCurrencyCell(row, 4, data.getAmountOver60Days(), currencyCellStyle);
                
                // Over 90 Days
                createCurrencyCell(row, 5, data.getAmountOver90Days(), currencyCellStyle);
                
                // Over Threshold
                createCurrencyCell(row, 6, data.getAmountOverThreshold(), currencyCellStyle);
                
                // Total Credits
                createCurrencyCell(row, 7, data.getTotalCredits(), currencyCellStyle);
                
                // % Over 90 Days
                Cell percentCell = row.createCell(8);
                percentCell.setCellValue(data.getPercentOver90Days().doubleValue() / 100);
                percentCell.setCellStyle(percentageCellStyle);
            }

            // Auto-size columns
            for (int i = 0; i < HEADERS.length; i++) {
                sheet.autoSizeColumn(i);
            }

            // Write to byte array
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    private void createCurrencyCell(Row row, int cellIndex, BigDecimal value, CellStyle style) {
        Cell cell = row.createCell(cellIndex);
        cell.setCellValue(value.doubleValue());
        cell.setCellStyle(style);
    }
    
    private CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        
        return style;
    }
    
    private CellStyle createDateCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }
    
    private CellStyle createCurrencyStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setDataFormat(workbook.createDataFormat().getFormat("$#,##0.00"));
        return style;
    }
    
    private CellStyle createPercentageStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setDataFormat(workbook.createDataFormat().getFormat("0.00%"));
        return style;
    }
}
