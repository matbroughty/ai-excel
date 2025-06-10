package com.example.service;

import com.example.model.AgeingByMonth;
import com.example.model.Customer;
import com.example.model.Item;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.util.List;

@Service
public class ExcelService {

    private static final String[] HEADERS = {
            "Month", "Sales Ledger Balance", "Amount Not Due", "Over 30 Days", 
            "Over 60 Days", "Over 90 Days", "Over Threshold", "Total Credits", 
            "% Over 90 Days"
    };
    
    // Theme colors
    private static final byte[] HEADER_BLUE_RGB = new byte[]{(byte)0x1F, (byte)0x49, (byte)0x7D};
    private static final byte[] TITLE_BLUE_RGB = new byte[]{(byte)0x4F, (byte)0x81, (byte)0xBD};
    private static final byte[] ALTERNATE_ROW_BLUE_RGB = new byte[]{(byte)0xDC, (byte)0xE6, (byte)0xF1};
    private static final byte[] NEGATIVE_VALUE_RED_RGB = new byte[]{(byte)0xC0, (byte)0x50, (byte)0x4D};

    /**
     * Generates an Excel report from the ageing data with professional theming
     * @param ageingData List of ageing data by month
     * @return byte array containing the Excel file
     * @throws IOException if there's an error generating the Excel file
     */
    // Customer list sheet headers
    private static final String[] CUSTOMER_HEADERS = {
            "Customer ID", "Customer Name", "Balance", "Reference", 
            "Address Line 1", "Address Line 2", "City", "State/Province", 
            "Postal Code", "Country", "Notified", "Last Updated"
    };
    
    // Open items sheet headers
    private static final String[] ITEM_HEADERS = {
            "Customer ID", "Document Type", "Document Number", "Document Reference",
            "Document Date", "Due Date", "Entry Date", "Entry User",
            "Amount", "Balance"
    };
    
    /**
     * Generates an Excel report with three sheets: Ageing Report, Customer List, and Open Items
     * @param ageingData List of ageing data by month
     * @param customerData List of customers with outstanding balances
     * @param itemData List of open items for customers
     * @return byte array containing the Excel file
     * @throws IOException if there's an error generating the Excel file
     */
    public byte[] generateAgeingReport(List<AgeingByMonth> ageingData, List<Customer> customerData, 
                                      List<Item> itemData) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            // Create the Ageing Report sheet
            Sheet ageingSheet = workbook.createSheet("Ageing Report");
            createAgeingReportSheet(workbook, ageingSheet, ageingData);
            
            // Create the Customer List sheet
            Sheet customerSheet = workbook.createSheet("Customer List");
            createCustomerListSheet(workbook, customerSheet, customerData);
            
            // Create the Open Items sheet
            Sheet itemsSheet = workbook.createSheet("Open Items");
            createOpenItemsSheet(workbook, itemsSheet, itemData, customerData);
            
            // Write to byte array
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }
    
    /**
     * Creates the Open Items sheet
     */
    private void createOpenItemsSheet(XSSFWorkbook workbook, Sheet sheet, List<Item> itemData, List<Customer> customerData) {
        // Create title section
        Row titleRow = sheet.createRow(0);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("Open Items");
        
        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        Font titleFont = workbook.createFont();
        titleFont.setFontHeightInPoints((short) 14);
        titleFont.setBold(true);
        titleFont.setColor(IndexedColors.DARK_BLUE.getIndex());
        titleStyle.setFont(titleFont);
        
        titleCell.setCellStyle(titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, ITEM_HEADERS.length - 1));
        
        // Add generation date
        Row dateRow = sheet.createRow(1);
        Cell dateCell = dateRow.createCell(0);
        dateCell.setCellValue("Generated on: " + LocalDate.now().format(DateTimeFormatter.ofPattern("MMMM d, yyyy")));
        
        CellStyle dateStyle = workbook.createCellStyle();
        Font dateFont = workbook.createFont();
        dateFont.setItalic(true);
        dateFont.setColor(IndexedColors.GREY_50_PERCENT.getIndex());
        dateStyle.setFont(dateFont);
        
        dateCell.setCellStyle(dateStyle);
        
        // Create a blank row
        sheet.createRow(2);
        
        // Create header row
        int tableStartRow = 3;
        Row headerRow = sheet.createRow(tableStartRow);
        CellStyle headerStyle = createHeaderStyle(workbook);
        
        for (int i = 0; i < ITEM_HEADERS.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(ITEM_HEADERS[i]);
            cell.setCellStyle(headerStyle);
        }
        
        // Create styles for data rows
        CellStyle textCellStyle = createTextCellStyle(workbook);
        CellStyle dateCellStyle = createDateCellStyle(workbook);
        
        // Currency styles for different item types
        CellStyle invoiceCurrencyStyle = createCurrencyStyle(workbook);
        CellStyle negativeCurrencyStyle = createNegativeCurrencyStyle(workbook);
        
        // Type-specific text styles
        CellStyle invoiceTextStyle = createItemTypeStyle(workbook, IndexedColors.BLACK);
        CellStyle paymentTextStyle = createItemTypeStyle(workbook, IndexedColors.BLUE);
        CellStyle creditNoteTextStyle = createItemTypeStyle(workbook, IndexedColors.RED);
        
        // Create alternating row styles
        CellStyle alternateRowTextCellStyle = createAlternateRowStyle(workbook, textCellStyle);
        CellStyle alternateRowDateCellStyle = createAlternateRowStyle(workbook, dateCellStyle);
        CellStyle alternateRowInvoiceCurrencyStyle = createAlternateRowStyle(workbook, invoiceCurrencyStyle);
        CellStyle alternateRowNegativeCurrencyStyle = createAlternateRowStyle(workbook, negativeCurrencyStyle);
        CellStyle alternateRowInvoiceTextStyle = createAlternateRowStyle(workbook, invoiceTextStyle);
        CellStyle alternateRowPaymentTextStyle = createAlternateRowStyle(workbook, paymentTextStyle);
        CellStyle alternateRowCreditNoteTextStyle = createAlternateRowStyle(workbook, creditNoteTextStyle);
        
        // Create map of customer IDs to names for lookups
        java.util.Map<String, String> customerNames = new java.util.HashMap<>();
        for (Customer customer : customerData) {
            customerNames.put(customer.getCustomerId(), customer.getCustomerName());
        }
        
        // Add data rows
        int rowNum = tableStartRow + 1;
        for (Item item : itemData) {
            Row row = sheet.createRow(rowNum);
            boolean isAlternateRow = (rowNum - tableStartRow) % 2 == 0;
            
            // Choose appropriate style based on row parity and item type
            CellStyle rowTextStyle = isAlternateRow ? alternateRowTextCellStyle : textCellStyle;
            CellStyle rowDateStyle = isAlternateRow ? alternateRowDateCellStyle : dateCellStyle;
            
            // Get appropriate currency style based on item type
            CellStyle rowCurrencyStyle;
            if (item.isInvoice()) {
                rowCurrencyStyle = isAlternateRow ? alternateRowInvoiceCurrencyStyle : invoiceCurrencyStyle;
            } else {
                rowCurrencyStyle = isAlternateRow ? alternateRowNegativeCurrencyStyle : negativeCurrencyStyle;
            }
            
            // Get appropriate document type style based on item type
            CellStyle rowTypeStyle;
            if (item.isInvoice()) {
                rowTypeStyle = isAlternateRow ? alternateRowInvoiceTextStyle : invoiceTextStyle;
            } else if (item.isPayment()) {
                rowTypeStyle = isAlternateRow ? alternateRowPaymentTextStyle : paymentTextStyle;
            } else {
                rowTypeStyle = isAlternateRow ? alternateRowCreditNoteTextStyle : creditNoteTextStyle;
            }
            
            // Customer ID
            Cell customerIdCell = row.createCell(0);
            customerIdCell.setCellValue(item.getCustomerId() + " - " + 
                    (customerNames.containsKey(item.getCustomerId()) ? 
                    customerNames.get(item.getCustomerId()) : "Unknown"));
            customerIdCell.setCellStyle(rowTextStyle);
            
            // Document Type
            Cell typeCell = row.createCell(1);
            String docType;
            switch (item.getItemType()) {
                case "INV":
                    docType = "Invoice";
                    break;
                case "PAY":
                    docType = "Payment";
                    break;
                case "CRN":
                    docType = "Credit Note";
                    break;
                default:
                    docType = item.getItemType();
            }
            typeCell.setCellValue(docType);
            typeCell.setCellStyle(rowTypeStyle);
            
            // Document Number
            Cell docNumCell = row.createCell(2);
            docNumCell.setCellValue(item.getDocumentNumber());
            docNumCell.setCellStyle(rowTextStyle);
            
            // Document Reference
            Cell docRefCell = row.createCell(3);
            docRefCell.setCellValue(item.getDocumentReference());
            docRefCell.setCellStyle(rowTextStyle);
            
            // Document Date
            Cell docDateCell = row.createCell(4);
            docDateCell.setCellValue(item.getDocumentDate().format(DateTimeFormatter.ISO_LOCAL_DATE));
            docDateCell.setCellStyle(rowDateStyle);
            
            // Due Date
            Cell dueDateCell = row.createCell(5);
            dueDateCell.setCellValue(item.getDueDate().format(DateTimeFormatter.ISO_LOCAL_DATE));
            dueDateCell.setCellStyle(rowDateStyle);
            
            // Entry Date
            Cell entryDateCell = row.createCell(6);
            entryDateCell.setCellValue(item.getEntryDate().format(DateTimeFormatter.ISO_LOCAL_DATE));
            entryDateCell.setCellStyle(rowDateStyle);
            
            // Entry User
            Cell entryUserCell = row.createCell(7);
            entryUserCell.setCellValue(item.getEntryUser());
            entryUserCell.setCellStyle(rowTextStyle);
            
            // Amount
            Cell amountCell = row.createCell(8);
            amountCell.setCellValue(item.getAmount().doubleValue());
            amountCell.setCellStyle(rowCurrencyStyle);
            
            // Balance
            Cell balanceCell = row.createCell(9);
            balanceCell.setCellValue(item.getBalance().doubleValue());
            balanceCell.setCellStyle(rowCurrencyStyle);
            
            rowNum++;
        }
        
        // Add a total row at the bottom
        Row totalRow = sheet.createRow(rowNum);
        totalRow.setHeightInPoints(20);
        
        CellStyle totalLabelStyle = workbook.createCellStyle();
        totalLabelStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        totalLabelStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        totalLabelStyle.setBorderBottom(BorderStyle.THIN);
        totalLabelStyle.setBorderTop(BorderStyle.MEDIUM);
        totalLabelStyle.setBorderLeft(BorderStyle.THIN);
        totalLabelStyle.setBorderRight(BorderStyle.THIN);
        Font totalFont = workbook.createFont();
        totalFont.setBold(true);
        totalLabelStyle.setFont(totalFont);
        
        // Total label (spans several columns)
        Cell totalLabelCell = totalRow.createCell(0);
        totalLabelCell.setCellValue("TOTAL");
        totalLabelCell.setCellStyle(totalLabelStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, 0, 7));
        
        // Total amount
        CellStyle totalAmountStyle = workbook.createCellStyle();
        totalAmountStyle.cloneStyleFrom(totalLabelStyle);
        totalAmountStyle.setDataFormat(workbook.createDataFormat().getFormat("$#,##0.00"));
        
        Cell totalAmountCell = totalRow.createCell(8);
        totalAmountCell.setCellFormula(String.format("SUM(%s%d:%s%d)", 
                getColumnName(8), tableStartRow + 1, getColumnName(8), rowNum));
        totalAmountCell.setCellStyle(totalAmountStyle);
        
        // Total balance
        Cell totalBalanceCell = totalRow.createCell(9);
        totalBalanceCell.setCellFormula(String.format("SUM(%s%d:%s%d)", 
                getColumnName(9), tableStartRow + 1, getColumnName(9), rowNum));
        totalBalanceCell.setCellStyle(totalAmountStyle);
        
        // Add footer with info about color coding
        Row footerRow = sheet.createRow(rowNum + 2);
        Cell footerCell = footerRow.createCell(0);
        footerCell.setCellValue("Color coding: Black = Invoices, Blue = Payments, Red = Credit Notes");
        
        CellStyle footerStyle = workbook.createCellStyle();
        Font footerFont = workbook.createFont();
        footerFont.setItalic(true);
        footerFont.setColor(IndexedColors.GREY_50_PERCENT.getIndex());
        footerStyle.setFont(footerFont);
        footerCell.setCellStyle(footerStyle);
        
        // Merge cells for the footer
        sheet.addMergedRegion(new CellRangeAddress(rowNum + 2, rowNum + 2, 0, ITEM_HEADERS.length - 1));
        
        // Auto-size columns and add padding
        for (int i = 0; i < ITEM_HEADERS.length; i++) {
            sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 500);
        }
        
        // Freeze panes
        sheet.createFreezePane(0, tableStartRow + 1);
        
        // Add filter to headers
        sheet.setAutoFilter(new CellRangeAddress(
                tableStartRow, tableStartRow, 0, ITEM_HEADERS.length - 1));
    }
    
    /**
     * Creates a colored text style for item types
     */
    private CellStyle createItemTypeStyle(Workbook workbook, IndexedColors color) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        
        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(color.getIndex());
        style.setFont(font);
        
        return style;
    }
    
    /**
     * Creates the Ageing Report sheet
     */
    private void createAgeingReportSheet(XSSFWorkbook workbook, Sheet sheet, List<AgeingByMonth> ageingData) {
        // Create title and branding section
        createTitleSection(workbook, sheet, ageingData.size() > 0 ? ageingData.get(0).getMonth() : null);
        
        // Start the actual table at row 5
        int tableStartRow = 5;
        
        // Create header row
        Row headerRow = sheet.createRow(tableStartRow);
        CellStyle headerStyle = createHeaderStyle(workbook);
        
        for (int i = 0; i < HEADERS.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(HEADERS[i]);
            cell.setCellStyle(headerStyle);
        }
        
        // Create styles for data rows
        CellStyle dateCellStyle = createDateCellStyle(workbook);
        CellStyle currencyCellStyle = createCurrencyStyle(workbook);
        CellStyle negativeCurrencyStyle = createNegativeCurrencyStyle(workbook);
        CellStyle percentageCellStyle = createPercentageStyle(workbook);
        
        // Create alternating row styles
        CellStyle alternateRowDateCellStyle = createAlternateRowStyle(workbook, dateCellStyle);
        CellStyle alternateRowCurrencyCellStyle = createAlternateRowStyle(workbook, currencyCellStyle);
        CellStyle alternateRowNegativeCurrencyStyle = createAlternateRowStyle(workbook, negativeCurrencyStyle);
        CellStyle alternateRowPercentageCellStyle = createAlternateRowStyle(workbook, percentageCellStyle);
        
        DateTimeFormatter monthFormatter = DateTimeFormatter.ofPattern("MMM yyyy");
        
        int rowNum = tableStartRow + 1;
        
        // Create data rows with alternating colors
        for (AgeingByMonth data : ageingData) {
            Row row = sheet.createRow(rowNum);
            boolean isAlternateRow = (rowNum - tableStartRow) % 2 == 0;
            
            // Choose appropriate style based on row parity
            CellStyle rowDateStyle = isAlternateRow ? alternateRowDateCellStyle : dateCellStyle;
            CellStyle rowCurrencyStyle = isAlternateRow ? alternateRowCurrencyCellStyle : currencyCellStyle;
            CellStyle rowNegativeCurrencyStyle = isAlternateRow ? alternateRowNegativeCurrencyStyle : negativeCurrencyStyle;
            CellStyle rowPercentageStyle = isAlternateRow ? alternateRowPercentageCellStyle : percentageCellStyle;
            
            // Month
            Cell monthCell = row.createCell(0);
            monthCell.setCellValue(data.getMonth().format(monthFormatter));
            monthCell.setCellStyle(rowDateStyle);
            
            // Sales Ledger Balance
            createConditionalCurrencyCell(row, 1, data.getSalesLedgerBalance(), rowCurrencyStyle, rowNegativeCurrencyStyle);
            
            // Amount Not Due
            createConditionalCurrencyCell(row, 2, data.getAmountNotDue(), rowCurrencyStyle, rowNegativeCurrencyStyle);
            
            // Over 30 Days
            createConditionalCurrencyCell(row, 3, data.getAmountOver30Days(), rowCurrencyStyle, rowNegativeCurrencyStyle);
            
            // Over 60 Days
            createConditionalCurrencyCell(row, 4, data.getAmountOver60Days(), rowCurrencyStyle, rowNegativeCurrencyStyle);
            
            // Over 90 Days
            createConditionalCurrencyCell(row, 5, data.getAmountOver90Days(), rowCurrencyStyle, rowNegativeCurrencyStyle);
            
            // Over Threshold
            createConditionalCurrencyCell(row, 6, data.getAmountOverThreshold(), rowCurrencyStyle, rowNegativeCurrencyStyle);
            
            // Total Credits
            createConditionalCurrencyCell(row, 7, data.getTotalCredits(), rowCurrencyStyle, rowNegativeCurrencyStyle);
            
            // % Over 90 Days
            Cell percentCell = row.createCell(8);
            percentCell.setCellValue(data.getPercentOver90Days().doubleValue() / 100);
            percentCell.setCellStyle(rowPercentageStyle);
            
            rowNum++;
        }
        
        // Add summary row
        addSummaryRow(workbook, sheet, rowNum, tableStartRow + 1, rowNum - 1);
        
        // Add footer
        addFooter(workbook, sheet, rowNum + 2);
        
        // Apply table formatting
        for (int i = 0; i < HEADERS.length; i++) {
            sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 1000); // Add a bit of padding
        }
        
        // Freeze the header row
        sheet.createFreezePane(0, tableStartRow + 1);
    }
    
    /**
     * Creates the Customer List sheet
     */
    private void createCustomerListSheet(XSSFWorkbook workbook, Sheet sheet, List<Customer> customerData) {
        // Create title section
        Row titleRow = sheet.createRow(0);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("Customer List - Outstanding Balances");
        
        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        Font titleFont = workbook.createFont();
        titleFont.setFontHeightInPoints((short) 14);
        titleFont.setBold(true);
        titleFont.setColor(IndexedColors.DARK_BLUE.getIndex());
        titleStyle.setFont(titleFont);
        
        titleCell.setCellStyle(titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, CUSTOMER_HEADERS.length - 1));
        
        // Add generation date
        Row dateRow = sheet.createRow(1);
        Cell dateCell = dateRow.createCell(0);
        dateCell.setCellValue("Generated on: " + LocalDate.now().format(DateTimeFormatter.ofPattern("MMMM d, yyyy")));
        
        CellStyle dateStyle = workbook.createCellStyle();
        Font dateFont = workbook.createFont();
        dateFont.setItalic(true);
        dateFont.setColor(IndexedColors.GREY_50_PERCENT.getIndex());
        dateStyle.setFont(dateFont);
        
        dateCell.setCellStyle(dateStyle);
        
        // Create a blank row
        sheet.createRow(2);
        
        // Create header row
        int tableStartRow = 3;
        Row headerRow = sheet.createRow(tableStartRow);
        CellStyle headerStyle = createHeaderStyle(workbook);
        
        for (int i = 0; i < CUSTOMER_HEADERS.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(CUSTOMER_HEADERS[i]);
            cell.setCellStyle(headerStyle);
        }
        
        // Create styles for data rows
        CellStyle textCellStyle = createTextCellStyle(workbook);
        CellStyle currencyCellStyle = createCurrencyStyle(workbook);
        CellStyle dateCellStyle = createDateCellStyle(workbook);
        CellStyle booleanCellStyle = createBooleanCellStyle(workbook);
        
        // Create alternating row styles
        CellStyle alternateRowTextCellStyle = createAlternateRowStyle(workbook, textCellStyle);
        CellStyle alternateRowCurrencyCellStyle = createAlternateRowStyle(workbook, currencyCellStyle);
        CellStyle alternateRowDateCellStyle = createAlternateRowStyle(workbook, dateCellStyle);
        CellStyle alternateRowBooleanCellStyle = createAlternateRowStyle(workbook, booleanCellStyle);
        
        // Add data rows
        int rowNum = tableStartRow + 1;
        for (Customer customer : customerData) {
            Row row = sheet.createRow(rowNum);
            boolean isAlternateRow = (rowNum - tableStartRow) % 2 == 0;
            
            // Choose appropriate style based on row parity
            CellStyle rowTextStyle = isAlternateRow ? alternateRowTextCellStyle : textCellStyle;
            CellStyle rowCurrencyStyle = isAlternateRow ? alternateRowCurrencyCellStyle : currencyCellStyle;
            CellStyle rowDateStyle = isAlternateRow ? alternateRowDateCellStyle : dateCellStyle;
            CellStyle rowBooleanStyle = isAlternateRow ? alternateRowBooleanCellStyle : booleanCellStyle;
            
            // Customer ID
            Cell idCell = row.createCell(0);
            idCell.setCellValue(customer.getCustomerId());
            idCell.setCellStyle(rowTextStyle);
            
            // Customer Name
            Cell nameCell = row.createCell(1);
            nameCell.setCellValue(customer.getCustomerName());
            nameCell.setCellStyle(rowTextStyle);
            
            // Balance
            Cell balanceCell = row.createCell(2);
            balanceCell.setCellValue(customer.getBalance().doubleValue());
            balanceCell.setCellStyle(rowCurrencyStyle);
            
            // Reference
            Cell refCell = row.createCell(3);
            refCell.setCellValue(customer.getReference());
            refCell.setCellStyle(rowTextStyle);
            
            // Address Line 1
            Cell addr1Cell = row.createCell(4);
            addr1Cell.setCellValue(customer.getAddressLine1());
            addr1Cell.setCellStyle(rowTextStyle);
            
            // Address Line 2
            Cell addr2Cell = row.createCell(5);
            addr2Cell.setCellValue(customer.getAddressLine2());
            addr2Cell.setCellStyle(rowTextStyle);
            
            // City
            Cell cityCell = row.createCell(6);
            cityCell.setCellValue(customer.getCity());
            cityCell.setCellStyle(rowTextStyle);
            
            // State
            Cell stateCell = row.createCell(7);
            stateCell.setCellValue(customer.getState());
            stateCell.setCellStyle(rowTextStyle);
            
            // Postal Code
            Cell postalCell = row.createCell(8);
            postalCell.setCellValue(customer.getPostalCode());
            postalCell.setCellStyle(rowTextStyle);
            
            // Country
            Cell countryCell = row.createCell(9);
            countryCell.setCellValue(customer.getCountry());
            countryCell.setCellStyle(rowTextStyle);
            
            // Notified
            Cell notifiedCell = row.createCell(10);
            notifiedCell.setCellValue(customer.isNotified() ? "Yes" : "No");
            notifiedCell.setCellStyle(rowBooleanStyle);
            
            // Last Updated
            Cell updatedCell = row.createCell(11);
            updatedCell.setCellValue(customer.getLastUpdated().format(DateTimeFormatter.ISO_LOCAL_DATE));
            updatedCell.setCellStyle(rowDateStyle);
            
            rowNum++;
        }
        
        // Add a total row at the bottom
        Row totalRow = sheet.createRow(rowNum);
        totalRow.setHeightInPoints(20);
        
        CellStyle totalLabelStyle = workbook.createCellStyle();
        totalLabelStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        totalLabelStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        totalLabelStyle.setBorderBottom(BorderStyle.THIN);
        totalLabelStyle.setBorderTop(BorderStyle.MEDIUM);
        totalLabelStyle.setBorderLeft(BorderStyle.THIN);
        totalLabelStyle.setBorderRight(BorderStyle.THIN);
        Font totalFont = workbook.createFont();
        totalFont.setBold(true);
        totalLabelStyle.setFont(totalFont);
        
        // Total label (spans first two columns)
        Cell totalLabelCell = totalRow.createCell(0);
        totalLabelCell.setCellValue("TOTAL OUTSTANDING");
        totalLabelCell.setCellStyle(totalLabelStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, 0, 1));
        
        // Total balance
        CellStyle totalBalanceStyle = workbook.createCellStyle();
        totalBalanceStyle.cloneStyleFrom(totalLabelStyle);
        totalBalanceStyle.setDataFormat(workbook.createDataFormat().getFormat("$#,##0.00"));
        
        Cell totalBalanceCell = totalRow.createCell(2);
        totalBalanceCell.setCellFormula(String.format("SUM(%s%d:%s%d)", 
                getColumnName(2), tableStartRow + 1, getColumnName(2), rowNum));
        totalBalanceCell.setCellStyle(totalBalanceStyle);
        
        // Fill the rest of the total row with the same style
        for (int i = 3; i < CUSTOMER_HEADERS.length; i++) {
            Cell cell = totalRow.createCell(i);
            cell.setCellStyle(totalLabelStyle);
        }
        
        // Auto-size columns and add padding
        for (int i = 0; i < CUSTOMER_HEADERS.length; i++) {
            sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 500);
        }
        
        // Freeze panes
        sheet.createFreezePane(0, tableStartRow + 1);
        
        // Add filter to headers
        sheet.setAutoFilter(new CellRangeAddress(
                tableStartRow, tableStartRow, 0, CUSTOMER_HEADERS.length - 1));
    }
    
    /**
     * Creates a text cell style for the customer list
     */
    private CellStyle createTextCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }
    
    /**
     * Creates a boolean cell style for yes/no values
     */
    private CellStyle createBooleanCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }
    
    /**
     * Creates the title section with company branding
     */
    private void createTitleSection(XSSFWorkbook workbook, Sheet sheet, YearMonth reportMonth) {
        // Create a merged cell for the title
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, HEADERS.length - 1));
        
        // Title row
        Row titleRow = sheet.createRow(0);
        titleRow.setHeightInPoints(30); // Taller row for title
        
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("Ageing Report");
        
        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        
        // Set background color
        titleStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        // Set border
        titleStyle.setBorderBottom(BorderStyle.NONE);
        titleStyle.setBorderTop(BorderStyle.NONE);
        titleStyle.setBorderLeft(BorderStyle.NONE);
        titleStyle.setBorderRight(BorderStyle.NONE);
        
        // Create title font
        Font titleFont = workbook.createFont();
        titleFont.setFontHeightInPoints((short) 18);
        titleFont.setBold(true);
        titleFont.setColor(IndexedColors.DARK_BLUE.getIndex());
        titleStyle.setFont(titleFont);
        
        titleCell.setCellStyle(titleStyle);
        
        // Subtitle with date
        Row subtitleRow = sheet.createRow(1);
        Cell subtitleCell = subtitleRow.createCell(0);
        
        String reportDate = LocalDate.now().format(DateTimeFormatter.ofPattern("MMMM d, yyyy"));
        String monthInfo = reportMonth != null ? 
            "Data as of " + reportMonth.format(DateTimeFormatter.ofPattern("MMMM yyyy")) : 
            "Report generated on " + reportDate;
            
        subtitleCell.setCellValue(monthInfo);
        
        CellStyle subtitleStyle = workbook.createCellStyle();
        Font subtitleFont = workbook.createFont();
        subtitleFont.setItalic(true);
        subtitleFont.setColor(IndexedColors.GREY_50_PERCENT.getIndex());
        subtitleStyle.setFont(subtitleFont);
        subtitleCell.setCellStyle(subtitleStyle);
        
        // Company info (normally would come from a parameter)
        Row companyRow = sheet.createRow(2);
        Cell companyCell = companyRow.createCell(0);
        companyCell.setCellValue("Company: Financial Services Ltd.");
        
        // Leave a blank row
        sheet.createRow(3);
    }
    
    /**
     * Adds a summary row at the bottom of the data
     */
    private void addSummaryRow(Workbook workbook, Sheet sheet, int rowNum, int firstDataRow, int lastDataRow) {
        Row summaryRow = sheet.createRow(rowNum);
        summaryRow.setHeightInPoints(20); // Slightly taller row
        
        // Create summary row style
        CellStyle summaryStyle = workbook.createCellStyle();
        summaryStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        summaryStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        summaryStyle.setBorderBottom(BorderStyle.THIN);
        summaryStyle.setBorderTop(BorderStyle.MEDIUM);
        summaryStyle.setBorderLeft(BorderStyle.THIN);
        summaryStyle.setBorderRight(BorderStyle.THIN);
        summaryStyle.setDataFormat(workbook.createDataFormat().getFormat("$#,##0.00"));
        
        Font summaryFont = workbook.createFont();
        summaryFont.setBold(true);
        summaryStyle.setFont(summaryFont);
        
        // Create summary label
        Cell labelCell = summaryRow.createCell(0);
        labelCell.setCellValue("TOTAL");
        labelCell.setCellStyle(summaryStyle);
        
        // For columns 1-7 (the currency columns), create sum formulas
        for (int i = 1; i <= 7; i++) {
            Cell cell = summaryRow.createCell(i);
            // Create a SUM formula for this column
            cell.setCellFormula(String.format("SUM(%s%d:%s%d)", 
                    getColumnName(i), firstDataRow + 1, getColumnName(i), lastDataRow + 1));
            cell.setCellStyle(summaryStyle);
        }
        
        // Average for the percentage column
        CellStyle avgPercentStyle = workbook.createCellStyle();
        avgPercentStyle.cloneStyleFrom(summaryStyle);
        avgPercentStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00%"));
        
        Cell avgCell = summaryRow.createCell(8);
        avgCell.setCellFormula(String.format("AVERAGE(%s%d:%s%d)", 
                getColumnName(8), firstDataRow + 1, getColumnName(8), lastDataRow + 1));
        avgCell.setCellStyle(avgPercentStyle);
    }
    
    /**
     * Adds a footer with notes
     */
    private void addFooter(Workbook workbook, Sheet sheet, int rowNum) {
        Row footerRow = sheet.createRow(rowNum);
        Cell footerCell = footerRow.createCell(0);
        footerCell.setCellValue("Note: This report was automatically generated. For questions, contact finance@example.com");
        
        CellStyle footerStyle = workbook.createCellStyle();
        Font footerFont = workbook.createFont();
        footerFont.setItalic(true);
        footerFont.setColor(IndexedColors.GREY_50_PERCENT.getIndex());
        footerStyle.setFont(footerFont);
        footerCell.setCellStyle(footerStyle);
        
        // Merge cells for the footer
        sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, 0, HEADERS.length - 1));
    }
    
    /**
     * Helper method to convert column index to column name (A, B, C...)
     */
    private String getColumnName(int columnIndex) {
        StringBuilder sb = new StringBuilder();
        while (columnIndex >= 0) {
            sb.append((char) ('A' + (columnIndex % 26)));
            columnIndex = (columnIndex / 26) - 1;
        }
        return sb.reverse().toString();
    }
    
    /**
     * Writes the Excel report to a file at the specified path
     * @param excelContent byte array containing the Excel file
     * @param filePath path where the file should be saved
     * @param fileName name of the file to save
     * @throws IOException if there's an error writing the file
     */
    public void saveExcelToFile(byte[] excelContent, String filePath, String fileName) throws IOException {
        java.nio.file.Path directoryPath = java.nio.file.Paths.get(filePath);
        java.nio.file.Path fullPath = directoryPath.resolve(fileName);
        
        // Create directories if they don't exist
        java.nio.file.Files.createDirectories(directoryPath);
        
        // Write the file
        java.nio.file.Files.write(fullPath, excelContent);
    }
    
    /**
     * Creates a cell with currency value, using different styles for positive and negative values
     */
    private void createConditionalCurrencyCell(Row row, int cellIndex, BigDecimal value, CellStyle positiveStyle, CellStyle negativeStyle) {
        Cell cell = row.createCell(cellIndex);
        cell.setCellValue(value.doubleValue());
        cell.setCellStyle(value.compareTo(BigDecimal.ZERO) < 0 ? negativeStyle : positiveStyle);
    }
    
    /**
     * Creates a standard currency cell
     */
    private void createCurrencyCell(Row row, int cellIndex, BigDecimal value, CellStyle style) {
        Cell cell = row.createCell(cellIndex);
        cell.setCellValue(value.doubleValue());
        cell.setCellStyle(style);
    }
    
    /**
     * Creates the header style with custom blue background
     */
    private CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        
        // Set custom blue color for headers
        XSSFColor headerColor = new XSSFColor(HEADER_BLUE_RGB, null);
        style.setFillForegroundColor(IndexedColors.ROYAL_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        // Add borders
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        
        // Text alignment
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        
        // Create white bold font for header text
        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);
        
        return style;
    }
    
    /**
     * Creates date cell style
     */
    private CellStyle createDateCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.LEFT);
        return style;
    }
    
    /**
     * Creates currency style
     */
    private CellStyle createCurrencyStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setDataFormat(workbook.createDataFormat().getFormat("$#,##0.00"));
        style.setAlignment(HorizontalAlignment.RIGHT);
        return style;
    }
    
    /**
     * Creates a currency style specifically for negative values
     */
    private CellStyle createNegativeCurrencyStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setDataFormat(workbook.createDataFormat().getFormat("$#,##0.00"));
        style.setAlignment(HorizontalAlignment.RIGHT);
        
        // Red font for negative values
        Font font = workbook.createFont();
        font.setColor(IndexedColors.RED.getIndex());
        style.setFont(font);
        
        return style;
    }
    
    /**
     * Creates percentage style
     */
    private CellStyle createPercentageStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setDataFormat(workbook.createDataFormat().getFormat("0.00%"));
        style.setAlignment(HorizontalAlignment.RIGHT);
        return style;
    }
    
    /**
     * Creates alternating row style based on another style
     */
    private CellStyle createAlternateRowStyle(Workbook workbook, CellStyle baseStyle) {
        CellStyle style = workbook.createCellStyle();
        
        // Clone the base style first
        style.cloneStyleFrom(baseStyle);
        
        // Set alternate row color
        style.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        return style;
    }
}
