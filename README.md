# Excel Generator Application

This Spring Boot application generates Excel reports containing ageing data by month and customer information. The application exposes a REST endpoint that returns an Excel file populated with data from mocked APIs.

## Features

- REST endpoint `/generate/{companyId}` that returns an Excel file
- Mock data services that simulate retrieving data from external APIs
- Excel report generation using Apache POI
- Multi-sheet Excel report with professional styling
- Summary dashboard with company information and key statistics
- Data visualization with interactive charts

### Ageing Report Sheet

The first sheet includes monthly ageing data with:
  - Sales Ledger Balance
  - Amount Not Due
  - Amount Over 30 Days Past Due
  - Amount Over 60 Days Past Due
  - Amount Over 90 Days Past Due
  - Amount Over Threshold
  - Total Credits
  - Percentage Over 90 Days Past Due
  - Interactive trend charts showing:
    - Line chart of aging categories over time
    - Column chart showing percentage of debt over 90 days

### Customer List Sheet

The second sheet contains a list of all customers/debtors with outstanding balances:
  - Customer ID
  - Customer Name
  - Balance
  - Reference
  - Address Information (multiple fields)
  - Notification Status
  - Last Update Date
  - Total Outstanding Balance summary

### Open Items Sheet

The third sheet details all open items (invoices, payments, credit notes) for customers:
  - Customer ID and Name
  - Document Type (color-coded: Invoices in black, Payments in blue, Credit Notes in red)
  - Document Number and Reference
  - Document Date, Due Date and Entry Date
  - Entry User
  - Amount and Balance (with negative values for payments and credit notes)
  - Totals for Amount and Balance

### Summary Sheet

The first sheet provides a dashboard with key information:
  - Company Information (name, business number, address)
  - Contact Information (phone, email, website with hyperlink)
  - Report Data (last data load date, report generation details)
  - Summary Statistics (total debtors, active debtors, total open items)
  - Financial Statistics (total outstanding, overdue, and over 90 days balances)
  - Navigation links to other sheets

### Excel Theming Features

The generated Excel reports include professional styling:

- Corporate branded title section with report date
- Professional blue theme for headers
- Alternating row colors for improved readability
- Color-coded negative values (displayed in red)
- Summary rows with totals and averages
- Automatic column sizing with padding
- Frozen header panes for easier navigation
- Auto-filtering in the Customer List sheet
- Footer with contact information
- Properly formatted currency, date, and percentage values

## Running the Application

1. Ensure you have Java 21 installed
2. Build the application using Gradle:
   ```
   ./gradlew build
   ```
3. Run the application:
   ```
   ./gradlew bootRun
   ```
4. Access the endpoint in your browser or via API client:
   ```
   http://localhost:8080/generate/123
   ```
   Replace `123` with any company ID.

   To specify a custom output path for saving the file to disk:
   ```
   http://localhost:8080/generate/123?outputPath=/path/to/save
   ```
   By default, files are saved to `/users/mathewbroughton` if no path is specified.

## Technologies Used

- Spring Boot 3.2.0
- Apache POI 5.2.3
- Java 21
