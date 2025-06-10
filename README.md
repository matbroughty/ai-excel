# Excel Generator Application

This Spring Boot application generates Excel reports containing ageing data by month. The application exposes a REST endpoint that returns an Excel file populated with data from a mocked API.

## Features

- REST endpoint `/generate/{companyId}` that returns an Excel file
- Mock data service that simulates retrieving data from an external API
- Excel report generation using Apache POI
- Report includes monthly ageing data with:
  - Sales Ledger Balance
  - Amount Not Due
  - Amount Over 30 Days Past Due
  - Amount Over 60 Days Past Due
  - Amount Over 90 Days Past Due
  - Amount Over Threshold
  - Total Credits
  - Percentage Over 90 Days Past Due

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

## Technologies Used

- Spring Boot 3.2.0
- Apache POI 5.2.3
- Java 21
