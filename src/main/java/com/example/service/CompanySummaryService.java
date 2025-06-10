package com.example.service;

import com.example.model.CompanySummary;
import org.springframework.stereotype.Service;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.Random;

@Service
public class CompanySummaryService {

    private final Random random = new Random();

    private final String[] companyNames = {
            "Acme Financial Services", "Global Credit Solutions", "Pinnacle Banking Corp",
            "Summit Lending Group", "Atlas Credit Management", "Vanguard Financial",
            "Everest Capital Management", "Horizon Credit Services", "Precision Financial Group",
            "Liberty Finance Solutions"
    };

    private final String[] cities = {
            "New York", "London", "Sydney", "Toronto", "Singapore",
            "Hong Kong", "Dubai", "Los Angeles", "Chicago", "Boston"
    };

    private final String[] states = {
            "NY", "CA", "TX", "FL", "IL", "PA", "ON", "QC", "NSW", "VIC"
    };

    private final String[] countries = {
            "United States", "United Kingdom", "Australia", "Canada", "Singapore"
    };

    /**
     * Mock service that would normally call an external API to get company summary data
     * In a real application, this would call: someurl/summary/{companyId}
     * 
     * @param companyId the company ID to retrieve summary for
     * @return CompanySummary object with company information and summary data
     */
    public CompanySummary getCompanySummary(String companyId) {
        // In a real service, this would call the external API endpoint

        // Generate random company data
        String companyName = companyNames[random.nextInt(companyNames.length)];
        String businessNumber = "BN" + String.format("%09d", random.nextInt(1000000000));

        // Generate address information
        String address = (random.nextInt(999) + 1) + " " + 
                new String[]{"Main", "First", "Financial", "Commerce", "Market"}[random.nextInt(5)] + " " +
                new String[]{"Street", "Avenue", "Boulevard", "Plaza", "Tower"}[random.nextInt(5)];

        String city = cities[random.nextInt(cities.length)];
        String state = states[random.nextInt(states.length)];
        String postalCode = String.format("%05d", random.nextInt(100000));
        String country = countries[random.nextInt(countries.length)];

        // Generate contact information
        String phone = "+" + (random.nextInt(20) + 1) + " " + 
                (random.nextInt(900) + 100) + "-" + 
                (random.nextInt(900) + 100) + "-" + 
                (random.nextInt(9000) + 1000);

        String email = "finance@" + companyName.toLowerCase().replaceAll("[^a-z]", "") + ".com";
        String website = "www." + companyName.toLowerCase().replaceAll("[^a-z]", "") + ".com";

        // Generate last data load date (within the last 10 days)
        LocalDateTime lastDataLoadDate = LocalDateTime.now().minusDays(random.nextInt(10));

        // Generate summary statistics
        int totalDebtors = 50 + random.nextInt(950);
        int activeDebtors = (int)(totalDebtors * (0.4 + (random.nextDouble() * 0.5))); // 40% to 90% active
        int totalOpenItems = activeDebtors * (2 + random.nextInt(6)); // 2-7 items per active debtor

        // Generate financial summaries
        BigDecimal totalOutstandingBalance = new BigDecimal(activeDebtors * (1000 + random.nextInt(9000)))
                .setScale(2, java.math.RoundingMode.HALF_UP);
        BigDecimal totalOverdueBalance = totalOutstandingBalance.multiply(new BigDecimal(0.2 + (random.nextDouble() * 0.5)))
                .setScale(2, java.math.RoundingMode.HALF_UP); // 20% to 70% overdue
        BigDecimal totalOver90DaysBalance = totalOverdueBalance.multiply(new BigDecimal(0.1 + (random.nextDouble() * 0.4)))
                .setScale(2, java.math.RoundingMode.HALF_UP); // 10% to 50% of overdue is >90 days

        String reportGeneratedBy = "System";

        return new CompanySummary(
                companyId, companyName, businessNumber, address,
                city, state, postalCode, country, phone,
                email, website, lastDataLoadDate, totalDebtors,
                activeDebtors, totalOpenItems, totalOutstandingBalance,
                totalOverdueBalance, totalOver90DaysBalance, reportGeneratedBy
        );
    }
}
