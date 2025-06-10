package com.example.service;

import com.example.model.AgeingByMonth;
import org.springframework.stereotype.Service;

import java.math.BigDecimal;
import java.time.YearMonth;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

@Service
public class AgeingDataService {

    private final Random random = new Random();

    /**
     * Mock service that would normally call an external API
     * @param companyId the company ID to retrieve data for
     * @return List of ageing data by month
     */
    public List<AgeingByMonth> getAgeingByMonth(String companyId) {
        // In a real application, this would call the external endpoint
        // someurl/lending/companies/{companyId}/ageingByMonth
        
        List<AgeingByMonth> result = new ArrayList<>();
        YearMonth currentMonth = YearMonth.now();
        
        // Generate 12 months of data
        for (int i = 0; i < 12; i++) {
            YearMonth month = currentMonth.minusMonths(i);
            
            // Generate random values (in a real app, these would come from the API)
            BigDecimal salesLedgerBalance = new BigDecimal(random.nextInt(1000000) / 100.0);
            BigDecimal amountNotDue = new BigDecimal(random.nextInt(500000) / 100.0);
            BigDecimal amountOver30Days = new BigDecimal(random.nextInt(300000) / 100.0);
            BigDecimal amountOver60Days = new BigDecimal(random.nextInt(200000) / 100.0);
            BigDecimal amountOver90Days = new BigDecimal(random.nextInt(100000) / 100.0);
            BigDecimal amountOverThreshold = new BigDecimal(random.nextInt(50000) / 100.0);
            BigDecimal totalCredits = new BigDecimal(random.nextInt(200000) / 100.0);
            
            result.add(new AgeingByMonth(
                month,
                salesLedgerBalance,
                amountNotDue,
                amountOver30Days,
                amountOver60Days,
                amountOver90Days,
                amountOverThreshold,
                totalCredits
            ));
        }
        
        return result;
    }
}
