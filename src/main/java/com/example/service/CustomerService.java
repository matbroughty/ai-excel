package com.example.service;

import com.example.model.Customer;
import org.springframework.stereotype.Service;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

@Service
public class CustomerService {

    private final Random random = new Random();
    private final String[] companyNames = {
            "Acme Corp", "Globex Corporation", "Soylent Corp", "Initech", "Umbrella Corporation",
            "Stark Industries", "Wayne Enterprises", "Cyberdyne Systems", "Weyland-Yutani Corp",
            "Massive Dynamic", "Oceanic Airlines", "Oscorp Industries", "Rekall", "Wonka Industries",
            "Tyrell Corporation", "Gekko & Co", "Nakatomi Trading Corp", "Dunder Mifflin", "Bluth Company",
            "Slate Rock and Gravel Company", "Sterling Cooper", "Wernham Hogg", "Sirius Cybernetics Corp"
    };
    
    private final String[] cities = {
            "New York", "London", "Tokyo", "Paris", "Sydney", "Berlin", "Toronto", "Singapore",
            "Hong Kong", "Dubai", "Los Angeles", "Chicago", "Boston", "San Francisco", "Seattle"
    };
    
    private final String[] states = {
            "NY", "CA", "TX", "FL", "IL", "PA", "OH", "GA", "NC", "MI", "NJ", "VA", "WA", "MA", "AZ"
    };
    
    private final String[] countries = {
            "United States", "United Kingdom", "Canada", "Australia", "Germany", "Japan", "France",
            "Italy", "Spain", "China", "India", "Brazil", "Mexico", "Singapore", "South Korea"
    };

    /**
     * Mock service that would normally call an external API to get customer data
     * @param companyId the company ID to retrieve customers for
     * @return List of customers with outstanding balances
     */
    public List<Customer> getCustomersWithOutstandingBalance(String companyId) {
        // In a real application, this would call the external endpoint
        // someurl/accounting/companies/{companyId}/customers
        
        List<Customer> customers = new ArrayList<>();
        
        // Generate a random number of customers (between 10 and 30)
        int numCustomers = 10 + random.nextInt(21);
        
        // Generate the requested number of customers with random data
        for (int i = 0; i < numCustomers; i++) {
            String customerId = "CUST" + String.format("%06d", random.nextInt(1000000));
            String customerName = companyNames[random.nextInt(companyNames.length)];
            
            // Generate balance (positive values only, as these are outstanding balances)
            BigDecimal balance = new BigDecimal(100 + random.nextInt(1000000) / 100.0);
            
            String reference = "REF" + String.format("%08d", random.nextInt(100000000));
            
            // Generate address information
            String addressLine1 = (random.nextInt(999) + 1) + " " + 
                    new String[]{"Main", "First", "Oak", "Pine", "Maple", "Cedar", "Broadway"}[random.nextInt(7)] + " " +
                    new String[]{"Street", "Avenue", "Boulevard", "Road", "Lane", "Drive", "Way"}[random.nextInt(7)];
            
            String addressLine2 = random.nextBoolean() ? 
                    "Suite " + (random.nextInt(100) + 100) : 
                    "";
            
            String city = cities[random.nextInt(cities.length)];
            String state = states[random.nextInt(states.length)];
            String postalCode = String.format("%05d", random.nextInt(100000));
            String country = countries[random.nextInt(countries.length)];
            
            // Generate notification status (most customers have been notified)
            boolean notified = random.nextInt(100) < 80; // 80% are notified
            
            // Generate last updated date (within the last 90 days)
            LocalDate lastUpdated = LocalDate.now().minusDays(random.nextInt(90));
            
            customers.add(new Customer(
                    customerId, customerName, balance, reference,
                    addressLine1, addressLine2, city, state,
                    postalCode, country, notified, lastUpdated
            ));
        }
        
        return customers;
    }
}
