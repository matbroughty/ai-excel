package com.example.service;

import com.example.model.Customer;
import com.example.model.Item;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;
import java.util.Set;
import java.util.stream.Collectors;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

@Service
public class ItemService {

    private static final Logger logger = LoggerFactory.getLogger(ItemService.class);
    private final Random random = new Random();
    private final CustomerService customerService;

    private final String[] users = {
            "john.doe", "jane.smith", "robert.johnson", "sarah.williams", "michael.brown",
            "lisa.jones", "david.miller", "emily.davis", "james.wilson", "amanda.taylor"
    };

    @Autowired
    public ItemService(CustomerService customerService) {
        this.customerService = customerService;
    }

    /**
     * Mock service that would normally call an external API to get open items data
     * @param companyId the company ID to retrieve items for
     * @return List of open items for customers with IDs matching those from CustomerService
     */
    public List<Item> getOpenItems(String companyId) {
        // In a real application, this would call the external endpoint
        // someurl/accounting/companies/{companyId}/items

        List<Item> items = new ArrayList<>();

        // Get customer IDs from CustomerService
        List<Customer> customers = customerService.getCustomersWithOutstandingBalance(companyId);
        List<String> customerIds = customers.stream()
                .map(Customer::getCustomerId)
                .collect(Collectors.toList());

        // If no customers found, return empty list
        if (customerIds.isEmpty()) {
            logger.warn("No customers found for company ID: {}", companyId);
            return items;
        }

        // Log retrieved customer IDs for debugging
        logger.info("Retrieved {} customers for company ID: {}", customerIds.size(), companyId);

        // Generate a random number of items (between 50 and 150)
        // but make sure we don't generate more items than we have customers * 5
        // to ensure a reasonable distribution
        int maxItems = Math.min(150, customerIds.size() * 5);
        int minItems = Math.min(50, customerIds.size() * 2);
        int numItems = minItems + (maxItems > minItems ? random.nextInt(maxItems - minItems) : 0);

        logger.info("Generating {} items for {} customers", numItems, customerIds.size());
        
        // Generate the requested number of items with random data
        for (int i = 0; i < numItems; i++) {
            // Randomly select a customer ID from the list obtained from CustomerService
            String customerId = customerIds.get(random.nextInt(customerIds.size()));
            
            // Determine item type (60% invoices, 20% payments, 20% credit notes)
            String itemType;
            int typeRandom = random.nextInt(100);
            if (typeRandom < 60) {
                itemType = "INV";
            } else if (typeRandom < 80) {
                itemType = "PAY";
            } else {
                itemType = "CRN";
            }
            
            // Generate amount based on item type (payments and credit notes are negative)
            BigDecimal amount;
            if ("INV".equals(itemType)) {
                // Invoices: positive amounts between $10 and $10,000
                amount = new BigDecimal(10 + random.nextInt(9990) + random.nextDouble()).setScale(2, java.math.RoundingMode.HALF_UP);
            } else {
                // Payments and credit notes: negative amounts
                amount = new BigDecimal(-10 - random.nextInt(9990) - random.nextDouble()).setScale(2, java.math.RoundingMode.HALF_UP);
            }
            
            // Generate balance (for partially paid invoices, the balance might be less than the amount)
            BigDecimal balance;
            if ("INV".equals(itemType)) {
                // For invoices, balance is between 0 and the full amount
                if (random.nextInt(100) < 30) {
                    // 30% chance of partially paid invoice
                    double ratio = random.nextDouble();
                    balance = amount.multiply(new BigDecimal(ratio)).setScale(2, java.math.RoundingMode.HALF_UP);
                } else {
                    // 70% chance of unpaid invoice (balance = amount)
                    balance = amount;
                }
            } else {
                // For payments and credit notes, balance is the same as amount
                balance = amount;
            }
            
            // Generate document date (between 1 and 180 days ago)
            LocalDate documentDate = LocalDate.now().minusDays(1 + random.nextInt(180));
            
            // Generate due date (for invoices, usually 30 days after document date)
            LocalDate dueDate;
            if ("INV".equals(itemType)) {
                dueDate = documentDate.plusDays(30);
            } else {
                // For payments and credit notes, due date is same as document date
                dueDate = documentDate;
            }
            
            // Generate entry date (typically same day or a few days after document date)
            LocalDate entryDate = documentDate.plusDays(random.nextInt(3));
            
            // Generate entry user
            String entryUser = users[random.nextInt(users.length)];
            
            // Generate document number
            String documentPrefix = "INV".equals(itemType) ? "INV" : ("PAY".equals(itemType) ? "PMT" : "CRN");
            String documentNumber = documentPrefix + "-" + String.format("%08d", random.nextInt(100000000));
            
            // Generate document reference
            String documentReference = "REF-" + String.format("%06d", random.nextInt(1000000));
            
            items.add(new Item(
                    customerId, amount, balance, documentDate, dueDate, entryDate,
                    entryUser, documentNumber, documentReference, itemType
            ));
        }
        
        return items;
    }
}
