package com.example.model;

import java.math.BigDecimal;
import java.time.LocalDate;

public class Item {
    private String customerId;
    private BigDecimal amount;
    private BigDecimal balance;
    private LocalDate documentDate;
    private LocalDate dueDate;
    private LocalDate entryDate;
    private String entryUser;
    private String documentNumber;
    private String documentReference;
    private String itemType; // INV, PAY, or CRN

    public Item() {
    }

    public Item(String customerId, BigDecimal amount, BigDecimal balance, 
                LocalDate documentDate, LocalDate dueDate, LocalDate entryDate, 
                String entryUser, String documentNumber, String documentReference, 
                String itemType) {
        this.customerId = customerId;
        this.amount = amount;
        this.balance = balance;
        this.documentDate = documentDate;
        this.dueDate = dueDate;
        this.entryDate = entryDate;
        this.entryUser = entryUser;
        this.documentNumber = documentNumber;
        this.documentReference = documentReference;
        this.itemType = itemType;
    }

    public String getCustomerId() {
        return customerId;
    }

    public void setCustomerId(String customerId) {
        this.customerId = customerId;
    }

    public BigDecimal getAmount() {
        return amount;
    }

    public void setAmount(BigDecimal amount) {
        this.amount = amount;
    }

    public BigDecimal getBalance() {
        return balance;
    }

    public void setBalance(BigDecimal balance) {
        this.balance = balance;
    }

    public LocalDate getDocumentDate() {
        return documentDate;
    }

    public void setDocumentDate(LocalDate documentDate) {
        this.documentDate = documentDate;
    }

    public LocalDate getDueDate() {
        return dueDate;
    }

    public void setDueDate(LocalDate dueDate) {
        this.dueDate = dueDate;
    }

    public LocalDate getEntryDate() {
        return entryDate;
    }

    public void setEntryDate(LocalDate entryDate) {
        this.entryDate = entryDate;
    }

    public String getEntryUser() {
        return entryUser;
    }

    public void setEntryUser(String entryUser) {
        this.entryUser = entryUser;
    }

    public String getDocumentNumber() {
        return documentNumber;
    }

    public void setDocumentNumber(String documentNumber) {
        this.documentNumber = documentNumber;
    }

    public String getDocumentReference() {
        return documentReference;
    }

    public void setDocumentReference(String documentReference) {
        this.documentReference = documentReference;
    }

    public String getItemType() {
        return itemType;
    }

    public void setItemType(String itemType) {
        this.itemType = itemType;
    }

    /**
     * Convenience method to check if this is an invoice
     */
    public boolean isInvoice() {
        return "INV".equals(itemType);
    }

    /**
     * Convenience method to check if this is a payment
     */
    public boolean isPayment() {
        return "PAY".equals(itemType);
    }

    /**
     * Convenience method to check if this is a credit note
     */
    public boolean isCreditNote() {
        return "CRN".equals(itemType);
    }
}
