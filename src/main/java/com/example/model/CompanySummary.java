package com.example.model;

import java.math.BigDecimal;
import java.time.LocalDateTime;

public class CompanySummary {
    private String companyId;
    private String companyName;
    private String businessNumber;
    private String address;
    private String city;
    private String state;
    private String postalCode;
    private String country;
    private String phone;
    private String email;
    private String website;
    private LocalDateTime lastDataLoadDate;
    private Integer totalDebtors;
    private Integer activeDebtors;
    private Integer totalOpenItems;
    private BigDecimal totalOutstandingBalance;
    private BigDecimal totalOverdueBalance;
    private BigDecimal totalOver90DaysBalance;
    private String reportGeneratedBy;

    public CompanySummary() {
    }

    public CompanySummary(String companyId, String companyName, String businessNumber, String address, 
                         String city, String state, String postalCode, String country, String phone, 
                         String email, String website, LocalDateTime lastDataLoadDate, Integer totalDebtors, 
                         Integer activeDebtors, Integer totalOpenItems, BigDecimal totalOutstandingBalance, 
                         BigDecimal totalOverdueBalance, BigDecimal totalOver90DaysBalance, String reportGeneratedBy) {
        this.companyId = companyId;
        this.companyName = companyName;
        this.businessNumber = businessNumber;
        this.address = address;
        this.city = city;
        this.state = state;
        this.postalCode = postalCode;
        this.country = country;
        this.phone = phone;
        this.email = email;
        this.website = website;
        this.lastDataLoadDate = lastDataLoadDate;
        this.totalDebtors = totalDebtors;
        this.activeDebtors = activeDebtors;
        this.totalOpenItems = totalOpenItems;
        this.totalOutstandingBalance = totalOutstandingBalance;
        this.totalOverdueBalance = totalOverdueBalance;
        this.totalOver90DaysBalance = totalOver90DaysBalance;
        this.reportGeneratedBy = reportGeneratedBy;
    }

    public String getCompanyId() {
        return companyId;
    }

    public void setCompanyId(String companyId) {
        this.companyId = companyId;
    }

    public String getCompanyName() {
        return companyName;
    }

    public void setCompanyName(String companyName) {
        this.companyName = companyName;
    }

    public String getBusinessNumber() {
        return businessNumber;
    }

    public void setBusinessNumber(String businessNumber) {
        this.businessNumber = businessNumber;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public String getCity() {
        return city;
    }

    public void setCity(String city) {
        this.city = city;
    }

    public String getState() {
        return state;
    }

    public void setState(String state) {
        this.state = state;
    }

    public String getPostalCode() {
        return postalCode;
    }

    public void setPostalCode(String postalCode) {
        this.postalCode = postalCode;
    }

    public String getCountry() {
        return country;
    }

    public void setCountry(String country) {
        this.country = country;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public String getWebsite() {
        return website;
    }

    public void setWebsite(String website) {
        this.website = website;
    }

    public LocalDateTime getLastDataLoadDate() {
        return lastDataLoadDate;
    }

    public void setLastDataLoadDate(LocalDateTime lastDataLoadDate) {
        this.lastDataLoadDate = lastDataLoadDate;
    }

    public Integer getTotalDebtors() {
        return totalDebtors;
    }

    public void setTotalDebtors(Integer totalDebtors) {
        this.totalDebtors = totalDebtors;
    }

    public Integer getActiveDebtors() {
        return activeDebtors;
    }

    public void setActiveDebtors(Integer activeDebtors) {
        this.activeDebtors = activeDebtors;
    }

    public Integer getTotalOpenItems() {
        return totalOpenItems;
    }

    public void setTotalOpenItems(Integer totalOpenItems) {
        this.totalOpenItems = totalOpenItems;
    }

    public BigDecimal getTotalOutstandingBalance() {
        return totalOutstandingBalance;
    }

    public void setTotalOutstandingBalance(BigDecimal totalOutstandingBalance) {
        this.totalOutstandingBalance = totalOutstandingBalance;
    }

    public BigDecimal getTotalOverdueBalance() {
        return totalOverdueBalance;
    }

    public void setTotalOverdueBalance(BigDecimal totalOverdueBalance) {
        this.totalOverdueBalance = totalOverdueBalance;
    }

    public BigDecimal getTotalOver90DaysBalance() {
        return totalOver90DaysBalance;
    }

    public void setTotalOver90DaysBalance(BigDecimal totalOver90DaysBalance) {
        this.totalOver90DaysBalance = totalOver90DaysBalance;
    }

    public String getReportGeneratedBy() {
        return reportGeneratedBy;
    }

    public void setReportGeneratedBy(String reportGeneratedBy) {
        this.reportGeneratedBy = reportGeneratedBy;
    }
}
