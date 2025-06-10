package com.example.model;

import java.math.BigDecimal;
import java.time.YearMonth;

public class AgeingByMonth {
    private YearMonth month;
    private BigDecimal salesLedgerBalance;
    private BigDecimal amountNotDue;
    private BigDecimal amountOver30Days;
    private BigDecimal amountOver60Days;
    private BigDecimal amountOver90Days;
    private BigDecimal amountOverThreshold;
    private BigDecimal totalCredits;
    private BigDecimal percentOver90Days;

    public AgeingByMonth() {
    }

    public AgeingByMonth(YearMonth month, BigDecimal salesLedgerBalance, BigDecimal amountNotDue, 
                        BigDecimal amountOver30Days, BigDecimal amountOver60Days, BigDecimal amountOver90Days, 
                        BigDecimal amountOverThreshold, BigDecimal totalCredits) {
        this.month = month;
        this.salesLedgerBalance = salesLedgerBalance;
        this.amountNotDue = amountNotDue;
        this.amountOver30Days = amountOver30Days;
        this.amountOver60Days = amountOver60Days;
        this.amountOver90Days = amountOver90Days;
        this.amountOverThreshold = amountOverThreshold;
        this.totalCredits = totalCredits;
        
        // Calculate percentage over 90 days
        if (salesLedgerBalance.compareTo(BigDecimal.ZERO) > 0) {
            this.percentOver90Days = amountOver90Days.multiply(new BigDecimal("100")).divide(salesLedgerBalance, 2, BigDecimal.ROUND_HALF_UP);
        } else {
            this.percentOver90Days = BigDecimal.ZERO;
        }
    }

    public YearMonth getMonth() {
        return month;
    }

    public void setMonth(YearMonth month) {
        this.month = month;
    }

    public BigDecimal getSalesLedgerBalance() {
        return salesLedgerBalance;
    }

    public void setSalesLedgerBalance(BigDecimal salesLedgerBalance) {
        this.salesLedgerBalance = salesLedgerBalance;
    }

    public BigDecimal getAmountNotDue() {
        return amountNotDue;
    }

    public void setAmountNotDue(BigDecimal amountNotDue) {
        this.amountNotDue = amountNotDue;
    }

    public BigDecimal getAmountOver30Days() {
        return amountOver30Days;
    }

    public void setAmountOver30Days(BigDecimal amountOver30Days) {
        this.amountOver30Days = amountOver30Days;
    }

    public BigDecimal getAmountOver60Days() {
        return amountOver60Days;
    }

    public void setAmountOver60Days(BigDecimal amountOver60Days) {
        this.amountOver60Days = amountOver60Days;
    }

    public BigDecimal getAmountOver90Days() {
        return amountOver90Days;
    }

    public void setAmountOver90Days(BigDecimal amountOver90Days) {
        this.amountOver90Days = amountOver90Days;
    }

    public BigDecimal getAmountOverThreshold() {
        return amountOverThreshold;
    }

    public void setAmountOverThreshold(BigDecimal amountOverThreshold) {
        this.amountOverThreshold = amountOverThreshold;
    }

    public BigDecimal getTotalCredits() {
        return totalCredits;
    }

    public void setTotalCredits(BigDecimal totalCredits) {
        this.totalCredits = totalCredits;
    }

    public BigDecimal getPercentOver90Days() {
        return percentOver90Days;
    }

    public void setPercentOver90Days(BigDecimal percentOver90Days) {
        this.percentOver90Days = percentOver90Days;
    }
}
