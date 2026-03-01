package ru.objectsfill.dto;

public class ExcelImport {

    private String account;
    private String offerDate;
    private String currency;
    private String product;
    private String rateModNumber;

    public String getAccount() {
        return account;
    }

    public void setAccount(String account) {
        this.account = account;
    }

    public String getOfferDate() {
        return offerDate;
    }

    public void setOfferDate(String offerDate) {
        this.offerDate = offerDate;
    }

    public String getCurrency() {
        return currency;
    }

    public void setCurrency(String currency) {
        this.currency = currency;
    }

    public String getProduct() {
        return product;
    }

    public void setProduct(String product) {
        this.product = product;
    }

    public String getRateModNumber() {
        return rateModNumber;
    }

    public void setRateModNumber(String rateModNumber) {
        this.rateModNumber = rateModNumber;
    }
}
