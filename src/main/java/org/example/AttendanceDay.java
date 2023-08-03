package org.example;

import java.util.Date;
import java.util.List;

public class AttendanceDay {
    private String date;
    private double hours;
    private List<String> shifts;
    private double amount;

    public AttendanceDay() {
    }

    public AttendanceDay(String date, double hours, List<String> shifts, double amount) {
        this.date = date;
        this.hours = hours;
        this.shifts = shifts;
        this.amount =amount;
    }

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }

    public double getHours() {
        return hours;
    }

    public void setHours(double hours) {
        this.hours = hours;
    }

    public List<String> getShifts() {
        return shifts;
    }

    public void setShifts(List<String> shifts) {
        this.shifts = shifts;
    }

    public double getAmount() {
        return amount;
    }

    public void setAmount(double amount) {
        this.amount = amount;
    }
}
