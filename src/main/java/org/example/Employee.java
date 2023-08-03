package org.example;

import java.util.ArrayList;
import java.util.List;

public class Employee {
    private String name;
    private double compareAmount;
    private List<AttendanceDay> attendanceDays = new ArrayList<>();
    private double totalAmount;

    public Employee(){
    }

    public Employee(String name, double compareAmount, List<AttendanceDay> attendanceDays, double totalAmount) {
        this.name = name;
        this.compareAmount = compareAmount;
        this.attendanceDays = new ArrayList<>();
        this.totalAmount = totalAmount;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public double getCompareAmount() {
        return compareAmount;
    }

    public void setCompareAmount(double compareAmount) {
       this.compareAmount= compareAmount;
    }

    public List<AttendanceDay> getAttendanceDays() {
        return attendanceDays;
    }

    public void addAttendanceDay(AttendanceDay attendanceDay) {
        this.attendanceDays.add(attendanceDay);
        totalAmount += attendanceDay.getAmount();
    }

    public double getTotalAmount() {
        return totalAmount;
    }

    public void setTotalAmount(double totalAmount) {
        this.totalAmount = totalAmount;
    }
}
