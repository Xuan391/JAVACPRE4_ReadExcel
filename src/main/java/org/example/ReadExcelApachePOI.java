package org.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ReadExcelApachePOI {

    public static void main(String[] args) throws FileNotFoundException {
        String filePath = "excel/BangCong_3.xlsx";
        try {
            List<Employee> employees = analyzeAttendance(filePath);
            System.out.println(employees);

             //Hiển thị kết quả
            for (Employee employee : employees) {
                System.out.println("Employee: " + employee.getName());
                System.out.println("Working days:");
                for (AttendanceDay attendanceDay : employee.getAttendanceDays()) {
                    System.out.println("Date: " + attendanceDay.getDate());
                    System.out.println("Total hours: " + attendanceDay.getHours());
                    System.out.println("Shifts: " + attendanceDay.getShifts());
                    System.out.println("day amount: "+attendanceDay.getAmount());
                }
                System.out.println("Total amount: " + employee.getTotalAmount());
                System.out.println("Compare amount: " + employee.getCompareAmount());
                System.out.println("Total compared to column Q: " + compareDouble(employee.getCompareAmount(),employee.getTotalAmount(),0.01)  );
                System.out.println("------------------------");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static List<Employee> analyzeAttendance(String filePath) throws IOException {
        List<Employee> employees = new ArrayList<>();

        FileInputStream fis = new FileInputStream(new File(filePath));
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0); // Trang tính đầu tiên trong tệp Excel

        Map<Integer, Integer> map1 = new HashMap<>(); // lưu ngày , index cột
        Map<Integer,String> map2 = new HashMap<>(); // lưu index cột, shift name
        Map<String, Integer> map3 = new HashMap<>(); // lưu shift name, cột chỉ giá tiền

        boolean isFirstDate = true;
        int index = 0; // đánh dấu số thứ tự chỉ cột đầu tiên trong tháng.
        int indexQ = 0; // đánh dấu số thứ tự chỉ cột tổng lương
        // Loop1 : row 3
        Row row3 = sheet.getRow(3);
        int day = 0;
        for(int i=0; i<=row3.getLastCellNum(); i++){
            Cell cell = row3.getCell(i);
            if(cell!=null && cell.getCellTypeEnum() == CellType.NUMERIC){
                day = (int) cell.getNumericCellValue();
                if(isFirstDate == true){
                    index = i;
                    indexQ = index - 1;
                    isFirstDate = false;
                }
            }
            map1.put(i,day);
        }
        int firstDay = map1.get(index);
//        System.out.println(map1);

        //Loop2 :row5
        Row row5 = sheet.getRow(5);
        String idShift ="";
        for(int i = index; i<= row5.getLastCellNum(); i++ ){
            Cell cell = row5.getCell(i);
            if(cell != null && cell.getCellTypeEnum() == CellType.STRING){
                idShift = cell.getStringCellValue();
            }
            map2.put(i,idShift);
        }
//        System.out.println(map2);

        //Loop3 : row5
        List<String> shifts = new ArrayList<>();
        for (int i=0; i< index ; i++){
            Cell cell = row5.getCell(i);
                if(cell != null && cell.getCellTypeEnum() == CellType.STRING){
                    if(cell.getStringCellValue().equals("$")){
                        new ArrayList<>(shifts);
                        for (String shift: shifts){
                            map3.put(shift,i);
                        }
                        shifts.clear();
                    }else{
                        String shift = cell.getStringCellValue();
                        shifts.add(shift);
                    }
                }
        }
//        System.out.println(map3);

        //loop4: duyet nhan vien
        for (int i =6; i<=9; i++){

            Row row6 = sheet.getRow(0);
            int month = (int) row6.getCell(0).getNumericCellValue();
            int year = (int) row6.getCell(1).getNumericCellValue();
//            System.out.println(month+"/"+year);

            Row row = sheet.getRow(i);
            String name = row.getCell(2).getStringCellValue();
            double compareAmount = row.getCell(indexQ).getNumericCellValue();
            Employee employee = new Employee();
            employee.setName(name);
            employee.setCompareAmount(compareAmount);
            AttendanceDay attendanceDay = null;
            int currentDate = -1; // Biến để lưu trữ ngày hiện tại, khởi tạo bằng giá trị không hợp lệ
            double hour = 0.0;
            double amount = 0.0;
            List<String> dayShifts = new ArrayList<>();
            boolean hasEncounteredFirstDay = false;
            for (int j = index; j <= row.getLastCellNum(); j++) {
                int indexDay = map1.get(j);
                if (firstDay != 1) {
                    if (indexDay == 1 && !hasEncounteredFirstDay) {
                        if (month == 12) {
                            month = 1;
                            year++;
                        } else {
                            month++;
                        }
                        hasEncounteredFirstDay = true;
                    }
                }
                Cell cell = row.getCell(j);
                if(cell!= null) {
                    if (cell.getCellTypeEnum() == CellType.NUMERIC && cell.getNumericCellValue() > 0.0) {
                        int date = map1.get(j); //---
                        if (date != currentDate) {
                            if (attendanceDay != null) {
                                employee.addAttendanceDay(attendanceDay);
                            }
                            hour = 0.0;
                            amount = 0.0;
                            dayShifts = new ArrayList<>();
                            attendanceDay = new AttendanceDay();
                        }

                        String shift = map2.get(j);
                        dayShifts.add(shift); //---
                        double hourCell = cell.getNumericCellValue();
                        hour += hourCell; //---

                        if (map3.containsKey(shift)) {
                            int indexMoney = map3.get(shift);
                            double money = row.getCell(indexMoney).getNumericCellValue();
                            amount += hourCell * money; //---
                        }

                        attendanceDay.setDate(date+"/"+month+"/"+year);
                        attendanceDay.setHours(hour);
                        attendanceDay.setAmount(amount);
                        attendanceDay.setShifts(dayShifts);

                        currentDate = date;

                    }
                }

            }

            // Sau khi kết thúc vòng lặp, thêm attendanceDay cuối cùng vào employee
            if (attendanceDay != null) {
                employee.addAttendanceDay(attendanceDay);
            }
            employees.add(employee);
        }

        workbook.close();
        fis.close();

        return employees;
    }

    public static boolean compareDouble (double num1, double num2, double epsilon){
        return Math.abs(num1-num2) < epsilon;
    }
}



