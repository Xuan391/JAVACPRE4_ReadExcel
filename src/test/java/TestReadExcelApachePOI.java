import org.example.AttendanceDay;
import org.example.Employee;
import org.example.ReadExcelApachePOI;
import org.junit.Assert;
import org.junit.Test;

import java.io.IOException;
import java.util.List;

public class TestReadExcelApachePOI {
    private static final String TEST_FILE_PATH1 = "src/test/excel/BangCongTest.xlsx";
    private static final String TEST_FILE_PATH2 = "src/test/excel/BangCongTest_2.xlsx";
    private static final String TEST_FILE_PATH3 = "src/test/excel/BangCongTest_3.xlsx";

    private List<Employee> employees;

    @Test
    public void testEmployeeDetails1() {
        try {
            // Thực hiện phân tích dữ liệu từ tệp Excel test và lưu kết quả vào danh sách nhân viên
            employees = ReadExcelApachePOI.analyzeAttendance(TEST_FILE_PATH1);
            // Kiểm tra thông tin của nhân viên trong danh sách
            Employee firstEmployee = employees.get(0);
            Assert.assertEquals("Nguyen Van A", firstEmployee.getName());
            Assert.assertEquals(firstEmployee.getTotalAmount(), firstEmployee.getCompareAmount(), 0.01);
            Assert.assertEquals("1/12/2021", firstEmployee.getAttendanceDays().get(0).getDate());
            Assert.assertEquals("2/12/2021", firstEmployee.getAttendanceDays().get(1).getDate());
            Assert.assertEquals(8.0, firstEmployee.getAttendanceDays().get(0).getHours(), 0.001);
            Assert.assertEquals(8.0, firstEmployee.getAttendanceDays().get(1).getHours(), 0.001);
            Assert.assertEquals(1730769.2307, firstEmployee.getAttendanceDays().get(0).getAmount(), 0.0001);
            Assert.assertEquals(1730769.2307, firstEmployee.getAttendanceDays().get(1).getAmount(), 0.0001);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void testEmployeeDetailsWithExcelFileAddShift() {
        try {
            // Thực hiện phân tích dữ liệu từ tệp Excel test và lưu kết quả vào danh sách nhân viên
            employees = ReadExcelApachePOI.analyzeAttendance(TEST_FILE_PATH2);
            // Kiểm tra thông tin của nhân viên trong danh sách
            Employee firstEmployee = employees.get(0);
            Assert.assertEquals("Nguyen Van A", firstEmployee.getName());
            Assert.assertEquals(firstEmployee.getTotalAmount(), firstEmployee.getCompareAmount(), 0.01);
            Assert.assertEquals("1/12/2021", firstEmployee.getAttendanceDays().get(0).getDate());
            Assert.assertEquals(12.0, firstEmployee.getAttendanceDays().get(0).getHours(), 0.001);
            Assert.assertEquals(2788461.5384, firstEmployee.getAttendanceDays().get(0).getAmount(), 0.0001);
            Assert.assertEquals("GC", firstEmployee.getAttendanceDays().get(0).getShifts().get(0));
            Assert.assertEquals("TC1", firstEmployee.getAttendanceDays().get(0).getShifts().get(1));
            Assert.assertEquals("TC2", firstEmployee.getAttendanceDays().get(0).getShifts().get(2));

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void testEmployeeDetailsWithExcelFileFirstDayEqual10() {
        try {
            // Thực hiện phân tích dữ liệu từ tệp Excel test và lưu kết quả vào danh sách nhân viên
            employees = ReadExcelApachePOI.analyzeAttendance(TEST_FILE_PATH3);
            // Kiểm tra thông tin của nhân viên trong danh sách
            Employee firstEmployee = employees.get(0);
            Assert.assertEquals("Nguyen Van A", firstEmployee.getName());
            Assert.assertEquals(firstEmployee.getTotalAmount(), firstEmployee.getCompareAmount(), 0.01);
            Assert.assertEquals("10/12/2021", firstEmployee.getAttendanceDays().get(0).getDate());
            Assert.assertEquals("31/12/2021", firstEmployee.getAttendanceDays().get(1).getDate());
            Assert.assertEquals("1/1/2022", firstEmployee.getAttendanceDays().get(2).getDate());
            Assert.assertEquals(8.0, firstEmployee.getAttendanceDays().get(0).getHours(), 0.001);
            Assert.assertEquals(8.0, firstEmployee.getAttendanceDays().get(1).getHours(), 0.001);
            Assert.assertEquals(4.0, firstEmployee.getAttendanceDays().get(2).getHours(), 0.001);
            Assert.assertEquals("GC", firstEmployee.getAttendanceDays().get(0).getShifts().get(0));
            Assert.assertEquals("GC", firstEmployee.getAttendanceDays().get(1).getShifts().get(0));
            Assert.assertEquals("GC", firstEmployee.getAttendanceDays().get(2).getShifts().get(0));

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
