package pack15;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;

import java.io.FileOutputStream;
import java.io.IOException;

public class Program1 {
	
	public static void main(String[] args) {
		try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            // Create a new sheet named "Sheet1"
            HSSFSheet sheet = workbook.createSheet("Sheet1");

            // Create the header row
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Name");
            headerRow.createCell(1).setCellValue("Age");
            headerRow.createCell(2).setCellValue("Email");

            // Data to be written
            Object[][] data = {
                {"John Deo", 30, "john@test.com"},
                {"Jane Deo", 28, "jane@test.com"},
                {"Bob Smith", 35, "jacky@example.com"},
                {"Swapnil", 37, "swapnil@example.com"}
            };

            // Write data to the sheet
            int rowNum = 1;
            for (Object[] rowData : data) {
                Row row = sheet.createRow(rowNum++);
                for (int i = 0; i < rowData.length; i++) {
                    Cell cell = row.createCell(i);
                    if (rowData[i] instanceof String) {
                        cell.setCellValue((String) rowData[i]);
                    } else if (rowData[i] instanceof Integer) {
                        cell.setCellValue((Integer) rowData[i]);
                    }
                }
            }

            // Write the workbook to a file
            try (FileOutputStream fileOut = new FileOutputStream("workbook.xls")) {
                workbook.write(fileOut);
            }

            System.out.println("Workbook with data created successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
	}

}
