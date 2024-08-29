package pack15;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;

import java.io.FileInputStream;
import java.io.IOException;


public class Program2 {
	
	 public static void main(String[] args) {
	        try (FileInputStream fileIn = new FileInputStream("workbook.xls");
	             HSSFWorkbook workbook = new HSSFWorkbook(fileIn)) {

	            // Get the first sheet
	            HSSFSheet sheet = workbook.getSheetAt(0);

	            // Iterate over rows
	            for (Row row : sheet) {
	                // Iterate over cells in each row
	                for (Cell cell : row) {
	                    switch (cell.getCellType()) {
	                        case STRING:
	                            System.out.print(cell.getStringCellValue() + "\t\t\t");
	                            break;
	                        case NUMERIC:
	                            System.out.print(cell.getNumericCellValue() + "\t\t");
	                            break;
	                        case BOOLEAN:
	                            System.out.print(cell.getBooleanCellValue() + "\t");
	                            break;
	                        default:
	                            System.out.print("UNKNOWN\t");
	                    }
	                }
	                System.out.println();
	            }

	        } catch (IOException e) {
	            e.printStackTrace();
	        }

}
}
