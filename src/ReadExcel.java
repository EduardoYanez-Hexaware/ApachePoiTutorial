import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

public class ReadExcel {
  public static void main(String[] args) {
    try {
      FileInputStream file = new FileInputStream(new File("ExcelPoiTutorial.xlsx"));

      // Create workbook instance holing reference to .xlsx file
      XSSFWorkbook workbook = new XSSFWorkbook(file);

      // Get first/desired sheet from the workbook
      XSSFSheet sheet = workbook.getSheetAt(0);

      // Iterate through each row one by one
      Iterator<Row> rowIterator = sheet.iterator();

      while (rowIterator.hasNext()) {
        Row row = rowIterator.next();
        // For each row, iterate through all the columns
        Iterator<Cell> cellIterator = row.cellIterator();

        while (cellIterator.hasNext()) {
          Cell cell = cellIterator.next();
          // Check the cell type and format accordingly
          switch (cell.getCellType()) {
            case NUMERIC:
              System.out.println(cell.getNumericCellValue() + "N");
              break;
            case STRING:
              System.out.println(cell.getStringCellValue() + "S");
              break;
          }
        }
        System.out.println("");
      }
      file.close();
    } catch (Exception e) {
      e.printStackTrace();
    }
  }
}
