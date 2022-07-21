// imports
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

// WriteExcel class
public class WriteExcel {
  public static void main(String[] args) {
    // Create blank workbook
    XSSFWorkbook workbook = new XSSFWorkbook();

    // Create blank sheet
    XSSFSheet sheet = workbook.createSheet();

    // This data need to be written (Object [])
    Map<String, Object[]> data = new TreeMap<String, Object[]>();
    data.put("1", new Object[] {"ID", "NAME", "LastName"});
    data.put("2", new Object[] {1, "Amit", "Shukla"});
    data.put("3", new Object[] {2, "Rian", "Gupta"});
    data.put("4", new Object[] {3, "John", "Edwards"});
    data.put("5", new Object[] {4, "Brian", "Schultz"});

    // Iterate over data and write to sheet
    Set<String> keySet = data.keySet();
    int rowNum = 0;
    for (String key : keySet) {
      Row row = sheet.createRow(rowNum++);
      Object [] objArr = data.get(key);
      int cellNum = 0;
      for (Object obj : objArr) {
        Cell cell = row.createCell(cellNum++);
        if (obj instanceof String) {
          cell.setCellValue((String)obj);
        } else if (obj instanceof Integer) {
          cell.setCellValue((Integer)obj);
        }
      }
    }
    try {
      // Write the workbook in file system
      FileOutputStream out = new FileOutputStream(new File("ExcelPoiTutorial.xlsx"));
      workbook.write(out);
      System.out.println("ExcelPoiTutorial.xlsx written successfully on disk.");
    } catch (Exception e) {
      e.printStackTrace();
    }
  }
}