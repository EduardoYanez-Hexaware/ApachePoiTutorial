import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;

public class FormulaExcel {
  public static void main(String[] args) {
    // Create workbook
    writeFormulaSheet();
    // Console log formula data
    readFormulaSheet();
  }

  public static void writeFormulaSheet() {
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet sheet = workbook.createSheet("Calculate simple interest");

    Row header = sheet.createRow(0);
    header.createCell(0).setCellValue("Principal");
    header.createCell(1).setCellValue("RoI");
    header.createCell(2).setCellValue("T");
    header.createCell(3).setCellValue("Interest (P r t)");

    Row dataRow = sheet.createRow(1);
    dataRow.createCell(0).setCellValue(14500d);
    dataRow.createCell(1).setCellValue(9.25);
    dataRow.createCell(2).setCellValue(3d);
    dataRow.createCell(3).setCellFormula("A2 * B2 * C2");

    try {
      FileOutputStream out = new FileOutputStream(new File("FormulaExcel.xlsx"));
      workbook.write(out);
      out.close();
      System.out.println("Excel with formula cells written successfully.");
    } catch (FileNotFoundException e) {
      e.printStackTrace();
    } catch (IOException e) {
      e.printStackTrace();
    }
  }

  public static void readFormulaSheet() {
    try {
      FileInputStream file = new FileInputStream(new File("FormulaExcel.xlsx"));

      // Create workbook instance holding reference to .xlsx file
      XSSFWorkbook workbook = new XSSFWorkbook(file);
      FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

      // Get first/desired sheet from the workbook
      XSSFSheet sheets = workbook.getSheetAt(0);

      // Iterate through each row one by one
      Iterator<Row> rowIterator = sheets.iterator();
      while (rowIterator.hasNext()) {
        Row row = rowIterator.next();
        // For each row, iterate through all the columns
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
          Cell cell = cellIterator.next();
          // Check the cell type after evaluating formula
          // IF it's formula cell, it will be evaluated, otherwise no change will happen
          switch (evaluator.evaluateInCell(cell).getCellType()) {
            case NUMERIC:
              System.out.println(cell.getNumericCellValue() + "NN");
              break;
            case STRING:
              System.out.println(cell.getStringCellValue() + "SS");
              break;
            case FORMULA:
              // Not again
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
