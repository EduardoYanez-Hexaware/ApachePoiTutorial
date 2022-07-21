import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class FormattingExcel {
  public static void main(String[] args) {
    // creating workbook
    XSSFWorkbook workbook = new XSSFWorkbook();

    // all calls here
    basedOnValue(workbook.createSheet("Value based formatting"));
    formatDuplicates(workbook.createSheet("Duplicates formatting"));
    shadeAlt(workbook.createSheet("Alternate rows"));
    expiryInNext30Days(workbook.createSheet("Soon expiring payments"));

    // output here
    try {
      FileOutputStream out = new FileOutputStream(new File("FormattingExcel.xlsx"));
      workbook.write(out);
      workbook.close();
      out.close();
      System.out.println("FormattingExcel.xlsx written successfully on disk.");
    } catch (IOException e) {
      e.printStackTrace();
    }
  }

  // Cell value is among a range
  static void basedOnValue(Sheet sheet) {
    // Creating some random values
    sheet.createRow(0).createCell(0).setCellValue(84);
    sheet.createRow(1).createCell(0).setCellValue(74);
    sheet.createRow(2).createCell(0).setCellValue(50);
    sheet.createRow(3).createCell(0).setCellValue(51);
    sheet.createRow(4).createCell(0).setCellValue(49);
    sheet.createRow(5).createCell(0).setCellValue(41);

    SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

    // Condition 1: Cell value is greater than 70 -> blue fill
    ConditionalFormattingRule ruleOne = sheetCF.createConditionalFormattingRule(ComparisonOperator.GT, "70");
    PatternFormatting fillOne = ruleOne.createPatternFormatting();
    fillOne.setFillBackgroundColor(IndexedColors.BLUE.index);
    fillOne.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

    // Condition 2: Cell value is less than 50 -> green fill
    ConditionalFormattingRule ruleTwo = sheetCF.createConditionalFormattingRule(ComparisonOperator.LT, "50");
    PatternFormatting fillTwo = ruleTwo.createPatternFormatting();
    fillTwo.setFillBackgroundColor(IndexedColors.GREEN.index);
    fillTwo.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

    CellRangeAddress[] regions = {
        CellRangeAddress.valueOf("A1:A6")
    };
    sheetCF.addConditionalFormatting(regions, ruleOne, ruleTwo);
  }

  // Use Excel conditional formatting to highlight duplicate entries in a column
  static void formatDuplicates(Sheet sheet) {
    sheet.createRow(0).createCell(0).setCellValue("Code");
    sheet.createRow(1).createCell(0).setCellValue(4);
    sheet.createRow(2).createCell(0).setCellValue(3);
    sheet.createRow(3).createCell(0).setCellValue(6);
    sheet.createRow(4).createCell(0).setCellValue(3);
    sheet.createRow(5).createCell(0).setCellValue(5);
    sheet.createRow(6).createCell(0).setCellValue(8);
    sheet.createRow(7).createCell(0).setCellValue(0);
    sheet.createRow(8).createCell(0).setCellValue(2);
    sheet.createRow(9).createCell(0).setCellValue(8);
    sheet.createRow(10).createCell(0).setCellValue(6);

    SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();
    // Condition 1: formula is = A2 = A1 -> white font
    ConditionalFormattingRule ruleOne = sheetCF.createConditionalFormattingRule("COUNTIF($A$2:$A$11,A2)>1");
    FontFormatting font = ruleOne.createFontFormatting();
    font.setFontStyle(false, true);
    font.setFontColorIndex(IndexedColors.BLUE.index);

    CellRangeAddress[] regions = {
      CellRangeAddress.valueOf("A2:A11")
    };
    sheetCF.addConditionalFormatting(regions, ruleOne);
  }

  // Use Excel conditional formatting to shade alternating rows on the worksheet
  static void shadeAlt(Sheet sheet) {
    SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

    // Condition 1: formula is = A2 = A1 -> white font
    ConditionalFormattingRule ruleOne = sheetCF.createConditionalFormattingRule("MOD(ROW(),2)");
    PatternFormatting fillOne = ruleOne.createPatternFormatting();
    fillOne.setFillBackgroundColor(IndexedColors.LIGHT_GREEN.index);
    fillOne.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

    CellRangeAddress[] regions = {
      CellRangeAddress.valueOf("A1:Z100")
    };
    sheetCF.addConditionalFormatting(regions, ruleOne);
  }

  // Code for financial projects which keeps track of deadlines
  static void expiryInNext30Days(Sheet sheet) {
    CellStyle style = sheet.getWorkbook().createCellStyle();
    style.setDataFormat((short) BuiltinFormats.getBuiltinFormat("d-mmm"));

    sheet.createRow(0).createCell(0).setCellValue("Date");
    sheet.createRow(1).createCell(0).setCellFormula("TODAY()+29");
    sheet.createRow(2).createCell(0).setCellFormula("A2+1");
    sheet.createRow(3).createCell(0).setCellFormula("A3+1");

    for (int rowNum = 1; rowNum <= 3; rowNum++) sheet.getRow(rowNum).getCell(0).setCellStyle(style);

    SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

    // Condition 1: formula is = A2 = A1 -> white font
    ConditionalFormattingRule ruleOne = sheetCF.createConditionalFormattingRule("AND(A2-TODAY()>=0,A2-TODAY()<=30)");
    FontFormatting font = ruleOne.createFontFormatting();
    font.setFontStyle(false, true);
    font.setFontColorIndex(IndexedColors.BLUE.index);

    CellRangeAddress[] regions = {
      CellRangeAddress.valueOf("A2:A4")
    };
    sheetCF.addConditionalFormatting(regions, ruleOne);
    sheet.getRow(0).createCell(1).setCellValue("Dates within the next 30 days are highlighted");
  }
}
