package org.test.excelIO;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadFromExcelFile {

  private static void readExcelFile(String filePath, String fileName, String sheetName) throws IOException {

    File file = new File(filePath + "\\" + fileName);

    FileInputStream inputStream = null;

    try {
      inputStream = new FileInputStream(file);

      Workbook workBook = null;

      String fileExtension = fileName.substring(fileName.indexOf("."));

      if (fileExtension.equals(".xls")) {
        workBook = new HSSFWorkbook(inputStream);

      } else if (fileExtension.equals(".xlsx")) {
        workBook = new XSSFWorkbook(inputStream);
      } else {
        System.out.println("Unsupported file to perform excel IO");
      }

      Sheet sheet = workBook.getSheet(sheetName);

      int rows = sheet.getLastRowNum() - sheet.getFirstRowNum();

      for (int i = 1; i <= rows; i++) {
        Row row = sheet.getRow(i);
        System.out.println();

        for (int j = 0; j < row.getLastCellNum(); j++) {
          System.out.print(row.getCell(j).getStringCellValue() + " ");
        }
      }
    } catch (Exception e) {
      e.printStackTrace();
    } finally {
      if (inputStream != null) {
        inputStream.close();
      }
    }
  }

  public static void main(String[] args) throws IOException {
    String filePath = System.getProperty("user.dir") + "\\src\\test\\resources";
    String sheetName = "Roommate Details";
    String fileName = "read.xlsx";

    readExcelFile(filePath, fileName, sheetName);
  }

}
