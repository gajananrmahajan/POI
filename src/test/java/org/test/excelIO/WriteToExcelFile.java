package org.test.excelIO;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteToExcelFile {

  private static void writeExcelFile(String filePath, String fileName, String sheetName, String[] toWrite) throws IOException {

    File file = new File(filePath + "\\" + fileName);

    FileInputStream inputStream = null;

    try {
      inputStream = new FileInputStream(file);

      Workbook workbook = null;

      String fileExtension = fileName.substring(fileName.indexOf("."));

      if (fileExtension.equals(".xls")) {
        workbook = new HSSFWorkbook(inputStream);
      } else if (fileExtension.equals(".xlsx")) {
        workbook = new XSSFWorkbook(inputStream);
      } else {
        System.out.println("Unsupported file to perform excel IO");
      }

      Sheet sheet = workbook.getSheet(sheetName);

      int rows = sheet.getLastRowNum() - sheet.getFirstRowNum();

      // Row row = sheet.getRow(rows);

      Row newRow = sheet.createRow(rows + 1);

      for (int i = 0; i < toWrite.length; i++) {
        Cell cell = newRow.createCell(i);
        cell.setCellValue(toWrite[i]);
      }

      FileOutputStream outputStream = new FileOutputStream(file);

      workbook.write(outputStream);

      outputStream.close();

    } catch (Exception e) {
      // TODO Auto-generated catch block
      e.printStackTrace();
    } finally {
      if (inputStream != null) {
        inputStream.close();
      }
    }

  }

  public static void main(String[] args) throws IOException {

    String filePath = System.getProperty("user.dir") + "\\src\\test\\resources";
    String fileName = "read.xlsx";
    String sheetName = "Roommate Details";
    String[] toWrite = { "Pravin Nawale", "Accountant" };

    writeExcelFile(filePath, fileName, sheetName, toWrite);
  }

}
