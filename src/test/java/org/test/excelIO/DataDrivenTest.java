package org.test.excelIO;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.Assert;

public class DataDrivenTest {

  public static void openUrl(String url, String title){
    WebDriver driver = new FirefoxDriver();
        
    driver.get(url);
    Assert.assertTrue(driver.getTitle().contains(title));
    driver.quit();
  }
  
  public static void main(String[] args) throws IOException {
    
    File file = new File(System.getProperty("user.dir") + "\\src\\test\\resources\\test_data.xlsx");
    
    FileInputStream fis = new FileInputStream(file);
    
    XSSFWorkbook workbook = new XSSFWorkbook(fis);
    XSSFSheet sheet = workbook.getSheet("Sheet1");
    
    for(int i=1; i <= sheet.getLastRowNum(); i++){
      XSSFRow row = sheet.getRow(i);
      openUrl(row.getCell(1).toString(), row.getCell(2).toString());
    }
    fis.close();
  }

}
