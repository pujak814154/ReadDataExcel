package readexcel;
import java.io.*;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelData {
   public static void main(String args[]) throws Exception
      {
	   
	   FileInputStream file = new FileInputStream(new File("E:\\Workspace\\details.xlsx"));
	   
	   XSSFWorkbook excelbook = new XSSFWorkbook(file);
	   XSSFSheet sheet = excelbook.getSheetAt(0);
	   
	   int rowcount = sheet.getLastRowNum();
	   
	   for(int i=0;i<=rowcount;i++)
	   {
	   String data = sheet.getRow(i).getCell(1).getStringCellValue();
	   
	   System.out.println(data);
	   }
      }
}
