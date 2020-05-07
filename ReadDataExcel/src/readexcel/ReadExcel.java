package readexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public static void main(String[] args) throws Exception {
		File path = new File("E:\\Workspace\\CompanyDetail.xlsx");
		System.out.println("File path "+path);
       FileInputStream streem = new FileInputStream(path);
       System.out.println("file stream "+streem);
       XSSFWorkbook wb = new XSSFWorkbook(streem);
       System.out.println("Workbook excel "+wb);
       XSSFSheet sheet  = wb.getSheetAt(0);
       System.out.println("Sheet excel "+sheet);
       int lastrow = sheet.getLastRowNum();
       System.out.println("Last no of row "+lastrow);
       for(int i = 0;i<=lastrow;i++)
       {
    	  Row  ro = sheet.getRow(i);
//    	  System.out.println("Row interface "+ro);
    	  int lstecell = ro.getLastCellNum();
//    	  System.out.println("lastcell "+lstecell);
    	  for(int j = 0;j<lstecell;j++)
    	  {
    		  String data = ro.getCell(j).getStringCellValue();
    		  System.out.print(ro.getCell(j).getStringCellValue()+"||");
    	  }
    	  System.out.println("");
       }
	}

}
