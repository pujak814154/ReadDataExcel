package readexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ReadAllExcel {
	
	Workbook wb = null;
	public void readExcel(String filePath,String fileName,String sheetName) throws IOException
	{
		File file = new File(filePath + "\\" + fileName); 
		FileInputStream inputStream = new FileInputStream(file);
		
		String fileExtensionName = fileName.substring(fileName.indexOf("."));
		if(fileExtensionName.equals(".xlsx"))
		    {
			wb = new XSSFWorkbook(inputStream);
			}
	   else if(fileExtensionName.equals(".xls"))
	   {
				wb = new HSSFWorkbook(inputStream);
	   }
		Sheet mysheet = (Sheet) wb.getSheet(sheetName);
		
		int rowcount = mysheet.getLastRowNum();
		for (int i = 0; i <rowcount; i++) {
			Row row = mysheet.getRow(i);
			
			for (int j = 0; j < row.getLastCellNum(); j++) {
			
			System.out.print(row.getCell(j).getStringCellValue()+"|| ");
			}
			System.out.println();
			}
	}
	public static void main(String[] args) throws IOException {
		ReadAllExcel obj = new ReadAllExcel(); 
//		obj.readExcel("E:\\Workspace","details.xlsx","FirstSheet");
		obj.readExcel("E:\\Workspace","CompanyDetail.xlsx","Sheet2");

	
	}

}
