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
public class SecondSheet {
	 Workbook wb=null ;
	  
   public void exceldata(String path,String name,String sheetname) throws IOException
   {
	   File filepath = new File(path+ "\\" +name);
//	   try {
		   FileInputStream strem = new FileInputStream(filepath);
		   String extention = name.substring(name.indexOf("."));
		  
		   if(extention.equals(".xls"))
		   {
			  wb = new HSSFWorkbook(strem);
		   }
		   else if(extention.equals(".xlsx"))
		   {
			   wb= new XSSFWorkbook(strem);
		   }
		   Sheet mysheet = (Sheet) wb.getSheet(sheetname);
		   int  row = mysheet.getLastRowNum();
		   for(int i = 0;i<=row;i++)
		   {
			 Row ro = mysheet.getRow(i);
			 for(int j =0;j<ro.getLastCellNum();j++)
			 {
				 System.out.print(ro.getCell(j).getStringCellValue()+"|" );
			 }
			 System.out.println("");
		   }
		   
//	} catch (Exception e) {
//		System.out.println(e.toString());
//	}
   }
	
	public static void main(String[] args) throws Exception {
		SecondSheet obj = new SecondSheet();
//		File path = new File("E:\\Workspace\\");
//		File name = new File("CompanyDetail.xlsx");
		String sheetname = "Sheet2";
		obj.exceldata("E:\\Workspace","CompanyDetail.xlsx",sheetname);
	}

}
