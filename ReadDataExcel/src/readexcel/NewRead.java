package readexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class NewRead {

	public static void main(String[] args) throws IOException {
		File path = new File("E:\\Workspace\\details.xlsx");
		FileInputStream readfile = new FileInputStream(path);
		Workbook wb = new XSSFWorkbook(readfile);
		XSSFSheet sheet1 = (XSSFSheet)wb.getSheetAt(0);
		XSSFSheet sheet12 = (XSSFSheet)wb.getSheet("FirstSheet");
        
		
		int rowcount = sheet1.getLastRowNum();
		System.out.println("total no of row " + rowcount);
		for ( int i=0;i<=rowcount;i++)
		{
		for ( int j=0;j<2;j++)
		{
		String data0 = sheet1.getRow(i).getCell(j).getStringCellValue();
		System.out.println(data0);
		//System.out.println("Data from row" + i + " is ==>> " +data0);
		}
		wb.close();
		}
		}
		
		
		
		
	}


