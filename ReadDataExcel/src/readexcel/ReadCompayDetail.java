package readexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadCompayDetail {

	public static void main(String[] args) throws IOException {
		File path = new File("E:\\Workspace\\CompanyDetail.xlsx");
		FileInputStream file = new FileInputStream(path);
		XSSFWorkbook work = new XSSFWorkbook(file);
        XSSFSheet seet = work.getSheet("TestSheet1");
//        XSSFSheet seet1 = work.getSheetAt(0);
        int row = seet.getLastRowNum();
        System.out.println("Total no of row in sheet "+row);
        for(int i = 0;i<=row;i++)
        {
        	  
          Row ro = seet.getRow(i);
			
			for (int j = 0; j < ro.getLastCellNum(); j++) {
			
			System.out.print(ro.getCell(j).getStringCellValue()+"|| ");
			}
			System.out.println();
			}
        }
	}


