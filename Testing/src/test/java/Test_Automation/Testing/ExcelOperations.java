package Test_Automation.Testing;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
//import org.apache.poi.xssf.model.InternalWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelOperations {
public static void main(String[] args) throws IOException {
		
		//Read Excel
		FileInputStream fileloc = new FileInputStream(".\\Test-data\\Automate.xlsx");
		XSSFWorkbook wbook = new XSSFWorkbook(fileloc);
		XSSFSheet wsheet = wbook.getSheetAt(0);
		//int lastrownumber = wsheet.getLastRowNum();
		
		XSSFRow wrow = wsheet.getRow(1);
		//int lastcellnumber = wrow.getLastCellNum();
		XSSFCell wcell = wrow.getCell(0);
		wcell.setCellType(CellType.STRING);
		String value = wcell.getStringCellValue();
		System.out.println(value);
		wbook.close();
		
		/*for(int i =1; i <=2; i++) {
			XSSFRow r1 = wsheet.getRow(i);
			for (int j = 0; j<=2; j++) {
				XSSFCell c1 = r1.getCell(j);
				c1.setCellType(CellType.STRING);
				String val = c1.getStringCellValue();
				{
			
			System.out.println(val);
			
			}
		wbook.close();
		}*/
			
		String[][] data = new String[lastrownumber][lastcellnumber];
			for(int i = 1; i <= lastrownumber; i++) {
				XSSFRow r1 = wsheet.getRow(i);
				for (int j = 0; j<=lastcellnumber; j++) {
					XSSFCell c1 = r1.getCell(j);
					c1.setCellType(CellType.STRING);
					String val = c1.getStringCellValue();
					{
				
				System.out.println(val);
				
				}
			wbook.close();
			}
			
		
		

	}

}
}


