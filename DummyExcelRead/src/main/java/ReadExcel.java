
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
//"C:\Users\Manu Sakthi\OneDrive\Desktop\Anjaly\Java\ExcelRead.xlsx"
	public static void main(String[] args) throws IOException {
		
		FileInputStream fis=new FileInputStream(new File("C:\\Users\\Manu Sakthi\\OneDrive\\Desktop\\Anjaly\\Java\\ExcelRead.xlsx"));  
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sheet=wb.getSheetAt(0);  
		for(Row row: sheet)     //iteration over row using for each loop  
		{  
		for(Cell cell: row)    //iteration over cell using for each loop  
		{  
		System.out.println(cell);
	}
		}
	}
}

	


