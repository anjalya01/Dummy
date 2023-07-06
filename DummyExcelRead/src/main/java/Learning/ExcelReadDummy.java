package Learning;

import java.io.FileInputStream;

public class ExcelReadDummy {
	
	static FileInputStream f;
	static XSSFWorkbook w;
	static XSSFSheet sh;
	public static String getStringData(int a,int b)
	{
		f=new FileInputStream("C:\\Users\\Manu Sakthi\\OneDrive\\Desktop\\Anjaly\\Java\\ExcelRead.xlsx");
		w=new XSSFWorkbook(f);
		sh=w.getsheet("Sheet1");
		Row r=sh.getRow(a);
		Cell c=r.getCell(b);
		return c.getStringCellValue();
		
		
	}
public static String getIntegerData(int a,int b) {
	f=new FileInputStream("C:\\Users\\Manu Sakthi\\OneDrive\\Desktop\\Anjaly\\Java\\ExcelRead.xlsx");
	w=new XSSFWorkbook(f);
	sh=w.getsheet("Sheet1");
	Row r=sh.getRow(a);
	Cell c=r.getCell(b);
	int x=(int)c.getNumericCellValue();
	return String.valueOf(x);
	
	
}
	public static void main(String[] args) {
		// TODO Auto-generated method stub

String d=Excelcode.getStringData(0,1);
System.out.println(d);
String e=ExcelCode.getIntegerData(3,0);
System.out.println(e);
	}

}
