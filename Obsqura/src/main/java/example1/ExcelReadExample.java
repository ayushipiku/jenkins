package example1;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadExample {
	
	FileInputStream file;
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	Row row;
	Cell cell;
	String data;
	double num;
	boolean b;
	
	public String readData(int rowNo,int cellNo) {

		try {
		file = new FileInputStream("C:\\Users\\saiju\\git\\jenkins\\Obsqura\\src\\main\\resources\\Book1.xlsx");

		// Create Workbook instance holding reference to .xlsx file
		workbook = new XSSFWorkbook(file);
		} catch (Exception e) {
		e.printStackTrace();
		}
		// Get first/desired sheet from the workbook
		sheet = workbook.getSheet("Sheet1");

		row = sheet.getRow(rowNo);

		cell = row.getCell(cellNo);
		System.out.println(cell);
		CellType cType = cell.getCellType();
		//int c=cell.getCellType();
		
		
		switch(cType)
		{
			case STRING :
				data = cell.getStringCellValue();
				break;
			
			case NUMERIC : 
				num = cell.getNumericCellValue();
				data = String.valueOf(num);
				break;
			
			case BOOLEAN:
				b = cell.getBooleanCellValue();
				data = String.valueOf(b);
				break;
			default:
				data = null;
		}
		
		return data;
			
		}
	

	public static void main(String[] args) {

		 ExcelReadExample obj= new ExcelReadExample();	
		 
		 System.out.println(obj.readData(0, 0));
		 
		 System.out.println(obj.readData(1, 0));
	
		
			


			

	}

}
