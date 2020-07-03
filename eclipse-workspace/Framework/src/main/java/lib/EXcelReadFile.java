package lib;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EXcelReadFile {
	static Sheet sh;

	public static Object[][] excelToObjectArray(String fileName, String sheetName) throws IOException{
		Object [][] tabArray;
		FileInputStream fis = new FileInputStream(fileName);
		Workbook wb = new XSSFWorkbook(fis);
	    sh = wb.getSheet(sheetName);

	int totalRow = sh.getPhysicalNumberOfRows();    
	int totalCell = sh.getRow(0).getPhysicalNumberOfCells();
	    tabArray = new Object [totalRow - 1][totalCell];
	    
	for(int row=1; row<totalRow; row++) {
	for(int cell =0; cell<totalCell; cell++) {
		
		tabArray [row-1][cell] = cellData(row, cell);
			
	}
	}
		return tabArray;	
	}
		
	public static String cellData(int row, int cell) {
		
		Cell c= sh.getRow(row).getCell(cell);
		String data="";
		if (c.getCellType()==Cell.CELL_TYPE_STRING) {
			String textData = c.getStringCellValue();
		data= textData;
	}
		else if (c.getCellType()==Cell.CELL_TYPE_NUMERIC) {
			int number = (int)c.getNumericCellValue();	
			data=number+"";
		if (c.getNumericCellValue()%1==0) {
			data = ""+(int)c.getNumericCellValue();
	}
		else {
			data = ""+c.getNumericCellValue();
	}
	}
		
		return data;
		
	}	
	}

