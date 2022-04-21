package selenium_myproject1;

import java.io.File;
import java.io.FileInputStream;

import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;



import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;




public class test {

	public static void main(String[] args) throws IOException {
  File file = new File("C:\\Users\\user\\eclipse-workspace\\selenium_myproject1\\excel\\demo.xlsx");
  
 FileInputStream stream  = new FileInputStream(file);
 Workbook Workbook = new XSSFWorkbook(stream);
 Sheet sheet = Workbook.getSheet("Data");
 for (int i =0; i<sheet.getPhysicalNumberOfRows();i++) {
Row row = sheet.getRow(i);
 for (int j =0; j<row.getPhysicalNumberOfCells(); j++) {
	Cell Cell = row.getCell(j);
	CellType type = Cell.getCellType();
	switch (type) {
	case STRING:
    String text = Cell.getStringCellValue();
    System.out.println(text);
		
		break;
	case NUMERIC:
		double d = Cell.getNumericCellValue();
		BigDecimal b = BigDecimal.valueOf(d);
		String num = b.toString();
		System.out.println(num);

	default:
		break;
	}
	
}
}
	}	
}

