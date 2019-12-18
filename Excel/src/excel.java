import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excel {
	public static void main(String[] args) throws Exception {
		
		FileInputStream inputStream = new FileInputStream("D:\\LaedSuit.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		Sheet sheet = workbook.getSheet("TestSteps");
		Iterator iterator = sheet.iterator();
		
		while (iterator.hasNext()) {
			Row rowIterator = (Row) iterator.next();
			Iterator cellIterator = rowIterator.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cellData = (Cell) cellIterator.next();
				switch (cellData.getCellType()) {
				case STRING:
					System.out.println(cellData.getStringCellValue());
					break;
				case NUMERIC:
					System.out.println(cellData.getNumericCellValue());
					break;
				case BOOLEAN:
					System.out.println(cellData.getBooleanCellValue());
					break;

				}
			}
		}
	}

}
