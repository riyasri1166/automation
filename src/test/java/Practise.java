import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Practise {

	public static void main(String[] args) throws IOException {
		try {
			File file = new File("C:\\Users\\Dharshini\\OneDrive\\Desktop\\Priya.xlsx");
			FileInputStream set = new FileInputStream(file);
			Workbook book = new XSSFWorkbook(set);
			Sheet sheet = book.getSheet("Sheet1");
			// Row row = sheet.getRow(1);
			// Cell cell = row.getCell(1);
			// String stringCellValue = cell.getStringCellValue();
			// System.out.println(stringCellValue);

			for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
				Row row = sheet.getRow(i);
				for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
					Cell cell = row.getCell(j);
					// String stringCellValue = cell.getStringCellValue();
					// System.out.print(stringCellValue+"\t");

					CellType cellType = cell.getCellType();

					switch (cellType) {

					case STRING:
						String stringCellValue = cell.getStringCellValue();
						System.out.print(stringCellValue + "\t");

						break;

					default:

						if (DateUtil.isCellDateFormatted(cell)) {

							Date dateCellValue = cell.getDateCellValue();

							SimpleDateFormat s = new SimpleDateFormat("dd/MM/yyyy");

							String format = s.format(dateCellValue);

							System.out.println(format + "\t");

						} else {

							double numericCellValue = cell.getNumericCellValue();

							long l = (long) numericCellValue;

							System.out.print(l + "\t");
						}

						break;
					}
				}
				System.out.println(" ");

			}

			
		} catch (Exception e) {
			// TODO: handle exception
		}
		
	}
}
