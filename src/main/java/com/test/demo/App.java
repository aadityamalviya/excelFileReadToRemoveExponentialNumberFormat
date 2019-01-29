package com.test.demo;

import java.io.File;
import java.io.FileInputStream;
import java.math.BigDecimal;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App {
	public static void main(String[] args) {
		try {
			FileInputStream file = new FileInputStream(new File("C:\\Users\\Aditya\\Desktop\\csv\\Settlement1.xlsx"));
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(0);
			Iterator rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				Row currentRow = (Row) rowIterator.next();
				if (currentRow.getRowNum() == 1) {
					Iterator cellIterator = currentRow.iterator();
					while (cellIterator.hasNext()) {
						Cell nextCell = (Cell) cellIterator.next();
						int columnIndex = nextCell.getColumnIndex();
						if (nextCell.getCellType() != CellType.STRING) {
							Double doubleValue = nextCell.getNumericCellValue();
							BigDecimal bd = new BigDecimal(doubleValue.toString());
							long lonVal = bd.longValue();
							String columnsValues = Long.toString(lonVal).trim();
							System.out.print("Numbers from columns " + columnsValues+"\n");
						}

					}
				}
				System.out.println();

			}
			file.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}
}
