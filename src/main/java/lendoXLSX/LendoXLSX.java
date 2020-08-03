package lendoXLSX;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LendoXLSX {

	public static void main(String[] args) throws IOException {

		String fileLocation = "D:\\projetos\\eclipse\\lendoXLSX\\Extra\\PlanilhaIn\\arquivo.xlsx";

		FileInputStream file = new FileInputStream(new File(fileLocation));
		Workbook workbook = new XSSFWorkbook(file);

		Sheet sheet = workbook.getSheetAt(0);

		// lendo as linhas
		Iterator<Row> rowIterator = sheet.iterator();

		while (rowIterator.hasNext()) {

			Row row = rowIterator.next();

			// lendo as celulas
			Iterator<Cell> cellIterator = row.iterator();

			while (cellIterator.hasNext()) {

				Cell cell = cellIterator.next();

				switch (cell.getCellTypeEnum()) {

				case STRING:
					System.out.print("TIPO STRING: " + cell.getStringCellValue());
					break;
				case NUMERIC:
					System.out.print("TIPO NUMERIC: " + cell.getNumericCellValue());
					break;
				case FORMULA:
					System.out.print("TIPO FORMULA: " + cell.getCellFormula());
				case BLANK:
					System.out.print("CELULA EM BRANCO");
				case _NONE:
					System.out.print("CELULA VAZIA");
				default:
					System.out.print("TIPO NÃO CONHECIDO");
				}
				
				System.out.println("\n");
			}
		}

	}

}
