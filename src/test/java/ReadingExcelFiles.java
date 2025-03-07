import java.io.FileInputStream;

import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcelFiles {

	public static void main(String[] args) throws IOException {

		// FileInputStream - reading
		// FileOutPutStream - writing

		// XSSFWorkbook - workbook
		// XSSFSheet - sheet
		// XSSFRow - row
		// XSSFCell - cell

		// -----------------------------------------------------------------------------------------------------------------------------

		// XSSFWorkbook accepts FileStream as an argument so must convert first
		// create object of FileInputStream and pass in file location
		// create object for XSSFWorkbook class
		// pass in the FileInputStream object in XSSFWorkbook as an argument

		FileInputStream file = new FileInputStream(System.getProperty("user.dir") + "\\testdata\\download_latest.xlsx");

		// Excel File > Workbook > Sheets > Rows > Cells

		XSSFWorkbook workbook = new XSSFWorkbook(file);

		XSSFSheet sheet = workbook.getSheet("Sheet1");

		int totalRows = sheet.getLastRowNum();

		int totalCells = sheet.getRow(1).getLastCellNum();

		System.out.println("Number of rows:" + " " + totalRows);
		System.out.println("Number of cells:" + " " + totalCells);

		for (int r = 0; r <= totalRows; r++) {

			XSSFRow currentRow = sheet.getRow(r);

			for (int c = 0; c < totalCells; c++) {

				XSSFCell currentCell = currentRow.getCell(c);
				String cell = currentCell.toString();
				System.out.print(cell + "\t");

			}
			System.out.println();
		}

		workbook.close();
		file.close();
		
		

	}

}
