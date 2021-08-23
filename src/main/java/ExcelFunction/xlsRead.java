package ExcelFunction;

import java.awt.Font;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class xlsRead extends ExcelFunctionList {
	public static void main(String[] args) throws IOException {
		int getPhysicalNumberOfRows, getFirstRowNum, getLastRowNum, getTopRow, rowCounter = 0, actualRowCount, columnCount = 0, getLastCellNum;
		String getSheetName, FILE_NAME;
		Footer getFooter;
		Header getHeader;
		FileInputStream file = null;
		int rowNumber, columnNumber;
		ExcelFunctionList excelFunctionList;
		FILE_NAME = System.getProperty("user.dir") + "\\bookNew.xlsx";
		System.out.println("Working Directory = "
				+ System.getProperty("user.dir"));

		try {
			excelFunctionList = new ExcelFunctionList();

			file = new FileInputStream(new File(FILE_NAME));
			Workbook workbook = new XSSFWorkbook(file);
			Sheet sheet = workbook.getSheet("TestData");

			getTopRow = sheet.getTopRow();
			getFooter = sheet.getFooter();
			getHeader = sheet.getHeader();

			getPhysicalNumberOfRows = sheet.getPhysicalNumberOfRows();
			getFirstRowNum = sheet.getFirstRowNum();
			getLastRowNum = sheet.getLastRowNum();
			getSheetName = sheet.getSheetName();

			actualRowCount = getLastRowNum - getFirstRowNum + 1;
			System.out.println("rowCount " + actualRowCount);
			rowCounter = getFirstRowNum;
			Row row = sheet.getRow(1);
			// getLastCellNum = row.getLastCellNum();

			System.out.println(" row.getLastCellNum() " + row.getLastCellNum());
			System.out.println(" row.getPhysicalNumberOfCells() "
					+ row.getPhysicalNumberOfCells());

			Row row2 = sheet.getRow(2);
			// getLastCellNum = row.getLastCellNum();

			System.out.println(" row.getLastCellNum() 2 "
					+ row2.getLastCellNum());
			System.out.println(" row.getPhysicalNumberOfCells() 2 "
					+ row2.getPhysicalNumberOfCells());
		
			
			excelFunctionList.getRowAndColumnPostion("Report_TC_01", "Status",
					workbook, sheet, rowCounter, getLastRowNum, "");
			excelFunctionList.getRowAndColumnPostion("Report_TC_02", "Status",
					workbook, sheet, rowCounter, getLastRowNum, "");
			excelFunctionList.getRowAndColumnPostion("Report_TC_03", "Status",
					workbook, sheet, rowCounter, getLastRowNum, "");
			excelFunctionList.getRowAndColumnPostion("Report_TC_04", "Status",
					workbook, sheet, rowCounter, getLastRowNum, "");
			excelFunctionList.getRowAndColumnPostion("Report_TC_05", "Status",
					workbook, sheet, rowCounter, getLastRowNum, "");
			excelFunctionList.getRowAndColumnPostion("Report_TC_06", "Status",
					workbook, sheet, rowCounter, getLastRowNum, "Pass");
			excelFunctionList.getRowAndColumnPostion("Report_TC_06", "Name",
					workbook, sheet, rowCounter, getLastRowNum, java.util.Calendar.getInstance().getTime().toString());
					
			file.close();

			FileOutputStream outFile = new FileOutputStream(new File(FILE_NAME));
			workbook.write(outFile);
			outFile.close();

		} catch (Exception e) {
			// System.out.println(e.getMessage());
			e.printStackTrace();
		}

	}

}
