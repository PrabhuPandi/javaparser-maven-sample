package ExcelFunction;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelFunctionList {

	public void addCellBackgroundColor(Workbook workbook, Cell cell,String colour) {
		CellStyle style = workbook.createCellStyle();

		if (colour.equalsIgnoreCase("green")) {
			style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		}
		if (colour.equalsIgnoreCase("white")) {
			style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		}
		if (colour.equalsIgnoreCase("red")) {
			style.setFillForegroundColor(IndexedColors.RED.getIndex());
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		}
		cell.setCellStyle(style);		
	}

	public void updateExcelCell(Workbook workbook, Sheet sheet, int rowNumber,
			int columnNumber, String cellValue) {
		Cell createCell,cell; 
		 cell = sheet.getRow(rowNumber).getCell(columnNumber);

		if (cellValue.isEmpty() || cellValue == null) {
			System.out.println("inside empty");
			if (cell != null && cell.getCellTypeEnum() != CellType.BLANK) {
				cell.setCellType(CellType.BLANK);
				addCellBackgroundColor(workbook, cell, "white");

			} else {
				System.out.println("By Default CELL Value is blank");
				createCell =sheet.getRow(rowNumber).createCell(columnNumber);
				addCellBackgroundColor(workbook, createCell, "white");

			}
		} else {
			try {
				cell.setCellValue(cellValue);
			
			//System.out.println("Cell is update with " + cellValue + " " + cell);
			if (cellValue.equalsIgnoreCase("pass"))
				addCellBackgroundColor(workbook, cell, "green");
			else if (cellValue.equalsIgnoreCase("fail"))
				addCellBackgroundColor(workbook, cell, "red");
			} catch (java.lang.NullPointerException exception) {
				System.out.println("inside update Excel cell");
				System.out.println("rowNumber "+rowNumber);
				System.out.println("columnNumber "+columnNumber);
				createCell=sheet.getRow(rowNumber).createCell(columnNumber);
				createCell.setCellValue(cellValue);
				System.out.println("Cell is update with " + cellValue + " " + cell);
				updateExcelCell(workbook, sheet, rowNumber, columnNumber, cellValue);

			}
		}
	}

	public void PrintAllValues(int rowCounter, int getLastRowNum, Sheet sheet) {

		int getLastCellNum, columnCount;
		while (rowCounter <= getLastRowNum) {
			System.out.println("rowCounter " + rowCounter);
			Row row = sheet.getRow(rowCounter);
			getLastCellNum = row.getLastCellNum();
			// System.out.println("getLastCellNum " + getLastCellNum);
			// System.out.println();
			for (columnCount = 0; columnCount < getLastCellNum; columnCount++) {
				// System.out.println("rowCount "+rowCounter+" columnCount "+columnCount);

				Cell cell = row.getCell(columnCount);
				String value = cell.getStringCellValue();
				System.out.print(value + " | ");

				columnCount++;
			}
			System.out.println();
			rowCounter++;
		}
	}

	public int getCellPostion(int rowCounter, int getLastRowNum, Sheet sheet,
			String searchValue, boolean rowValue) {
		Row row = sheet.getRow(1);
		int getLastCellNum, columnCount, searchNumber = 0;
		boolean flag = false;

		getLastCellNum = row.getPhysicalNumberOfCells();

		while (rowCounter <= getLastRowNum) {
			// System.out.println("rowCounter " + rowCounter);
			row = sheet.getRow(rowCounter);
			// getLastCellNum = row.getLastCellNum();

			//System.out.println("getLastCellNum " + getLastCellNum);
			// System.out.println();
			for (columnCount = 0; columnCount < getLastCellNum; columnCount++) {
				// System.out.println("rowCount "+rowCounter+" columnCount "+columnCount);
				try {
					if (row.getCell(columnCount).getCellTypeEnum() == CellType.STRING) {

						if (row.getCell(columnCount).getStringCellValue()
								.equals(searchValue)) {
							if (rowValue) {
								searchNumber = rowCounter;
								flag = true;
							} else {
								searchNumber = columnCount;
								// System.out.println("columnCount " +
								// searchNumber);
								flag = true;
							}
							break;
						}
					}
				
				if (row.getCell(columnCount).getCellTypeEnum() == CellType.NUMERIC) {

					// if
					// (row.getCell(columnCount).getNumericCellValue()==searchIntValue)
					// {
					searchNumber = rowCounter;
					// System.out.println("searchNumber " + searchNumber);
					// flag = true;
					// break;
					// }

					// System.out.print(Math.round(row.getCell(columnCount)
					// .getNumericCellValue()) + "|| ");

					// System.out.print(row.getCell(columnCount)
					// .getNumericCellValue() + "|| ");

				}
				} catch (java.lang.NullPointerException exp) {
					System.out.println("null pointer exception");
				}
			}
			if (flag == true) {
				break;
			}
			rowCounter++;

		}
		if (flag == false) {
			System.out.println("Search Record "+searchValue+" is not found ");
			
		}
		return searchNumber;
	}

	public void getRowAndColumnPostion(String rowValueCheck,
			String columnValueCheck, Workbook workbook, Sheet sheet,
			int rowCounter, int getLastRowNum, String updateValue) {
		int rowNumber, columnNumber;
		rowNumber = getCellPostion(rowCounter, getLastRowNum, sheet,
				rowValueCheck, true);
		columnNumber = getCellPostion(rowCounter, getLastRowNum, sheet,
				columnValueCheck, false);
		//System.out.println(rowValueCheck + " " + rowNumber);
		//System.out.println(columnValueCheck + " " + columnNumber);
		updateExcelCell(workbook, sheet, rowNumber, columnNumber, updateValue);
	}

}
