package tool;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {
	public static final int COLUMN_INDEX_ELEMENT = 0;
	public static final int COLUMN_INDEX_L = 1;
	public static final int COLUMN_INDEX_A = 2;
	public static final int COLUMN_INDEX_H = 3;
	public static final int COLUMN_INDEX_S = 4;
	int number[][] = new int[6][4];
	private static CellStyle cellStyleFormatNumber = null;
	Workbook workbook;

	public void main(String[] args) throws IOException {
		final List<FunctionPoint> books = getFunctionPoint();
		final String excelFilePath = "D:/Fp1.xlsx";
		writeExcel(books, excelFilePath);
	}

	public void get(int[][] number, Workbook w) throws IOException {
		this.number = number;
		final List<FunctionPoint> books = getFunctionPoint();
		final String excelFilePath = "D:/Fp1.xlsx";
		writeExcel(books, excelFilePath);
	}

	public void renew() throws IOException {
		Workbook workbook = getWorkbook("D:/Fp1.xlsx");
		Sheet sheet = workbook.getSheet("FunctionPoint");
		Cell cell2Update = sheet.getRow(1).getCell(3);
		cell2Update.setCellValue(49);

	}

	public  void writeExcel(List<FunctionPoint> books, String excelFilePath) throws IOException {
		// Create Workbook
		workbook = getWorkbook(excelFilePath);
		// Create sheet
		Sheet sheet = workbook.createSheet("FunctionPoint"); // Create sheet with sheet name

		int rowIndex = 0;

		// Write header
		writeHeader(sheet, rowIndex);

		// Write data
		rowIndex++;
		for (FunctionPoint book : books) {
			// Create row
			Row row = sheet.createRow(rowIndex);
			// Write data on row
			writeBook(book, row);
			rowIndex++;
		}

		// Write footer
		writeFooter(sheet, rowIndex);

		// Auto resize column witdth
		int numberOfColumn = sheet.getRow(0).getPhysicalNumberOfCells();
		autosizeColumn(sheet, numberOfColumn);

		// Create file excel
		createOutputFile(workbook, excelFilePath);
		System.out.println("Done!!!");
	}

	// Create dummy data
	private List<FunctionPoint> getFunctionPoint() {
		List<FunctionPoint> listBook = new ArrayList<>();
		FunctionPoint book;
		for (int i = 1; i <= 5; i++) {
			if (i == 1) {
				book = new FunctionPoint("External Inputs (EI)", number[1][1], number[1][2], number[1][3]);
				listBook.add(book);
			}
			if (i == 2) {
				book = new FunctionPoint("External Outputs (EO)", number[2][1], number[2][2], number[2][3]);
				listBook.add(book);
			}
			if (i == 3) {
				book = new FunctionPoint("External Inquiries (EQ)", number[3][1], number[3][2], number[3][3]);
				listBook.add(book);
			}
			if (i == 4) {
				book = new FunctionPoint("External Interface Files (EIF)", number[4][1], number[4][2], number[4][3]);
				listBook.add(book);
			}
			if (i == 5) {
				book = new FunctionPoint("Internal Logical Files (ILF)", number[5][1], number[5][2], number[5][3]);
				listBook.add(book);
			}
		}
		return listBook;
	}

	// Create workbook
	private static Workbook getWorkbook(String excelFilePath) throws IOException {
		Workbook workbook = null;

		if (excelFilePath.endsWith("xlsx")) {
			workbook = new XSSFWorkbook();
		} else if (excelFilePath.endsWith("xls")) {
			workbook = new HSSFWorkbook();
		} else {
			throw new IllegalArgumentException("The specified file is not Excel file");
		}

		return workbook;
	}

	// Write header with format
	private void writeHeader(Sheet sheet, int rowIndex) {
		// create CellStyle
		CellStyle cellStyle = createStyleForHeader(sheet);

		// Create row
		Row row = sheet.createRow(rowIndex);

		// Create cells
		Cell cell = row.createCell(COLUMN_INDEX_ELEMENT);
		cell.setCellStyle(cellStyle);
		cell.setCellValue("Elements");

		cell = row.createCell(COLUMN_INDEX_L);
		cell.setCellStyle(cellStyle);
		cell.setCellValue("L");

		cell = row.createCell(COLUMN_INDEX_A);
		cell.setCellStyle(cellStyle);
		cell.setCellValue("A");

		cell = row.createCell(COLUMN_INDEX_H);
		cell.setCellStyle(cellStyle);
		cell.setCellValue("H");

		cell = row.createCell(COLUMN_INDEX_S);
		cell.setCellStyle(cellStyle);
		cell.setCellValue("Sum");
	}

	// Write data
	private static void writeBook(FunctionPoint book, Row row) {
		if (cellStyleFormatNumber == null) {
			// Format number
			short format = (short) BuiltinFormats.getBuiltinFormat("#,##0");
			// DataFormat df = workbook.createDataFormat();
			// short format = df.getFormat("#,##0");

			// Create CellStyle
			Workbook workbook = row.getSheet().getWorkbook();
			cellStyleFormatNumber = workbook.createCellStyle();
			cellStyleFormatNumber.setDataFormat(format);
		}

		Cell cell = row.createCell(COLUMN_INDEX_ELEMENT);
		cell.setCellValue(book.getEle());

		cell = row.createCell(COLUMN_INDEX_L);
		cell.setCellValue(book.getL());

		cell = row.createCell(COLUMN_INDEX_A);
		cell.setCellValue(book.getA());
		cell.setCellStyle(cellStyleFormatNumber);

		cell = row.createCell(COLUMN_INDEX_H);
		cell.setCellValue(book.getH());

		// Create cell formula
		// totalMoney = price * quantity
		cell = row.createCell(COLUMN_INDEX_S, CellType.FORMULA);
		cell.setCellStyle(cellStyleFormatNumber);
		int currentRow = row.getRowNum() + 1;
		String columnL = CellReference.convertNumToColString(COLUMN_INDEX_L);
		String columnA = CellReference.convertNumToColString(COLUMN_INDEX_A);
		String columnH = CellReference.convertNumToColString(COLUMN_INDEX_H);
		cell.setCellFormula(columnA + currentRow + "+" + columnL + currentRow + "+" + columnH + currentRow);
	}

	// Create CellStyle for header
	public CellStyle createStyleForHeader(Sheet sheet) {
		// Create font
		Font font = sheet.getWorkbook().createFont();
		font.setFontName("Times New Roman");
		font.setBold(true);
		font.setFontHeightInPoints((short) 14); // font size
		font.setColor(IndexedColors.WHITE.getIndex()); // text color

		// Create CellStyle
		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setFont(font);
		cellStyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellStyle.setBorderBottom(BorderStyle.THIN);
		return cellStyle;
	}

	// Write footer
	private static void writeFooter(Sheet sheet, int rowIndex) {
		// Create row
		Row row = sheet.createRow(rowIndex);
		Cell cell = row.createCell(COLUMN_INDEX_S, CellType.FORMULA);
		cell.setCellFormula("SUM(E2:E6)");
	}

	// Auto resize column width
	private static void autosizeColumn(Sheet sheet, int lastColumn) {
		for (int columnIndex = 0; columnIndex < lastColumn; columnIndex++) {
			sheet.autoSizeColumn(columnIndex);
		}
	}

	// Create output file
	private static void createOutputFile(Workbook workbook, String excelFilePath) throws IOException {
		try (OutputStream os = new FileOutputStream(excelFilePath)) {
			workbook.write(os);
		}
	}

}
