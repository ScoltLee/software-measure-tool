package tool;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
 
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class ReadExcel {
	
	public static final int COLUMN_INDEX_ELEMENT = 0;
	public static final int COLUMN_INDEX_L = 1;
	public static final int COLUMN_INDEX_A = 2;
	public static final int COLUMN_INDEX_H = 3;
	public static final int COLUMN_INDEX_S = 4;
 
    public void main(String[] args) throws IOException {
    	
        final String excelFilePath = "D:/Fp1.xlsx";
        
        final List<FunctionPoint> books = readExcel(excelFilePath);
      
        for (FunctionPoint book : books) {
            System.out.println(book);
        }
    }
 
    public static List<FunctionPoint> readExcel(String excelFilePath) throws IOException {
        List<FunctionPoint> listBooks = new ArrayList<>();
 
        // Get file
        InputStream inputStream = new FileInputStream(new File(excelFilePath));
 
        // Get workbook
        Workbook workbook = getWorkbook(inputStream, excelFilePath);
 
        // Get sheet
        Sheet sheet = workbook.getSheet("FunctionPoint");
        
        // Get all rows
        Iterator<Row> iterator = sheet.iterator();
        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            if (nextRow.getRowNum() == 0) {
                // Ignore header
                continue;
            }
         // Get all cells
            Iterator<Cell> cellIterator = nextRow.cellIterator();
 
            // Read cells and set value for book object
            FunctionPoint book = new FunctionPoint();
            while (cellIterator.hasNext()) {
                //Read cell
                Cell cell = cellIterator.next();
                Object cellValue = getCellValue(cell);
                if (cellValue == null || cellValue.toString().isEmpty()) {
                    continue;
                }
                // Set value for book object
                int columnIndex = cell.getColumnIndex();
                switch (columnIndex) {
                case COLUMN_INDEX_ELEMENT:
                    book.setEle((String) getCellValue(cell));
                    break;
                case COLUMN_INDEX_L:
                    book.setL((int)getCellValue(cell));
                    break;
                case COLUMN_INDEX_H:
                    book.setH((int)getCellValue(cell));
                    break;
                case COLUMN_INDEX_A:
                    book.setA((int)getCellValue(cell));
                    break;
                case COLUMN_INDEX_S:
                    book.setS((int)getCellValue(cell));
                    break;
                default:
                    break;
                }
 
            }
            listBooks.add(book);
        }
 
        workbook.close();
        inputStream.close();
 
        return listBooks;
    }
 
    // Get Workbook
    private static Workbook getWorkbook(InputStream inputStream, String excelFilePath) throws IOException {
        Workbook workbook = null;
        if (excelFilePath.endsWith("xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        } else if (excelFilePath.endsWith("xls")) {
            workbook = new HSSFWorkbook(inputStream);
        } else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        }
 
        return workbook;
    }
 // Get cell value
    private static Object getCellValue(Cell cell) {
        CellType cellType = cell.getCellType();
        Object cellValue = null;
        switch (cellType) {
        case BOOLEAN:
            cellValue = cell.getBooleanCellValue();
            break;
        case FORMULA:
            Workbook workbook = cell.getSheet().getWorkbook();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            cellValue = evaluator.evaluate(cell).getNumberValue();
            break;
        case NUMERIC:
            cellValue = cell.getNumericCellValue();
            break;
        case STRING:
            cellValue = cell.getStringCellValue();
            break;
        case _NONE:
        case BLANK:
        case ERROR:
            break;
        default:
            break;
        }
 
        return cellValue;
    }
}
