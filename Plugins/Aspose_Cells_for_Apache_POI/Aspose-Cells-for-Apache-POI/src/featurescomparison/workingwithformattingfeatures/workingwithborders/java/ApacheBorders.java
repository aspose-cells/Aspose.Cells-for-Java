package featurescomparison.workingwithformattingfeatures.workingwithborders.java;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ApacheBorders
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithformattingfeatures/workingwithborders/data/";
		
		Workbook wb = new HSSFWorkbook();
	    Sheet sheet = wb.createSheet("new sheet");

	    // Create a row and put some cells in it. Rows are 0 based.
	    Row row = sheet.createRow(1);

	    // Create a cell and put a value in it.
	    Cell cell = row.createCell(1);
	    cell.setCellValue(4);

	    // Style the cell with borders all around.
	    CellStyle style = wb.createCellStyle();
	    style.setBorderBottom(CellStyle.BORDER_THIN);
	    style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
	    
	    style.setBorderLeft(CellStyle.BORDER_THIN);
	    style.setLeftBorderColor(IndexedColors.GREEN.getIndex());
	    
	    style.setBorderRight(CellStyle.BORDER_THIN);
	    style.setRightBorderColor(IndexedColors.BLUE.getIndex());
	    
	    style.setBorderTop(CellStyle.BORDER_MEDIUM_DASHED);
	    style.setTopBorderColor(IndexedColors.BLACK.getIndex());
	    
	    cell.setCellStyle(style);

	    // Write the output to a file
	    FileOutputStream fileOut = new FileOutputStream(dataPath + "ApacheBorders_Out.xls");
	    wb.write(fileOut);
	    fileOut.close();
	    
	    System.out.println("Process Completed.");
	}
}
