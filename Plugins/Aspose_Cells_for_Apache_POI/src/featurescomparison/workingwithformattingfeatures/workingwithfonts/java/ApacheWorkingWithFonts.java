package featurescomparison.workingwithformattingfeatures.workingwithfonts.java;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ApacheWorkingWithFonts
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithformattingfeatures/workingwithfonts/data/";
		
		Workbook wb = new HSSFWorkbook();
	    Sheet sheet = wb.createSheet("new sheet");

	    // Create a row and put some cells in it. Rows are 0 based.
	    Row row = sheet.createRow(1);

	    // Create a new font and alter it.
	    Font font = wb.createFont();
	    font.setFontHeightInPoints((short)24);
	    font.setFontName("Courier New");
	    font.setItalic(true);
	    font.setStrikeout(true);

	    // Fonts are set into a style so create a new one to use.
	    CellStyle style = wb.createCellStyle();
	    style.setFont(font);

	    // Create a cell and put a value in it.
	    Cell cell = row.createCell(1);
	    cell.setCellValue("This is a test of fonts");
	    cell.setCellStyle(style);

	    // Write the output to a file
	    FileOutputStream fileOut = new FileOutputStream(dataPath + "ApacheFonts.xls");
	    wb.write(fileOut);
	    fileOut.close();
	    
	    System.out.println("Apache Fonts Created.");
	}
}
