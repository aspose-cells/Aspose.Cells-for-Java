package featurescomparison.workingwithdata.newlineincells.java;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ApacheNewLineinCell
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithdata/newlineincells/data/";
		
		Workbook wb = new XSSFWorkbook();   //or new HSSFWorkbook();
	    Sheet sheet = wb.createSheet();

	    Row row = sheet.createRow(2);
	    Cell cell = row.createCell(2);
	    cell.setCellValue("Use \n with word wrap on to create a new line");

	    //to enable newlines you need set a cell styles with wrap=true
	    CellStyle cs = wb.createCellStyle();
	    cs.setWrapText(true);
	    cell.setCellStyle(cs);

	    //increase row height to accommodate two lines of text
	    row.setHeightInPoints((2*sheet.getDefaultRowHeightInPoints()));

	    //adjust column width to fit the content
	    sheet.autoSizeColumn((short)2);

	    FileOutputStream fileOut = new FileOutputStream(dataPath + "Apache-Newlines.xlsx");
	    wb.write(fileOut);
	    fileOut.close();
	    
	    System.out.println("Done...");
	}
}