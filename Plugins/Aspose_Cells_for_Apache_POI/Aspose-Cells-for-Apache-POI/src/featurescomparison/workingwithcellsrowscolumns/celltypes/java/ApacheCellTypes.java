package featurescomparison.workingwithcellsrowscolumns.celltypes.java;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ApacheCellTypes
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithcellsrowscolumns/celltypes/data/";
		
		Workbook wb = new HSSFWorkbook();
	    Sheet sheet = wb.createSheet("new sheet");
	    Row row = sheet.createRow((short)2);
	    row.createCell(0).setCellValue(1.1);
	    row.createCell(1).setCellValue(new Date());
	    row.createCell(2).setCellValue(Calendar.getInstance());
	    row.createCell(3).setCellValue("a string");
	    row.createCell(4).setCellValue(true);
	    row.createCell(5).setCellType(Cell.CELL_TYPE_ERROR);

	    // Write the output to a file
	    FileOutputStream fileOut = new FileOutputStream(dataPath + "ApacheCellTypes.xls");
	    wb.write(fileOut);
	    fileOut.close();
	    
	    System.out.println("Done.");
	}
}
