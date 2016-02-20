package featurescomparison.workingwithworkbook.adddataincells.java;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ApacheInsertCellsData
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithworkbook/adddataincells/data/";
		
		Workbook wb = new HSSFWorkbook();
		
		CreationHelper createHelper = wb.getCreationHelper();
		Sheet sheet = wb.createSheet("new sheet");
		
		// Create a row and put some cells in it. Rows are 0 based.
		Row row = sheet.createRow((short)0);
		// Create a cell and put a value in it.
		Cell cell = row.createCell(0);
		cell.setCellValue(1);
		
		// Or do it on one line.
		row.createCell(1).setCellValue(1.2);
		row.createCell(2).setCellValue(
		createHelper.createRichTextString("This is a string"));
		row.createCell(3).setCellValue(true);
		
		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream(dataPath + "DataInCells_Apache_Out.xls");
		wb.write(fileOut);
		fileOut.close();
		
		// Print message
		System.out.println("Data Added Successfully");
	}
}
