package featurescomparison.workingwithworksheets.copysheetwithinworkbook.java;

import com.aspose.cells.Workbook;
import com.aspose.cells.WorksheetCollection;

public class AsposeCopySheetWithinWorkbook
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithworksheets/copysheetwithinworkbook/data/";
		
		//Create a new Workbook by excel file path
		Workbook wb = new Workbook();
		
		//Create a Worksheets object with reference to the sheets of the Workbook.
		WorksheetCollection sheets = wb.getWorksheets();
		
		//Copy data to a new sheet from an existing
		//sheet within the Workbook.
		sheets.addCopy("Sheet1");

		//Save the excel file.
		wb.save(dataPath + "AsposeCopyWorkbook_Out.xls");
		
		System.out.println("Sheet copied successfully."); // Print Message
	}
}
