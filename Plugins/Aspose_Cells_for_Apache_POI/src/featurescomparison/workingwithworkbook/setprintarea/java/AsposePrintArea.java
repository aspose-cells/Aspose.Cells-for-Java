package featurescomparison.workingwithworkbook.setprintarea.java;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.PageSetup;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

public class AsposePrintArea
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithworkbook/setprintarea/data/";
		
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Accessing the first worksheet in the Workbook file
		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet sheet = worksheets.get(0);

		// Obtaining the reference of the PageSetup of the worksheet
		PageSetup pageSetup = sheet.getPageSetup();

		// Specifying the cells range (from A1 cell to F20 cell) of the print area
		pageSetup.setPrintArea("A1:F20");

		// Workbooks can be saved in many formats
		workbook.save(dataPath + "AsposePrintArea_Out.xlsx", FileFormatType.XLSX);

		System.out.println("Print Area Set successfully."); // Print Message
	}
}
