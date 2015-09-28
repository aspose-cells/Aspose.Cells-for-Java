package featurescomparison.workingwithworkbook.createnewworkbook.java;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

/**
 * @author Shoaib Khan
 */
public class NewSpreadSheetAspose
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithworkbook/createnewworkbook/data/";
		
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Adding a new worksheet to the Workbook object
		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet worksheet = worksheets.add("My Worksheet");

		Cells cells = worksheet.getCells();

		// Adding some value to cell
		Cell cell = cells.get("A1");
		cell.setValue("This is Aspose test of fonts!");

		// Saving the Excel file
		workbook.save(dataPath + "newWorksheet_Aspose.xls");

		// Print Message
		System.out.println("Sheet added successfully.");
	}
}