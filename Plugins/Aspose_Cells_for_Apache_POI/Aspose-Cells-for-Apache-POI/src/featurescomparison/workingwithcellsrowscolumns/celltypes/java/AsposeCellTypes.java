package featurescomparison.workingwithcellsrowscolumns.celltypes.java;

import java.util.Calendar;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class AsposeCellTypes
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithcellsrowscolumns/celltypes/data/";
		
		//Instantiating a Workbook object
		Workbook workbook = new Workbook();

		//Accessing the added worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);
		Cells cells = worksheet.getCells();

		//Adding a string value to the cell
		Cell cell = cells.get("A1");
		cell.setValue("Hello World");

		//Adding a double value to the cell
		cell = cells.get("A2");
		cell.setValue(20.5);

		//Adding an integer  value to the cell
		cell = cells.get("A3");
		cell.setValue(15);

		//Adding a boolean value to the cell
		cell = cells.get("A4");
		cell.setValue(true);

		//Adding a date/time value to the cell
		cell = cells.get("A5");
		cell.setValue(Calendar.getInstance());

		//Setting the display format of the date
		Style style = cell.getStyle();
		style.setNumber(15);
		cell.setStyle(style);

		//Saving the Excel file
		workbook.save(dataPath + "AsposeCellTypes.xls");
		
		System.out.println("Done.");
	}
}
