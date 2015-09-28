package asposefeatures.datahandlingfeatures.findvalueincells.java;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.FindOptions;
import com.aspose.cells.LookAtType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class AsposeFindCellsWithString
{
    public static void main(String[] args) throws Exception
    {
	String dataPath = "src/asposefeatures/datahandlingfeatures/findvalueincells/data/";

	// Instantiating a Workbook object
	Workbook workbook = new Workbook(dataPath + "workbook.xls");

	// Accessing the first worksheet in the Excel file
	Worksheet worksheet = workbook.getWorksheets().get(0);

	// Finding the cell containing the specified formula
	Cells cells = worksheet.getCells();

	// Instantiate FindOptions
	FindOptions findOptions = new FindOptions();

	// Finding the cell containing a string value that starts with "Or"
	findOptions.setLookAtType(LookAtType.START_WITH);

	Cell cell = cells.find("SH", null, findOptions);

	// Printing the name of the cell found after searching worksheet
	System.out.println("Name of the cell containing String: " + cell.getName());
    }
}
