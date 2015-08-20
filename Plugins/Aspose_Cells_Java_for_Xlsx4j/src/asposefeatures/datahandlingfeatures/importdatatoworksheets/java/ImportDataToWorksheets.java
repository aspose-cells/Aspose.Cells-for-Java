package asposefeatures.datahandlingfeatures.importdatatoworksheets.java;

import java.util.ArrayList;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class ImportDataToWorksheets
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/asposefeatures/datahandlingfeatures/importdatatoworksheets/data/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Obtaining the reference of the newly added worksheet by passing its
		// sheet index
		int sheetIndex = workbook.getWorksheets().add();
		Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);

		// ==================================================
		// Creating an array containing names as string values
		String[] names = new String[] { "laurence chen", "roman korchagin",
				"kyle huang" };

		// Importing the array of names to 1st row and first column vertically
		Cells cells = worksheet.getCells();
		cells.importArray(names, 0, 0, false);

		// ==================================================
		ArrayList<String> list = new ArrayList<String>();

		// Add few names to the list as string values
		list.add("laurence chen");
		list.add("roman korchagin");
		list.add("kyle huang");

		// Importing the contents of ArrayList to 1st row and first column
		// vertically
		cells.importArrayList(list, 2, 0, true);
		// ==================================================

		// Saving the Excel file
		workbook.save(dataPath + "AsposeDataImport.xls");
		System.out.println("Done.");
	}
}
