package asposefeatures.workingwithworksheets.settabcolor.java;

import com.aspose.cells.Color;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class SetWorksheetTabColor
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/asposefeatures/workingwithworksheets/settabcolor/data/";

		// Instantiate a new Workbook
		Workbook workbook = new Workbook(dataPath + "workbook.xls");

		// Get the first worksheet in the book
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Set the tab color
		worksheet.setTabColor(Color.getRed());

		// Save the Excel file
		workbook.save(dataPath + "AsposeColoredTab_Out.xls");

		System.out.println("Tab is now Colored.");
	}
}