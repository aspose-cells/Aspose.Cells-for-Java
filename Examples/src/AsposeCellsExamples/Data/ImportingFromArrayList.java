package AsposeCellsExamples.Data;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

import java.util.ArrayList;

public class ImportingFromArrayList {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ImportingFromArrayList.class) + "Data/";
		//ExStart: 1

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Obtaining the reference of the worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Instantiating an ArrayList object
		ArrayList<String> list = new ArrayList<>();

		// Add few names to the list as string values
		list.add("laurence chen");
		list.add("roman korchagin");
		list.add("kyle huang");
		list.add("tommy wang");

		// Importing the contents of ArrayList to 1st row and first column
		// vertically
		worksheet.getCells().importArrayList(list, 0, 0, true);

		// Saving the Excel file
		workbook.save(dataDir + "IFromArrayList_out.xls");
		// ExEnd: 1

		// Printing the name of the cell found after searching worksheet
		System.out.println("Process completed successfully");
	}
}
