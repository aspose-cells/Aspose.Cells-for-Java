package AsposeCellsExamples.Data;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;
import java.io.*;

public class ExportingDataFromWorksheets {

	public static void main(String[] args) throws Exception {
		// ExStart: 1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ExportingDataFromWorksheets.class) + "Data/";

		// Creating a file stream containing the Excel file to be opened
		FileInputStream fstream = new FileInputStream(dataDir + "book1.xls");

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(fstream);

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Exporting the contents of 7 rows and 2 columns starting from 1st cell
		// to Array.
		Object dataTable[][] = worksheet.getCells().exportArray(0, 0, 7, 2);

		// Printing the number of rows exported
		System.out.println("No. Of Rows Exported: " + dataTable.length);

		// Closing the file stream to free all resources
		fstream.close();
		// ExEnd: 1

		System.out.println("ExportingDataFromWorksheets executed successfully.");
	}
}
