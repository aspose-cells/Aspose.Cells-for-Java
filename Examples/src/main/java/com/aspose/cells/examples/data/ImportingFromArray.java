package com.aspose.cells.examples.data;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ImportingFromArray {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ImportingFromArray.class) + "data/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Obtaining the reference of the worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Creating an array containing names as string values
		String[] names = new String[] { "laurence chen", "roman korchagin", "kyle huang" };

		// Importing the array of names to 1st row and first column vertically
		Cells cells = worksheet.getCells();
		cells.importArray(names, 0, 0, false);

		// Saving the Excel file
		workbook.save(dataDir + "ImportingFromArray_out.xls");

		// Printing the name of the cell found after searching worksheet
		System.out.println("Process completed successfully");

	}
}
