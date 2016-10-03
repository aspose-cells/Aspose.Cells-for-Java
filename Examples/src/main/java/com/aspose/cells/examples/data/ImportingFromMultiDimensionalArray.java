package com.aspose.cells.examples.data;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ImportingFromMultiDimensionalArray {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ImportingFromMultiDimensionalArray.class) + "data/";
		// Instantiate a new Workbook
		Workbook workbook = new Workbook();
		// Get the first worksheet (default sheet) in the Workbook
		Cells cells = workbook.getWorksheets().get("Sheet1").getCells();

		// Define a multi-dimensional array and store some data into it.
		String[][] strArray = { { "A", "1A", "2A" }, { "B", "2B", "3B" } };

		// Import the multi-dimensional array to the sheet
		cells.importArray(strArray, 0, 0);

		// Save the Excel file
		workbook.save(dataDir + "IFMDA_out.xlsx");
	}
}
