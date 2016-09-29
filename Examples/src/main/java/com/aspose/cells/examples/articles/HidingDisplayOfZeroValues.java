package com.aspose.cells.examples.articles;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class HidingDisplayOfZeroValues {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(HidingDisplayOfZeroValues.class) + "articles/";

		// Create a new Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xlsx");

		// Get First worksheet of the workbook
		Worksheet sheet = workbook.getWorksheets().get(0);

		// Hide the zero values in the sheet
		sheet.setDisplayZeros(false);

		// Save the workbook
		workbook.save(dataDir + "HDOfZeroValues_out.xls");

	}
}
