package com.aspose.cells.examples.data.addon.namedranges;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class CopyNamedRanges {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CopyNamedRanges.class) + "data/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		WorksheetCollection worksheets = workbook.getWorksheets();

		// Accessing the first worksheet in the Excel file
		Worksheet sheet = worksheets.get(0);
		Cells cells = sheet.getCells();

		// Creating a named range
		Range namedRange = cells.createRange("B4", "G14");
		namedRange.setName("TestRange");

		// Input some data with some formattings into
		// a few cells in the range.
		namedRange.get(0, 0).setValue("Test");
		namedRange.get(0, 4).setValue("123");

		// Creating a named range
		Range namedRange2 = cells.createRange("H4", "M14");
		namedRange2.setName("TestRange2");

		namedRange2.copy(namedRange);

		workbook.save(dataDir + "CNamedRanges-out.xls");

		// Print message
		System.out.println("Process completed successfully");

	}
}
