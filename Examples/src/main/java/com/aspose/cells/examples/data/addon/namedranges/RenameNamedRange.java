package com.aspose.cells.examples.data.addon.namedranges;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class RenameNamedRange {

	public static void main(String[] args) throws Exception {
		// ExStart:RenameNamedRange
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(RenameNamedRange.class) + "data/";

		// Open an existing Excel file that has a (global) named range
		// "TestRange" in it
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Get the first worksheet
		Worksheet sheet = workbook.getWorksheets().get(0);

		// Get the Cells of the sheet
		Cells cells = sheet.getCells();

		// Get the named range "MyRange"
		Name name = workbook.getWorksheets().getNames().get("TestRange");

		// Rename it
		name.setText("NewRange");

		// Save the Excel file
		workbook.save(dataDir + "RNamedRange-out.xlsx");

		// Print message
		System.out.println("Process completed successfully");
		// ExEnd:RenameNamedRange
	}
}
