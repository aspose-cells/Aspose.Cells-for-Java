package com.aspose.cells.examples.data;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class InputDataInCellsInRange {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(InputDataInCellsInRange.class) + "data/";

		// Instantiate a new Workbook.
		Workbook workbook = new Workbook();

		// Get the first worksheet in the workbook.
		Worksheet worksheet1 = workbook.getWorksheets().get(0);

		// Create a range of cells and specify its name based on H1:J4.
		Range range = worksheet1.getCells().createRange("H1:J4");
		range.setName("MyRange");

		// Input some data into cells in the range.
		range.get(0, 0).setValue("USA");
		range.get(0, 1).setValue("SA");
		range.get(0, 2).setValue("Israel");
		range.get(1, 0).setValue("UK");
		range.get(1, 1).setValue("AUS");
		range.get(1, 2).setValue("Canada");
		range.get(2, 0).setValue("France");
		range.get(2, 1).setValue("India");
		range.get(2, 2).setValue("Egypt");
		range.get(3, 0).setValue("China");
		range.get(3, 1).setValue("Philipine");
		range.get(3, 2).setValue("Brazil");

		// Save the excel file.
		workbook.save(dataDir + "IDICInRange_out.xls");

		// Print message
		System.out.println("Process completed successfully");

	}
}
