package com.aspose.cells.examples.data;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class RemoveANamedRange {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(RemoveANamedRange.class) + "data/";

		// Instantiate a new Workbook.
		Workbook workbook = new Workbook();

		// Get all the worksheets in the book.
		WorksheetCollection worksheets = workbook.getWorksheets();

		// Get the first worksheet in the worksheets collection.
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Create a range of cells.
		Range range1 = worksheet.getCells().createRange("E12", "I12");

		// Name the range.
		range1.setName("MyRange");

		// Set the outline border to the range.
		range1.setOutlineBorder(BorderType.TOP_BORDER, CellBorderType.MEDIUM, Color.fromArgb(0, 0, 128));
		range1.setOutlineBorder(BorderType.BOTTOM_BORDER, CellBorderType.MEDIUM, Color.fromArgb(0, 0, 128));
		range1.setOutlineBorder(BorderType.LEFT_BORDER, CellBorderType.MEDIUM, Color.fromArgb(0, 0, 128));
		range1.setOutlineBorder(BorderType.RIGHT_BORDER, CellBorderType.MEDIUM, Color.fromArgb(0, 0, 128));

		// Input some data with some formattings into
		// a few cells in the range.
		range1.get(0, 0).setValue("Test");
		range1.get(0, 4).setValue("123");

		// Create another range of cells.
		Range range2 = worksheet.getCells().createRange("B3", "F3");

		// Name the range.
		range2.setName("testrange");

		// Copy the first range into second range.
		range2.copy(range1);

		// Remove the previous named range (range1) with its contents.
		worksheet.getCells().clearRange(11, 4, 11, 8);
		worksheets.getNames().removeAt(0);

		// Save the excel file.
		workbook.save(dataDir + "RANRange_out.xls");

		// Print message
		System.out.println("Process completed successfully");

	}
}
