package com.aspose.cells.examples.worksheets.security.protecting;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class EditRangesWorksheet {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(EditRangesWorksheet.class) + "worksheets/";

		// Instantiating a Excel object by excel file path
		Workbook excel = new Workbook();

		// Accessing the first worksheet in the Excel file
		WorksheetCollection worksheets = excel.getWorksheets();
		Worksheet worksheet = worksheets.get(0);

		ProtectedRangeCollection allowranges = worksheet.getAllowEditRanges();
		ProtectedRange protected_range;

		int index = allowranges.add("r2", 1, 1, 3, 3);
		protected_range = allowranges.get(index);

		protected_range.setPassword("123");
		worksheet.protect(ProtectionType.ALL);

		// Saving the modified Excel file in default format
		excel.save(dataDir + "EditRangesWorksheet-out.xls");

		// Print Message
		System.out.println("you can Edit Range .");
		// ExEnd:1
	}
}
