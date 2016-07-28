package com.aspose.cells.examples.worksheets.display;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class FreezePanes {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(FreezePanes.class);

		// Instantiating a Excel object by excel file path
		Workbook workbook = new Workbook(dataDir + "book.xls");

		// Accessing the first worksheet in the Excel file
		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet worksheet = worksheets.get(0);

		// Applying freeze panes settings
		worksheet.freezePanes(3, 2, 3, 2);

		// Saving the modified Excel file in default format
		workbook.save(dataDir + "book.out.xls");

		// Print Message
		System.out.println("Panes freeze successfull.");
		// ExEnd:1
	}
}
