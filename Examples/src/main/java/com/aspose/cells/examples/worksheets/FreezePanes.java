package com.aspose.cells.examples.worksheets;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class FreezePanes {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(FreezePanes.class) + "worksheets/";

		// Instantiating a Excel object by excel file path
		Workbook workbook = new Workbook(dataDir + "book.xls");

		// Accessing the first worksheet in the Excel file
		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet worksheet = worksheets.get(0);

		// Applying freeze panes settings
		worksheet.freezePanes(3, 2, 3, 2);

		// Saving the modified Excel file in default format
		workbook.save(dataDir + "FreezePanes_out.xls");

		// Print Message
		System.out.println("Panes freeze successfull.");

	}
}
