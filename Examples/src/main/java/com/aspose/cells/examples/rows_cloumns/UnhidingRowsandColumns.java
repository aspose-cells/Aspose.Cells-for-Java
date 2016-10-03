package com.aspose.cells.examples.rows_cloumns;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class UnhidingRowsandColumns {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(UnhidingRowsandColumns.class) + "rows_cloumns/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);
		Cells cells = worksheet.getCells();

		// Unhiding the 3rd row and setting its height to 13.5
		cells.unhideRow(2, 13.5);

		// Unhiding the 2nd column and setting its width to 8.5
		cells.unhideColumn(1, 8.5);

		// Saving the modified Excel file in default (that is Excel 2003) format
		workbook.save(dataDir + "UnhidingRowsandColumns_out.xls");

		// Print message
		System.out.println("Rows and Columns unhidden successfully.");

	}
}
