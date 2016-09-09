package com.aspose.cells.examples.RowsColumns;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class UngroupingRowsandColumns {

	public static void main(String[] args) throws Exception {
		// ExStart:UngroupingRowsandColumns
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(UngroupingRowsandColumns.class) + "RowsColumns/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "workbook.xls");

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);
		Cells cells = worksheet.getCells();

		// Grouping first six rows (from 0 to 5) and making them hidden by
		// passing true
		cells.ungroupRows(0, 10);

		// Grouping first three columns (from 0 to 2) and making them hidden by
		// passing true
		cells.ungroupColumns(0, 6);

		// Saving the modified Excel file in default (that is Excel 2003) format
		workbook.save(dataDir + "UngroupingRowsandColumns-out.xls");

		// Print message
		System.out.println("Rows and Columns ungrouped successfully.");
		// ExEnd:UngroupingRowsandColumns
	}
}
