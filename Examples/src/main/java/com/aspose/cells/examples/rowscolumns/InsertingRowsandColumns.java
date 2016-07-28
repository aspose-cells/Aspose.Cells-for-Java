package com.aspose.cells.examples.RowsColumns;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class InsertingRowsandColumns {

	public static void main(String[] args) throws Exception {
		// ExStart:InsertingRowsandColumns
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(InsertingRowsandColumns.class);

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "workbook.xls");

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// ================== INSERTING ROWS ==================
		// Inserting a row into the worksheet at 6th position
		worksheet.getCells().insertRow(5);

		// Inserting 3 rows into the worksheet starting from 8th row
		worksheet.getCells().insertRows(7, 3);

		// ================== INSERTING COLUMNS ===============
		// Inserting a column into the worksheet at 4th position
		worksheet.getCells().insertColumn(3);

		// Inserting 3 columns into the worksheet at 6th position
		worksheet.getCells().insertColumns(5, 3);

		// Saving the modified Excel file in default (that is Excel 2000) format
		workbook.save(dataDir + "workbook.out.xls");

		// Print message
		System.out.println("Rows and Columns inserted successfully.");
		// ExEnd:InsertingRowsandColumns
	}
}
