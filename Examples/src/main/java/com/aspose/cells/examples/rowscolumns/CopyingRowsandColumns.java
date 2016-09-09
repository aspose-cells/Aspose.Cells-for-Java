package com.aspose.cells.examples.RowsColumns;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class CopyingRowsandColumns {

	public static void main(String[] args) throws Exception {
		// ExStart:CopyingRowsandColumns
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CopyingRowsandColumns.class) + "RowsColumns/";

		// Create a new Workbook.
		Workbook excelWorkbook = new Workbook(dataDir + "workbook.xls");

		// Get the first worksheet in the workbook.
		Worksheet wsTemplate = excelWorkbook.getWorksheets().get(0);

		// Copy the second row with data, formating, images and drawing objects
		// to the 12th row in the worksheet.
		wsTemplate.getCells().copyRow(wsTemplate.getCells(), 2, 10);

		// Copy the first column from the first worksheet of the first workbook
		// into
		// the first worksheet of the second workbook.
		wsTemplate.getCells().copyColumn(wsTemplate.getCells(), 1, 4);

		// Save the excel file.
		excelWorkbook.save(dataDir + "CopyingRowsandColumns-out.xls");

		// Print message
		System.out.println("Row and Column copied successfully.");
		// ExEnd:CopyingRowsandColumns
	}
}
