package com.aspose.cells.examples.rows_cloumns;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SummaryRowRight {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SummaryRowRight.class) + "rows_cloumns/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "BookStyles.xls");

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);
		Cells cells = worksheet.getCells();

		// Grouping first six rows (from 0 to 5) and making them hidden by passing true
		cells.ungroupRows(0, 5);

		// Grouping first three columns (from 0 to 2) and making them hidden by passing true
		cells.ungroupColumns(0, 2);

		// Saving the modified Excel file in default (that is Excel 2003) format
		workbook.save(dataDir + "SummaryRowRight_out.xls");

		// Print message
		System.out.println("Rows and Columns ungrouped successfully.");
	}
}
