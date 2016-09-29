package com.aspose.cells.examples.articles;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class CopyingMultipleRows {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CopyingMultipleRows.class) + "articles/";
		// Create an instance of Workbook class by loading the existing spreadsheet
		Workbook workbook = new Workbook(dataDir + "aspose-sample.xlsx");

		// Get the cells collection of worksheet by name Rows
		Cells cells = workbook.getWorksheets().get("Rows").getCells();

		// Copy the first 3 rows to 7th row
		cells.copyRows(cells, 0, 6, 3);

		// Save the result on disc
		workbook.save(dataDir + "CMultipleRows_out.xlsx");

	}
}
