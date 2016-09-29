package com.aspose.cells.examples.articles;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class CopyingMultipleColumns {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CopyingMultipleColumns.class) + "articles/";
		// Create an instance of Workbook class by loading the existing spreadsheet
		Workbook workbook = new Workbook(dataDir + "aspose-sample.xlsx");

		// Get the cells collection of worksheet by name Columns
		Cells cells = workbook.getWorksheets().get("Columns").getCells();

		// Copy the first 3 columns 7th column
		cells.copyColumns(cells, 0, 6, 3);

		// Save the result on disc
		workbook.save(dataDir + "CMultipleColumns_out.xlsx");

	}
}
