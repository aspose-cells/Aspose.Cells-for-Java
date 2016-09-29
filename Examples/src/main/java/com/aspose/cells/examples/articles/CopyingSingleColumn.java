package com.aspose.cells.examples.articles;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class CopyingSingleColumn {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CopyingSingleColumn.class) + "articles/";
		// Create an instance of Workbook class by loading the existing spreadsheet
		Workbook workbook = new Workbook(dataDir + "aspose-sample.xlsx");

		// Get the cells collection of first workshet
		Cells cells = workbook.getWorksheets().get("Columns").getCells();

		// Copy the first column to next 10 columns
		for (int i = 1; i <= 10; i++) {
			cells.copyColumn(cells, 0, i);
		}
		// Save the result on disc
		workbook.save(dataDir + "CSingleColumn_out.xlsx");

	}
}
