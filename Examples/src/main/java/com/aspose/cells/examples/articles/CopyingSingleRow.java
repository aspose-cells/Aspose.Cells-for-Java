package com.aspose.cells.examples.articles;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class CopyingSingleRow {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CopyingSingleRow.class) + "articles/";
		// Create an instance of Workbook class by loading the existing spreadsheet
		Workbook workbook = new Workbook(dataDir + "aspose-sample.xlsx");

		// Get the cells collection of first worksheet
		Cells cells = workbook.getWorksheets().get(0).getCells();

		// Copy the first row to next 10 rows
		for (int i = 1; i <= 10; i++) {
			cells.copyRow(cells, 0, i);
		}
		// Save the result on disc
		workbook.save(dataDir + "CSingleRow_out.xlsx");

	}
}
