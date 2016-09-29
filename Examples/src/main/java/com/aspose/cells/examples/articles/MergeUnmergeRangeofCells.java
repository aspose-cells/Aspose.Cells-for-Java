package com.aspose.cells.examples.articles;

import com.aspose.cells.Range;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class MergeUnmergeRangeofCells {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(MergeUnmergeRangeofCells.class) + "articles/";
		// Create a workbook
		Workbook workbook = new Workbook();

		// Access the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Create a range
		Range range = worksheet.getCells().createRange("A1:D4");

		// Merge range into a single cell
		range.merge();

		// Save the workbook
		workbook.save(dataDir + "MURangeofCells_out.xlsx");

	}
}
