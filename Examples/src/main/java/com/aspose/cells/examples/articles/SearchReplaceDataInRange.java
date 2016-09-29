package com.aspose.cells.examples.articles;

import com.aspose.cells.Cell;
import com.aspose.cells.CellArea;
import com.aspose.cells.FindOptions;
import com.aspose.cells.LookAtType;
import com.aspose.cells.LookInType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SearchReplaceDataInRange {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SearchReplaceDataInRange.class) + "articles/";

		Workbook workbook = new Workbook(dataDir + "input.xlsx");

		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Specify the range where you want to search
		// Here the range is E3:H6
		CellArea area = CellArea.createCellArea("E3", "H6");

		// Specify Find options
		FindOptions opts = new FindOptions();
		opts.setLookInType(LookInType.VALUES);
		opts.setLookAtType(LookAtType.ENTIRE_CONTENT);
		opts.setRange(area);

		Cell cell = null;

		do {
			// Search the cell with value search within range
			cell = worksheet.getCells().find("search", cell, opts);

			// If no such cell found, then break the loop
			if (cell == null)
				break;

			// Replace the cell with value replace
			cell.putValue("replace");

		} while (true);

		// Save the workbook
		workbook.save(dataDir + "SRDataInRange_out.xlsx");

	}
}
