package com.aspose.cells.examples.articles;

import com.aspose.cells.Cells;
import com.aspose.cells.Range;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class UsingDisplayRange {
	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(UsingDisplayRange.class) + "articles/";
		// Load a file in an instance of Workbook
		Workbook book = new Workbook(dataDir + "sample.xlsx");

		// Get Cells collection of first worksheet
		Cells cells = book.getWorksheets().get(0).getCells();

		// Get the MaxDisplayRange
		Range displayRange = cells.getMaxDisplayRange();

		// Loop over all cells in the MaxDisplayRange
		for (int row = displayRange.getFirstRow(); row < displayRange.getRowCount(); row++) {
			for (int col = displayRange.getFirstColumn(); col < displayRange.getColumnCount(); col++) {
				// Read the Cell value
				System.out.println(displayRange.get(row, col).getStringValue());
			}
		}

	}
}
