package com.aspose.cells.examples.articles;

import com.aspose.cells.Cell;
import com.aspose.cells.FindOptions;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class FindCellsWithSpecificStyle {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(FindCellsWithSpecificStyle.class) + "articles/";

		Workbook workbook = new Workbook(dataDir + "TestBook.xlsx");

		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access the style of cell A1
		Style style = worksheet.getCells().get("A1").getStyle();

		// Specify the style for searching
		FindOptions options = new FindOptions();
		options.setStyle(style);

		Cell nextCell = null;

		do {
			// Find the cell that has a style of cell A1
			nextCell = worksheet.getCells().find(null, nextCell, options);

			if (nextCell == null)
				break;

			// Change the text of the cell
			nextCell.putValue("Found");

		} while (true);

		workbook.save(dataDir + "FCWithSpecificStyle_out.xlsx");

	}
}
