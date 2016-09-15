package com.aspose.cells.examples.articles;

import com.aspose.cells.AutoFitterOptions;
import com.aspose.cells.Range;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AutoFitRowsforMergedCells {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AutoFitRowsforMergedCells.class) + "articles/";
		// Instantiate a new Workbook
		Workbook wb = new Workbook();

		// Get the first (default) worksheet
		Worksheet _worksheet = wb.getWorksheets().get(0);

		// Create a range A1:B1
		Range range = _worksheet.getCells().createRange(0, 0, 1, 2);

		// Merge the cells
		range.merge();

		// Insert value to the merged cell A1
		_worksheet.getCells().get(0, 0).setValue(
				"A quick brown fox jumps over the lazy dog. A quick brown fox jumps over the lazy dog....end");

		// Create a style object
		Style style = _worksheet.getCells().get(0, 0).getStyle();

		// Set wrapping text on
		style.setTextWrapped(true);

		// Apply the style to the cell
		_worksheet.getCells().get(0, 0).setStyle(style);

		// Create an object for AutoFitterOptions
		AutoFitterOptions options = new AutoFitterOptions();

		// Set auto-fit for merged cells
		options.setAutoFitMergedCells(true);

		// Autofit rows in the sheet(including the merged cells)
		_worksheet.autoFitRows(options);

		// Save the Excel file
		wb.save(dataDir + "AFRFMergedCells.xlsx");

	}
}
