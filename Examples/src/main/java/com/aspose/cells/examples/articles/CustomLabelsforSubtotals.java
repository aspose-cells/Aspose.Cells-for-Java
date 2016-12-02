package com.aspose.cells.examples.articles;

import com.aspose.cells.CellArea;
import com.aspose.cells.ConsolidationFunction;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class CustomLabelsforSubtotals {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CustomLabelsforSubtotals.class) + "articles/";

		// Loads an existing spreadsheet containing some data
		Workbook book = new Workbook(dataDir + "sample.xlsx");

		// Assigns the GlobalizationSettings property of the WorkbookSettings
		// class
		// to the class created in first step
		book.getSettings().setGlobalizationSettings(new CustomSettings());

		// Accesses the 1st worksheet from the collection which contains data
		// Data resides in the cell range A2:B9
		Worksheet sheet = book.getWorksheets().get(0);

		// Adds SubTotal of type Average to the worksheet
		sheet.getCells().subtotal(CellArea.createCellArea("A2", "B9"), 0, ConsolidationFunction.AVERAGE, new int[] { 1 });

		// Calculates Formulas
		book.calculateFormula();

		// Auto fits all columns
		sheet.autoFitColumns();

		// Saves the workbook on disc
		book.save(dataDir + "CustomLabelsforSubtotals_out.xlsx");
	}
}
