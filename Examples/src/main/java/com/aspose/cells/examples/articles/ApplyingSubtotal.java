package com.aspose.cells.examples.articles;

import com.aspose.cells.CellArea;
import com.aspose.cells.Cells;
import com.aspose.cells.ConsolidationFunction;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ApplyingSubtotal {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ApplyingSubtotal.class) + "articles/";
		// Create workbook from source Excel file
		Workbook workbook = new Workbook(dataDir + "Book1.xlsx");

		// Access the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Get the Cells collection in the first worksheet
		Cells cells = worksheet.getCells();

		// Create a cellarea i.e.., A2:B11
		CellArea ca = CellArea.createCellArea("A2", "B11");

		// Apply subtotal, the consolidation function is Sum and it will applied to
		// Second column (B) in the list
		cells.subtotal(ca, 0, ConsolidationFunction.SUM, new int[] { 1 }, true, false, true);

		// Set the direction of outline summary
		worksheet.getOutline().SummaryRowBelow = true;

		// Save the excel file
		workbook.save(dataDir + "ASubtotal_out.xlsx");

	}
}
