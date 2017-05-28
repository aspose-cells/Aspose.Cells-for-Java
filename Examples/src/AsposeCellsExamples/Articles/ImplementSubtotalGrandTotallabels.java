package com.aspose.cells.examples.articles;

import com.aspose.cells.CellArea;
import com.aspose.cells.ConsolidationFunction;
import com.aspose.cells.GlobalizationSettings;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ImplementSubtotalGrandTotallabels {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ImplementSubtotalGrandTotallabels.class) + "articles/";

		// Load your source workbook
		Workbook wb = new Workbook(dataDir + "sample.xlsx");

		// Set the glorbalization setting to change subtotal and grand total
		// names
		GlobalizationSettings gsi = new GlobalizationSettingsImp();
		wb.getSettings().setGlobalizationSettings(gsi);

		// Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		// Apply subtotal on A1:B10
		CellArea ca = CellArea.createCellArea("A1", "B10");
		ws.getCells().subtotal(ca, 0, ConsolidationFunction.SUM, new int[] { 2, 3, 4 });

		// Set the width of the first column
		ws.getCells().setColumnWidth(0, 40);

		// Save the output excel file
		wb.save(dataDir + "ImplementTotallabels_out.xlsx");
	}
}
