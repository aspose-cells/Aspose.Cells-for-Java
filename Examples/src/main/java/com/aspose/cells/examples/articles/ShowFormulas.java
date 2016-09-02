package com.aspose.cells.examples.articles;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ShowFormulas {
	public static void main(String[] args) throws Exception {
		// ExStart:ShowFormulas
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(ShowFormulas.class);
		// Load the source workbook
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		// Access the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Show formulas of the worksheet
		worksheet.setShowFormulas(true);

		// Save the workbook
		workbook.save(dataDir + "out.xlsx");
		// ExEnd:ShowFormulas
	}
}
