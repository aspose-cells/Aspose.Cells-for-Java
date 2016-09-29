package com.aspose.cells.examples.articles;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ShowFormulas {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ShowFormulas.class) + "articles/";
		// Load the source workbook
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		// Access the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Show formulas of the worksheet
		worksheet.setShowFormulas(true);

		// Save the workbook
		workbook.save(dataDir + "ShowFormulas_out.xlsx");

	}
}
