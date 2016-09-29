package com.aspose.cells.examples.articles;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SetExternalLinksInFormulas {
	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(SetExternalLinksInFormulas.class) + "articles/";
		// Instantiate a new Workbook.
		Workbook workbook = new Workbook();

		// Get first Worksheet
		Worksheet sheet = workbook.getWorksheets().get(0);

		// Get Cells collection
		Cells cells = sheet.getCells();

		// Set formula with external links
		cells.get("A1").setFormula("=SUM('[F:\\book1.xls]Sheet1'!A2, '[F:\\book1.xls]Sheet1'!A4)");

		// Set formula with external links
		cells.get("A2").setFormula("='[F:\\book1.xls]Sheet1'!A8");

		// Save the workbook
		workbook.save(dataDir + "SetExternalLinksInFormulas_out.xls");

	}
}
