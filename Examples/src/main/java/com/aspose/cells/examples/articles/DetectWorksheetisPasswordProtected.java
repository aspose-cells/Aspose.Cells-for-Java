package com.aspose.cells.examples.articles;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class DetectWorksheetisPasswordProtected {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(DetectWorksheetisPasswordProtected.class) + "articles/";
		// Create an instance of Workbook and load a spreadsheet
		Workbook book = new Workbook(dataDir + "sample.xlsx");

		// Access the protected Worksheet
		Worksheet sheet = book.getWorksheets().get(0);

		// Check if Worksheet is password protected
		if (sheet.getProtection().isProtectedWithPassword()) {
			System.out.println("Worksheet is password protected");
		}

	}
}
