package com.aspose.cells.examples.articles;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class VerifyPasswordtoProtectWorksheet {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(VerifyPasswordtoProtectWorksheet.class) + "articles/";
		// Create an instance of Workbook and load a spreadsheet
		Workbook book = new Workbook(dataDir + "book1.xlsx");

		// Access the protected Worksheet
		Worksheet sheet = book.getWorksheets().get(0);

		// Check if Worksheet is password protected
		if (sheet.getProtection().isProtectedWithPassword()) {
			// Verify the password used to protect the Worksheet
			if (sheet.getProtection().verifyPassword("password")) {
				System.out.println("Specified password has matched");
			} else {
				System.out.println("Specified password has not matched");
			}
		}

	}
}
