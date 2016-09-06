package com.aspose.cells.examples.worksheets.security.unprotecting;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.Protection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

/*import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;*/

public class UnprotectingPasswordProtectedWorksheet {

	public static void main(String[] args) throws Exception {
		// ExEnd:1
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(UnprotectingPasswordProtectedWorksheet.class);

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet worksheet = worksheets.get(0);

		Protection protection = worksheet.getProtection();

		// Unprotecting the worksheet with a password
		worksheet.unprotect("aspose");

		// Save the excel file.
		workbook.save(dataDir + "output.xls", FileFormatType.EXCEL_97_TO_2003);

		// Print Message
		System.out.println("Worksheet unprotected successfully.");
		// ExEnd:1
	}
}
