package com.aspose.cells.examples.worksheets.security.unprotecting;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class UnprotectSimpleSheet {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(UnprotectSimpleSheet.class) + "worksheets/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet worksheet = worksheets.get(0);

		// Unprotecting the worksheet
		worksheet.unprotect();

		// Save the excel file.
		workbook.save(dataDir + "USimpleSheet-out.xls", FileFormatType.EXCEL_97_TO_2003);

		// Print Message
		System.out.println("Worksheet unprotected successfully.");

	}
}
