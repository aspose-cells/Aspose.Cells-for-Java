package com.aspose.cells.examples.worksheets;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class MoveWorksheet {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(MoveWorksheet.class) + "worksheets/";
		// Create a new Workbook.
		Workbook wb = new Workbook(dataDir + "BkFinance.xls");

		// Get the first worksheet in the book.
		Worksheet sheet = wb.getWorksheets().get(0);

		// Move the first sheet to the third position in the workbook.
		sheet.moveTo(2);

		// Save the Excel file.
		wb.save(dataDir + "MoveWorksheet_out.xls");
	}
}
