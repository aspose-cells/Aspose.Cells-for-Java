package com.aspose.cells.examples.articles;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class DeletingBlankRows {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(DeletingBlankRows.class) + "articles/";
		// Create a new Workbook. Open an existing excel file.
		Workbook wb = new Workbook(dataDir + "Book1.xlsx");

		// Create a Worksheets object with reference to the sheets of the Workbook.
		WorksheetCollection sheets = wb.getWorksheets();

		// Get first Worksheet from WorksheetCollection
		Worksheet sheet = sheets.get(0);

		// Delete the Blank Rows from the worksheet
		sheet.getCells().deleteBlankRows();

		// Save the excel file.
		wb.save(dataDir + "DBlankRows_out.xlsx");


	}
}
