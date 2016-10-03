package com.aspose.cells.examples.worksheets;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class RemovePanes {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(RemovePanes.class) + "worksheets/";

		// Instantiating a Excel object by excel file path
		Workbook workbook = new Workbook(dataDir + "Book1.xls");

		// Accessing the first worksheet in the Excel file
		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet worksheet = worksheets.get(0);

		worksheet.setActiveCell("A1");
		worksheet.removeSplit();

		// Saving the modified Excel file in default format
		workbook.save(dataDir + "RemovePanes_out.xls");

		// Print Message
		System.out.println("Panes Removed successfully.");

	}
}
