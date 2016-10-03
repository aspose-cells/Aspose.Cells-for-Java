package com.aspose.cells.examples.worksheets;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class HideUnhideWorksheet {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(HideUnhideWorksheet.class) + "worksheets/";

		// Instantiating a Workbook object by excel file path
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Accessing the first worksheet in the Excel file
		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet worksheet = worksheets.get(0);

		// Hiding the first worksheet of the Excel file
		worksheet.setVisible(false);

		// Saving the modified Excel file in default (that is Excel 2003) format
		workbook.save(dataDir + "HideUnhideWorksheet_out.xls");

		// Print message
		System.out.println("Worksheet 1 is now hidden, please check the output document.");

	}
}
