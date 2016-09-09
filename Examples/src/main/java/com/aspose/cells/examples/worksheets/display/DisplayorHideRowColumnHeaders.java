package com.aspose.cells.examples.worksheets.display;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class DisplayorHideRowColumnHeaders {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(DisplayorHideRowColumnHeaders.class) + "worksheets/";

		// Instantiating a Workbook object by excel file path
		Workbook workbook = new Workbook(dataDir + "book.xls");

		// Accessing the worksheets in the Excel file
		WorksheetCollection worksheets = workbook.getWorksheets();

		// Adding a worksheet in last place
		int sheetIndex = worksheets.add();
		Worksheet worksheet = worksheets.get(sheetIndex);

		// Hiding the headers of rows and columns
		worksheet.setRowColumnHeadersVisible(false);

		// Saving the modified Excel file in default (that is Excel 2000) format
		workbook.save(dataDir + "DisplayorHideRowColumnHeaders-out.xls");

		// Print Message
		System.out.println("Headers hidden successfully.");
		// ExEnd:1
	}
}
