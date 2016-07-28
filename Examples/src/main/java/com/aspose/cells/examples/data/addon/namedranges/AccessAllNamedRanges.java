package com.aspose.cells.examples.data.addon.namedranges;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class AccessAllNamedRanges {

	public static void main(String[] args) throws Exception {
		// ExStart:AccessAllNamedRanges
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(AccessAllNamedRanges.class);

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		WorksheetCollection worksheets = workbook.getWorksheets();

		// Accessing the first worksheet in the Excel file
		Worksheet sheet = worksheets.get(0);
		Cells cells = sheet.getCells();

		// Getting all named ranges
		Range[] namedRanges = worksheets.getNamedRanges();

		// Print message
		System.out.println("Number of Named Ranges : " + namedRanges.length);
		// ExEnd:AccessAllNamedRanges
	}
}
