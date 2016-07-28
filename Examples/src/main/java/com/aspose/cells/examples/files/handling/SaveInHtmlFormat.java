package com.aspose.cells.examples.files.handling;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class SaveInHtmlFormat {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(SaveInHtmlFormat.class);

		// Creating an Workbook object with an Excel file path
		Workbook workbook = new Workbook();

		// Save in SpreadsheetML format
		workbook.save(dataDir + "output.html", FileFormatType.HTML);

		// Print Message
		System.out.println("Worksheets are saved successfully.");
		// ExEnd:1
	}
}
