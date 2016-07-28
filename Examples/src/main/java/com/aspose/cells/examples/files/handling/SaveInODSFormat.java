package com.aspose.cells.examples.files.handling;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class SaveInODSFormat {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(SaveInODSFormat.class);

		// Creating an Workbook object with an Excel file path
		Workbook workbook = new Workbook();

		// Save in ODS format
		workbook.save(dataDir + "output.ods", FileFormatType.ODS);

		// Print Message
		System.out.println("Worksheets are saved successfully.");
		// ExEnd:1
	}
}
