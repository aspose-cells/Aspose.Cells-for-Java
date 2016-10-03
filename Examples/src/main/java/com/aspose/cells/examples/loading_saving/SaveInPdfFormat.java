package com.aspose.cells.examples.loading_saving;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class SaveInPdfFormat {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SaveInPdfFormat.class) + "loading_saving/";

		// Creating an Workbook object with an Excel file path
		Workbook workbook = new Workbook();

		// Save in PDF format
		workbook.save(dataDir + "SIPdfFormat_out.pdf", FileFormatType.PDF);

		// Print Message
		System.out.println("Worksheets are saved successfully.");

	}
}
