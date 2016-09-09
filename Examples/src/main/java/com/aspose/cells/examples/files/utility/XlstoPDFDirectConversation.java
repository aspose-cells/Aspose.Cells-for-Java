package com.aspose.cells.examples.files.utility;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;
import com.aspose.cells.examples.files.handling.SavingFiletoSomeLocation;

public class XlstoPDFDirectConversation {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(XlstoPDFDirectConversation.class) + "files/";

		// Instantiate a Workbook object by excel file path
		Workbook workbook = new Workbook(dataDir + "Book1.xls");

		// Save the file
		workbook.save(dataDir + "XlsToPDFDC-out.pdf");

		// Print message
		System.out.println("Converted xls to Pdf successfully.");
		// ExEnd:1
	}
}
