package com.aspose.cells.examples.files.utility;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class Excel2PDFConversion {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(Excel2PDFConversion.class) + "files/";

		Workbook workbook = new Workbook(dataDir + "Book1.xls");

		// Save the document in PDF format
		workbook.save(dataDir + "E2PDFC-out.pdf", SaveFormat.PDF);

		// Print message
		System.out.println("Excel to PDF conversion performed successfully.");
		// ExEnd:1
	}
}
