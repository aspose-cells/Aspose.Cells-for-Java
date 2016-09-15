package com.aspose.cells.examples.files.utility;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class SetPDFCreationTime {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SetPDFCreationTime.class) + "files/";

		// Instantiate a Workbook object by excel file path
		Workbook workbook = new Workbook(dataDir + "Book1.xlsx");

		// Create an instance of PdfSaveOptions and pass SaveFormat to the
		// constructor
		PdfSaveOptions options = new PdfSaveOptions(FileFormatType.PDF);

		options.setCreatedTime(DateTime.getNow());
		// Save the file
		workbook.save(dataDir + "SPDFCTime-out.pdf", options);

		// Print message
		System.out.println("Set PDF Creation Time successfully.");

	}
}
