package com.aspose.cells.examples.articles;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class PreventExportingHiddenWorksheetContent {
	public static void main(String[] args) throws Exception {
		// ExStart:PreventExportingHiddenWorksheetContent
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(PreventExportingHiddenWorksheetContent.class) + "articles/";
		
		// Create workbook object
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		// Do not export hidden worksheet contents
		ImplementingIStreamProvider options = new ImplementingIStreamProvider();
		options.setExportHiddenWorksheet(false);

		// Save the workbook
		workbook.save(dataDir + "PEHWorksheetContent-out.html");
		// ExEnd:PreventExportingHiddenWorksheetContent
	}
}
