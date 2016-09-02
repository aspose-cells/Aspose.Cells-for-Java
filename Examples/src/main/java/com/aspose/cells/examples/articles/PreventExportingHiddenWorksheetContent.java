package com.aspose.cells.examples.articles;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class PreventExportingHiddenWorksheetContent {
	public static void main(String[] args) throws Exception {
		// ExStart:PreventExportingHiddenWorksheetContent
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(PreventExportingHiddenWorksheetContent.class);
		
		// Create workbook object
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		// Do not export hidden worksheet contents
		HtmlSaveOptions options = new HtmlSaveOptions();
		options.setExportHiddenWorksheet(false);

		// Save the workbook
		workbook.save(dataDir + ".out.html");
		// ExEnd:PreventExportingHiddenWorksheetContent
	}
}
