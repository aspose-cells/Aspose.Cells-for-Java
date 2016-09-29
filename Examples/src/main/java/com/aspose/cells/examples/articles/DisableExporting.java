package com.aspose.cells.examples.articles;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class DisableExporting {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(DisableExporting.class) + "articles/";
		// Open the required workbook to convert
		Workbook w = new Workbook(dataDir + "Sample1.xlsx");

		// Disable exporting frame scripts and document properties
		ImplementingIStreamProvider options = new ImplementingIStreamProvider();
		options.setExportFrameScriptsAndProperties(false);

		// Save workbook as HTML
		w.save(dataDir + "DisableExporting_out.html");

		System.out.println("File saved");

	}
}
