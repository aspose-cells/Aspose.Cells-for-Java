package com.aspose.cells.examples.articles;

import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class HtmlExportFrameScripts {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(HtmlExportFrameScripts.class) + "articles/";

		// Open the required workbook to convert
		Workbook w = new Workbook(dataDir + "Sample1.xlsx");

		// Disable exporting frame scripts and document properties
		HtmlSaveOptions options = new HtmlSaveOptions();
		options.setExportFrameScriptsAndProperties(false);

		// Save workbook as HTML
		w.save(dataDir + "HEFrameScripts_out.html", options);

		System.out.println("File saved");

	}
}
