package com.aspose.cells.examples.files.handling;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class OpeningHTMLFile {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(OpeningHTMLFile.class) + "files/";
		String filePath = dataDir + "Book1.html";

		// Opening html Files
		HTMLLoadOptions loadOptions = new HTMLLoadOptions(LoadFormat.HTML);
		// Create a Workbook object and opening the file from its path

		Workbook wb = new Workbook(filePath, loadOptions);
		// Print message
		System.out.println("Html format workbook has been opened successfully.");
		wb.save(dataDir + "OHTMLFile-out.xlsx", FileFormatType.XLSX);


	}
}
