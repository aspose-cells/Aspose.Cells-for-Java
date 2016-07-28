package com.aspose.cells.examples.files.utility;

import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ConvertingToHTMLFiles {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(ConvertingToHTMLFiles.class);

		// Specify the file path
		String filePath = dataDir + "Book1.xlsx";

		// Specify the HTML saving options
		HtmlSaveOptions sv = new HtmlSaveOptions(SaveFormat.M_HTML);

		// Instantiate a workbook and open the template XLSX file
		Workbook wb = new Workbook(filePath);

		// Save the HTML file
		wb.save(dataDir + "output.html", sv);

		// Print message
		System.out.println("Excel to HTML conversion performed successfully.");
		// ExEnd:1
	}
}
