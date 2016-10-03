package com.aspose.cells.examples.loading_saving;

import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ConvertingToHTMLFiles {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ConvertingToHTMLFiles.class) + "loading_saving/";

		// Specify the file path
		String filePath = dataDir + "Book1.xlsx";

		// Specify the HTML saving options
		HtmlSaveOptions sv = new HtmlSaveOptions(SaveFormat.M_HTML);

		// Instantiate a workbook and open the template XLSX file
		Workbook wb = new Workbook(filePath);

		// Save the HTML file
		wb.save(dataDir + "CToHTMLFiles_out.html", sv);

		// Print message
		System.out.println("Excel to HTML conversion performed successfully.");

	}
}
