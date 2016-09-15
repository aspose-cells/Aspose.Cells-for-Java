package com.aspose.cells.examples.articles;

import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class PrintingSelectedWorksheet {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(PrintingSelectedWorksheet.class) + "articles/";
		// Instantiate a new workbook
		Workbook book = new Workbook(dataDir + "Book1.xls");

		// Create an object for ImageOptions
		ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

		// Get the first worksheet
		Worksheet sheet = book.getWorksheets().get(0);
		// Create a SheetRender object with respect to your desired sheet
		SheetRender sr = new SheetRender(sheet, imgOptions);

		// Print the worksheet
		sr.toPrinter(strPrinterName);

	}
}
