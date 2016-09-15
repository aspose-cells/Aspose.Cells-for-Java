package com.aspose.cells.examples.articles;

import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.WorkbookRender;
import com.aspose.cells.examples.Utils;

public class PrintingWholeWorkbook {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(PrintingWholeWorkbook.class) + "articles/";
		// Instantiate a new workbook
		Workbook book = new Workbook(dataDir + "Book1.xls");

		// Create an object for ImageOptions
		ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

		// WorkbookRender only support TIFF file format
		imgOptions.setImageFormat(ImageFormat.getTiff());

		// Create a WorkbookRender object with respect to your workbook
		WorkbookRender wr = new WorkbookRender(book, imgOptions);

		// Print the workbook
		wr.toPrinter(strPrinterName);

	}
}
