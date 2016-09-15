package com.aspose.cells.examples.articles;

import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ConvertWorksheetToImageByPage {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ConvertWorksheetToImageByPage.class) + "articles/";
		// Create a new Workbook object
		// Open a template excel file
		Workbook book = new Workbook(dataDir + "bool1.xlsx");
		// Get the first worksheet
		Worksheet sheet = book.getWorksheets().get(0);
		// Define ImageOrPrintOptions
		ImageOrPrintOptions options = new ImageOrPrintOptions();
		// Set Resolution
		options.setHorizontalResolution(200);
		options.setVerticalResolution(200);
		options.setImageFormat(ImageFormat.getTiff());

		// Sheet2Image by page conversion
		SheetRender render = new SheetRender(sheet, options);
		for (int j = 0; j < render.getPageCount(); j++) {
			render.toImage(j, dataDir + sheet.getName() + " Page" + (j + 1) + ".tif");
		}

	}
}
