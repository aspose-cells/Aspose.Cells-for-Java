package com.aspose.cells.examples.articles;

import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class WorksheetToSeparateImage {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(WorksheetToSeparateImage.class) + "articles/";
		// Instantiate a new Workbook object
		// Open template
		Workbook book = new Workbook(dataDir + "source.xlsx");

		// Iterate over all worksheets in the workbook
		for (int i = 0; i < book.getWorksheets().getCount(); i++) {
			Worksheet sheet = book.getWorksheets().get(i);

			// Apply different Image and Print options
			ImageOrPrintOptions options = new ImageOrPrintOptions();

			// Set Horizontal Resolution
			options.setHorizontalResolution(300);

			// Set Vertical Resolution
			options.setVerticalResolution(300);

			// Set Image Format
			options.setImageFormat(ImageFormat.getJpeg());

			// If you want entire sheet as a single image
			options.setOnePagePerSheet(true);

			// Render to image
			SheetRender sr = new SheetRender(sheet, options);
			sr.toImage(0, dataDir + "WSheetToSImage_out-" + sheet.getName() + ".jpg");
		}

	}
}
