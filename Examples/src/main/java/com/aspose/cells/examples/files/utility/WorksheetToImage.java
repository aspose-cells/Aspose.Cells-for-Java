package com.aspose.cells.examples.files.utility;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class WorksheetToImage {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(WorksheetToImage.class) + "files/";

		// Instantiate a new workbook with path to an Excel file
		Workbook book = new Workbook(dataDir + "MyTestBook1.xls");

		// Create an object for ImageOptions
		ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

		// Set the image type
		imgOptions.setImageFormat(ImageFormat.getPng());

		// Get the first worksheet.
		Worksheet sheet = book.getWorksheets().get(0);

		// Create a SheetRender object for the target sheet
		SheetRender sr = new SheetRender(sheet, imgOptions);
		for (int j = 0; j < sr.getPageCount(); j++) {
			// Generate an image for the worksheet
			sr.toImage(j, dataDir + "WToImage-out" + j + ".png");
		}
		// Print message
		System.out.println("Images generated successfully.");
		// ExEnd:1
	}
}
