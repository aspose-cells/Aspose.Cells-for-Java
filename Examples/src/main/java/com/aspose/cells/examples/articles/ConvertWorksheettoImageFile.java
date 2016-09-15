package com.aspose.cells.examples.articles;

import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ConvertWorksheettoImageFile {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ConvertWorksheettoImageFile.class) + "articles/";
		// Create a new Workbook object
		// Open a template excel file
		Workbook book = new Workbook(dataDir + "book.xlsx");
		// Get the first worksheet
		Worksheet sheet = book.getWorksheets().get(0);

		// Define ImageOrPrintOptions
		ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
		// Specify the image format
		imgOptions.setImageFormat(ImageFormat.getJpeg());

		// Render the sheet with respect to specified image/print options
		SheetRender render = new SheetRender(sheet, imgOptions);
		// Render the image for the sheet
		render.toImage(0, dataDir + "CWToImageFile.jpg");

	}
}
