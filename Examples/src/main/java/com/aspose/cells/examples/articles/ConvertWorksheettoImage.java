package com.aspose.cells.examples.articles;

import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.PrintingPageType;
import com.aspose.cells.SheetRender;
import com.aspose.cells.TiffCompression;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ConvertWorksheettoImage {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ConvertWorksheettoImage.class) + "articles/";
		//Instantiate a new Workbook object
		//Open template
		Workbook book = new Workbook(dataDir + "book1.xlsx");

		//Get the first worksheet
		Worksheet sheet = book.getWorksheets().get(0);

		//Apply different Image and Print options
		ImageOrPrintOptions options = new ImageOrPrintOptions();

		//Set Horizontal Resolution
		options.setHorizontalResolution(300);

		//Set Vertical Resolution
		options.setVerticalResolution(300);

		//Set TiffCompression
		options.setTiffCompression(TiffCompression.COMPRESSION_LZW);

		//Set Image Format
		options.setImageFormat(ImageFormat.getTiff());

		//Set printing page type
		options.setPrintingPage(PrintingPageType.DEFAULT);

		//Render the sheet with respect to specified image/print options
		SheetRender sr = new SheetRender(sheet, options);

		//Render/save the image for the sheet
		sr.toImage(0, dataDir + "CWorksheettoImage_out.tiff");

	}
}
