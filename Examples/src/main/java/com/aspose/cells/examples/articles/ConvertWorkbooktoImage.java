package com.aspose.cells.examples.articles;

import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.WorkbookRender;
import com.aspose.cells.examples.Utils;

public class ConvertWorkbooktoImage {
	public static void main(String[] args) throws Exception {
		// ExStart:ConvertWorkbooktoImage
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(ConvertWorkbooktoImage.class);
		// Instantiate a new Workbook object
		Workbook book = new Workbook(dataDir + "book1.xlsx");

		// Apply different Image and Print options
		ImageOrPrintOptions options = new ImageOrPrintOptions();

		// Set Image Format
		options.setImageFormat(ImageFormat.getTiff());

		// If you want entire sheet as a single image
		options.setOnePagePerSheet(true);

		// Render to image
		WorkbookRender render = new WorkbookRender(book, options);
		render.toImage(dataDir + "output.tiff");
		// ExEnd:ConvertWorkbooktoImage
	}
}
