package com.aspose.cells.examples.articles;

import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ExportWorksheettoImage {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ExportWorksheettoImage.class) + "articles/";
		// Create workbook object from source file
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		/*
		 * Set image or print options, We want one page per sheet, The image format is in png And desired dimensions are
		 * 400x400
		 */
		ImageOrPrintOptions opts = new ImageOrPrintOptions();
		opts.setOnePagePerSheet(true);
		opts.setImageFormat(ImageFormat.getPng());
		opts.setDesiredSize(400, 400);

		// Render sheet into image
		SheetRender sr = new SheetRender(worksheet, opts);
		sr.toImage(0, dataDir + "EWSheetToImage_out.png");


	}
}
