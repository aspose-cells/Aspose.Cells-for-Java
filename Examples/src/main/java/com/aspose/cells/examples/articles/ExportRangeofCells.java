package com.aspose.cells.examples.articles;

import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ExportRangeofCells {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ExportRangeofCells.class) + "articles/";
		// Create workbook from source file.
		Workbook workbook = new Workbook(dataDir + "aspose-sample.xlsx");

		// Access the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Set the print area with your desired range
		worksheet.getPageSetup().setPrintArea("E8:H15");

		// Set all margins as 0
		worksheet.getPageSetup().setLeftMargin(0);
		worksheet.getPageSetup().setRightMargin(0);
		worksheet.getPageSetup().setTopMargin(0);
		worksheet.getPageSetup().setBottomMargin(0);

		// Set OnePagePerSheet option as true
		ImageOrPrintOptions options = new ImageOrPrintOptions();
		options.setOnePagePerSheet(true);
		options.setImageFormat(ImageFormat.getJpeg());

		// Take the image of your worksheet
		SheetRender sr = new SheetRender(worksheet, options);
		sr.toImage(0, dataDir + "ERangeofCells_out.jpg");

	}
}
