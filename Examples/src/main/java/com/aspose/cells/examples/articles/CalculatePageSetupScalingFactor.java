package com.aspose.cells.examples.articles;

import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.PaperSizeType;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class CalculatePageSetupScalingFactor {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CalculatePageSetupScalingFactor.class) + "articles/";
		// Create workbook object
		Workbook workbook = new Workbook();

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Put some data in these cells
		worksheet.getCells().get("A4").putValue("Test");
		worksheet.getCells().get("S4").putValue("Test");

		// Set paper size
		worksheet.getPageSetup().setPaperSize(PaperSizeType.PAPER_A_4);

		// Set fit to pages wide as 1
		worksheet.getPageSetup().setFitToPagesWide(1);

		// Calculate page scale via sheet render
		SheetRender sr = new SheetRender(worksheet, new ImageOrPrintOptions());

		// Write the page scale value
		System.out.println(sr.getPageScale());

	}
}
