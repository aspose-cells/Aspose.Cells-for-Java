package com.aspose.cells.examples.articles;

import com.aspose.cells.PageOrientationType;
import com.aspose.cells.PageSetup;
import com.aspose.cells.PaperSizeType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SettingPageSetupOptions {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SettingPageSetupOptions.class) + "articles/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "CustomerReport.xls");

		// Accessing the first worksheet in the Excel file
		Worksheet sheet = workbook.getWorksheets().get(0);

		PageSetup pageSetup = sheet.getPageSetup();

		// Setting the orientation to Portrait
		pageSetup.setOrientation(PageOrientationType.PORTRAIT);

		// Setting the scaling factor to 100
		// pageSetup.setZoom(100);
		// OR Alternately you can use Fit to Page Options as under

		// Setting the number of pages to which the length of the worksheet will be spanned
		pageSetup.setFitToPagesTall(1);

		// Setting the number of pages to which the width of the worksheet will be spanned
		pageSetup.setFitToPagesWide(1);

		// Setting the paper size to A4
		pageSetup.setPaperSize(PaperSizeType.PAPER_A_4);

		// Setting the print quality of the worksheet to 1200 dpi
		pageSetup.setPrintQuality(1200);

		// Setting the first page number of the worksheet pages
		pageSetup.setFirstPageNumber(2);

		// Save the workbook
		workbook.save(dataDir + "SettingPageSetupOptions_out.xls");

	}
}
