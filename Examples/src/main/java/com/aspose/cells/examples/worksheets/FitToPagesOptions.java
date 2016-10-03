package com.aspose.cells.examples.worksheets;

import com.aspose.cells.PageSetup;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class FitToPagesOptions {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(FitToPagesOptions.class) + "worksheets/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Accessing the first worksheet in the Excel file
		WorksheetCollection worksheets = workbook.getWorksheets();
		int sheetIndex = worksheets.add();
		Worksheet sheet = worksheets.get(sheetIndex);

		PageSetup pageSetup = sheet.getPageSetup();

		// Setting the number of pages to which the length of the worksheet will be spanned
		pageSetup.setFitToPagesTall(1);

		// Setting the number of pages to which the width of the worksheet will be spanned
		pageSetup.setFitToPagesWide(1);
		workbook.save(dataDir + "FitToPagesOptions_out.xls");
	}
}
