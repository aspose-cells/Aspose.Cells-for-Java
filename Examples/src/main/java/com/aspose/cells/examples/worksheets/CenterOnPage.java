package com.aspose.cells.examples.worksheets;

import com.aspose.cells.PageSetup;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class CenterOnPage {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(CenterOnPage.class) + "worksheets/";
		// Create a workbook object
		Workbook workbook = new Workbook();

		// Get the worksheets in the workbook
		WorksheetCollection worksheets = workbook.getWorksheets();

		// Get the first (default) worksheet
		Worksheet worksheet = worksheets.get(0);

		// Get the pagesetup object
		PageSetup pageSetup = worksheet.getPageSetup();

		// Set bottom,left,right and top page margins
		pageSetup.setCenterHorizontally(true);
		pageSetup.setCenterVertically(true);

		workbook.save(dataDir + "CenterOnPage_out.xls");
	}
}
