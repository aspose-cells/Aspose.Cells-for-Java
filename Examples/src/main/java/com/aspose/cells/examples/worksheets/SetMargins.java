package com.aspose.cells.examples.worksheets;

import com.aspose.cells.PageSetup;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class SetMargins {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(SetMargins.class) + "worksheets/";
		// Create a workbook object
		Workbook workbook = new Workbook();

		// Get the worksheets in the workbook
		WorksheetCollection worksheets = workbook.getWorksheets();

		// Get the first (default) worksheet
		Worksheet worksheet = worksheets.get(0);

		// Get the pagesetup object
		PageSetup pageSetup = worksheet.getPageSetup();

		// Set bottom,left,right and top page margins
		pageSetup.setBottomMargin(2);
		pageSetup.setLeftMargin(1);
		pageSetup.setRightMargin(1);
		pageSetup.setTopMargin(3);

		workbook.save(dataDir + "SetMargins_out.xls");
	}
}
