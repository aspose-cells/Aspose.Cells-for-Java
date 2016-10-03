package com.aspose.cells.examples.worksheets;

import com.aspose.cells.PageSetup;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class HeaderAndFooterMargins {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(HeaderAndFooterMargins.class) + "worksheets/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Accessing the first worksheet in the Excel file
		WorksheetCollection worksheets = workbook.getWorksheets();
		int sheetIndex = worksheets.add();
		Worksheet sheet = worksheets.get(sheetIndex);

		PageSetup pageSetup = sheet.getPageSetup();
		// Specify Header / Footer margins
		pageSetup.setHeaderMargin(2);
		pageSetup.setFooterMargin(2);

		workbook.save(dataDir + "HeaderAndFooterMargins_out.xls");
	}
}
