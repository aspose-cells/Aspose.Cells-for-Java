package com.aspose.cells.examples.worksheets;

import com.aspose.cells.PageSetup;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class SetFirstPageNumber {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(SetFirstPageNumber.class) + "worksheets/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Accessing the first worksheet in the Excel file
		WorksheetCollection worksheets = workbook.getWorksheets();
		int sheetIndex = worksheets.add();
		Worksheet sheet = worksheets.get(sheetIndex);

		// Setting the first page number of the worksheet pages
		PageSetup pageSetup = sheet.getPageSetup();
		pageSetup.setFirstPageNumber(2);

		workbook.save(dataDir + "SetFirstPageNumber_out.xls");
	}
}
