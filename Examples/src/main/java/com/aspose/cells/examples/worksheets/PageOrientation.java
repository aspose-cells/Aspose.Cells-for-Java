package com.aspose.cells.examples.worksheets;

import com.aspose.cells.PageOrientationType;
import com.aspose.cells.PageSetup;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class PageOrientation {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(PageOrientation.class) + "worksheets/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Accessing the first worksheet in the Excel file
		WorksheetCollection worksheets = workbook.getWorksheets();
		int sheetIndex = worksheets.add();
		Worksheet sheet = worksheets.get(sheetIndex);

		// Setting the orientation to Portrait
		PageSetup pageSetup = sheet.getPageSetup();
		pageSetup.setOrientation(PageOrientationType.PORTRAIT);
		workbook.save(dataDir + "PageOrientation_out.xls");
	}
}
