package com.aspose.cells.examples.worksheets;

import com.aspose.cells.PageSetup;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class SetPrintQuality {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(SetPrintQuality.class) + "worksheets/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Accessing the first worksheet in the Excel file
		WorksheetCollection worksheets = workbook.getWorksheets();
		int sheetIndex = worksheets.add();
		Worksheet sheet = worksheets.get(sheetIndex);

		// Setting the print quality of the worksheet to 180 dpi
		PageSetup pageSetup = sheet.getPageSetup();
		pageSetup.setPrintQuality(180);
		workbook.save(dataDir + "SetPrintQuality_out.xls");
	}
}
