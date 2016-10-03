package com.aspose.cells.examples.worksheets;

import com.aspose.cells.PageSetup;
import com.aspose.cells.PrintOrderType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class SetPageOrder {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(SetPageOrder.class) + "worksheets/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Accessing the first worksheet in the Workbook file
		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet sheet = worksheets.get(0);

		// Obtaining the reference of the PageSetup of the worksheet
		PageSetup pageSetup = sheet.getPageSetup();

		// Setting the printing order of the pages to over then down
		pageSetup.setOrder(PrintOrderType.OVER_THEN_DOWN);

		workbook.save(dataDir + "SetPageOrder_out.xls");
	}
}
