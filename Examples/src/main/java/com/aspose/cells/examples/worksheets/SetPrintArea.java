package com.aspose.cells.examples.worksheets;

import com.aspose.cells.PageSetup;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class SetPrintArea {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(SetPrintArea.class) + "worksheets/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Accessing the first worksheet in the Workbook file
		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet sheet = worksheets.get(0);

		// Obtaining the reference of the PageSetup of the worksheet
		PageSetup pageSetup = sheet.getPageSetup();

		// Specifying the cells range (from A1 cell to T35 cell) of the print area
		pageSetup.setPrintArea("A1:T35");
		workbook.save(dataDir + "SetPrintArea_out.xls");
	}
}
