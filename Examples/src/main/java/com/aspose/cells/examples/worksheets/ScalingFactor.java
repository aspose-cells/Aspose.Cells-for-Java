package com.aspose.cells.examples.worksheets;

import com.aspose.cells.PageSetup;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class ScalingFactor {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(ScalingFactor.class) + "worksheets/";
		// Instantiating a Excel object
		Workbook workbook = new Workbook();

		// Accessing the first worksheet in the Excel file
		WorksheetCollection worksheets = workbook.getWorksheets();
		int sheetIndex = worksheets.add();
		Worksheet sheet = worksheets.get(sheetIndex);

		// Setting the scaling factor to 100
		PageSetup pageSetup = sheet.getPageSetup();
		pageSetup.setZoom(100);
		workbook.save(dataDir + "ScalingFactor_out.xls");
	}
}
