package com.aspose.cells.examples.articles;

import com.aspose.cells.PivotTable;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class RefreshCalculatePivotTablehavingCalculatedItems {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(RefreshCalculatePivotTablehavingCalculatedItems.class) + "articles/";
		// Load source excel file containing a pivot table having calculated items
		Workbook wb = new Workbook(dataDir + "sample.xlsx");

		// Access first worksheet
		Worksheet sheet = wb.getWorksheets().get(0);

		// Change the value of cell D2
		sheet.getCells().get("D2").putValue(20);

		// Refresh and calculate all the pivot tables inside this sheet
		for (int i = 0; i < sheet.getPivotTables().getCount(); i++) {
			PivotTable pt = sheet.getPivotTables().get(i);
			pt.refreshData();
			pt.calculateData();
		}

		// Save the workbook in output pdf
		wb.save(dataDir + "RCPTHavingCItems_out.pdf", SaveFormat.PDF);

	}

}
