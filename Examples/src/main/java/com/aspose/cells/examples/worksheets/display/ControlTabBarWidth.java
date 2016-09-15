package com.aspose.cells.examples.worksheets.display;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;
import com.aspose.cells.examples.tables.ConvertTableToRange;

public class ControlTabBarWidth {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ControlTabBarWidth.class) + "worksheets/";

		// Instantiating a Workbook object by excel file path
		Workbook workbook = new Workbook(dataDir + "book1.xlsx");

		workbook.getSettings().setShowTabs(true);
		workbook.getSettings().setSheetTabBarWidth(100);

		// Saving the modified Excel file in default (that is Excel 2000) format
		workbook.save(dataDir + "ControlTabBarWidth-out.xls");

		// Print message
		System.out.println("Tab Bar width is updated, please check the output document.");

	}
}
