package com.aspose.cells.examples.worksheets.display;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class ControlTabBarWidth {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(ControlTabBarWidth.class);

		// Instantiating a Workbook object by excel file path
		Workbook workbook = new Workbook(dataDir + "book1.xlsx");

		workbook.getSettings().setShowTabs(true);
		workbook.getSettings().setSheetTabBarWidth(100);

		// Saving the modified Excel file in default (that is Excel 2000) format
		workbook.save(dataDir + "output.xls");

		// Print message
		System.out.println("Tab Bar width is updated, please check the output document.");
		// ExEnd:1
	}
}
