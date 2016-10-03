package com.aspose.cells.examples.worksheets;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class HideTabs {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(DisplayTab.class) + "worksheets/";

		// Instantiating a Workbook object by excel file path
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Hiding the tabs of the Excel file
		workbook.getSettings().setShowTabs(false);

		// Saving the modified Excel file in default (that is Excel 2003) format
		workbook.save(dataDir + "HideTabs_out.xls");

		// Print message
		System.out.println("Tabs are now hidden, please check the output file.");

	}
}
