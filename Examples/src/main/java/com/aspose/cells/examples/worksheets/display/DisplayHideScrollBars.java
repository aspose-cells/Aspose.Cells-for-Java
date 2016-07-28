package com.aspose.cells.examples.worksheets.display;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class DisplayHideScrollBars {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(DisplayHideScrollBars.class);

		// Instantiating a Excel object by excel file path
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Hiding the vertical scroll bar of the Excel file
		workbook.getSettings().setVScrollBarVisible(false);

		// Hiding the horizontal scroll bar of the Excel file
		workbook.getSettings().setHScrollBarVisible(false);

		// Saving the modified Excel file in default (that is Excel 2003) format
		workbook.save(dataDir + "output.xls");

		// Print message
		System.out.println("Scroll bars are now hidden, please check the output document.");
		// ExEnd:1
	}
}
