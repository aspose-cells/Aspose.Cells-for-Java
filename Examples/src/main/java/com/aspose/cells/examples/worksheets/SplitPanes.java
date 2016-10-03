package com.aspose.cells.examples.worksheets;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class SplitPanes {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SplitPanes.class) + "worksheets/";

		// Instantiate a new workbook
		// Open a template file
		Workbook book = new Workbook(dataDir + "book.xls");

		// Set the active cell
		book.getWorksheets().get(0).setActiveCell("A20");

		// Split the worksheet window
		book.getWorksheets().get(0).split();

		// Save the excel file
		book.save(dataDir + "SplitPanes_out.xls", SaveFormat.EXCEL_97_TO_2003);

		// Print Message
		System.out.println("Panes split successfully.");

	}
}
