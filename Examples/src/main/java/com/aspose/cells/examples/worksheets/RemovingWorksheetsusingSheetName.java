package com.aspose.cells.examples.worksheets;

import java.io.FileInputStream;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class RemovingWorksheetsusingSheetName {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(RemovingWorksheetsusingSheetName.class) + "worksheets/";

		// Creating a file stream containing the Excel file to be opened
		FileInputStream fstream = new FileInputStream(dataDir + "book.xls");

		// Instantiating a Workbook object with the stream
		Workbook workbook = new Workbook(fstream);

		// Removing a worksheet using its sheet name
		workbook.getWorksheets().removeAt("Sheet1");

		// Saving the Excel file
		workbook.save(dataDir + "RemovingWorksheetsusingSheetName_out.xls");

		// Closing the file stream to free all resources
		fstream.close();

		// Print Message
		System.out.println("Sheet removed successfully.");

	}
}
