package com.aspose.cells.examples.worksheets.management;

import java.io.FileInputStream;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class AddingWorksheetstoDesignerSpreadsheet {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(AddingWorksheetstoDesignerSpreadsheet.class);

		// Creating a file stream containing the Excel file to be opened
		FileInputStream fstream = new FileInputStream(dataDir + "book.xls");

		// Instantiating a Workbook object with the stream
		Workbook workbook = new Workbook(fstream);

		// Adding a new worksheet to the Workbook object
		WorksheetCollection worksheets = workbook.getWorksheets();
		int sheetIndex = worksheets.add();
		Worksheet worksheet = worksheets.get(sheetIndex);

		// Setting the name of the newly added worksheet
		worksheet.setName("My Worksheet");

		// Saving the Excel file
		workbook.save(dataDir + "output.xls");

		// Closing the file stream to free all resources
		fstream.close();

		// Print Message
		System.out.println("Sheet added successfully.");
		// ExEnd:1
	}
}
