package com.aspose.cells.examples.worksheets.management;

import com.aspose.cells.Cell;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

import java.io.FileInputStream;

public class AddWorksheetsToExistingExcelFile {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddWorksheetsToExistingExcelFile.class) + "worksheets/";
		String filePath = dataDir + "book1.xls";

		// Creating a file stream containing the Excel file to be opened
		FileInputStream fstream = new FileInputStream(filePath);

		// Instantiating a Workbook object with the stream
		Workbook workbook = new Workbook(fstream);

		// Adding a new worksheet to the Workbook object
		WorksheetCollection worksheets = workbook.getWorksheets();

		int sheetIndex = worksheets.add();
		Worksheet worksheet = worksheets.get(sheetIndex);

		// Setting the name of the newly added worksheet
		worksheet.setName("My Worksheet");

		// Saving the Excel file
		workbook.save(dataDir + "AWToExistingExcelFile-out.xls");

		// Print Message
		System.out.println("Sheet added successfully.");
		// ExEnd:1
	}
}
