package com.aspose.cells.examples.data;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

import java.util.ArrayList;

public class ImportingFromArrayList {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ImportingFromArrayList.class) + "data/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Obtaining the reference of the worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Instantiating an ArrayList object
		ArrayList list = new ArrayList();

		// Add few names to the list as string values
		list.add("laurence chen");
		list.add("roman korchagin");
		list.add("kyle huang");
		list.add("tommy wang");

		// Importing the contents of ArrayList to 1st row and first column
		// vertically
		worksheet.getCells().importArrayList(list, 0, 0, true);

		// Saving the Excel file
		workbook.save(dataDir + "IFromArrayList_out.xls");

		// Printing the name of the cell found after searching worksheet
		System.out.println("Process completed successfully");

	}
}
