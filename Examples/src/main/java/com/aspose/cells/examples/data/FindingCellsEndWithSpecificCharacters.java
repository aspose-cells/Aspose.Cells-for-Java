package com.aspose.cells.examples.data;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.FindOptions;
import com.aspose.cells.LookAtType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class FindingCellsEndWithSpecificCharacters {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(FindingCellsEndWithSpecificCharacters.class) + "data/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Finding the cell containing the specified formula
		Cells cells = worksheet.getCells();

		// Instantiate FindOptions
		FindOptions findOptions = new FindOptions();

		// Finding the cell containing a string value that ends with "es"
		findOptions.setLookAtType(LookAtType.END_WITH);
		Cell cell = cells.find("SH", null, findOptions);

		// Printing the name of the cell found after searching worksheet
		System.out.println("Name of the cell containing String: " + cell.getName());
	}
}
