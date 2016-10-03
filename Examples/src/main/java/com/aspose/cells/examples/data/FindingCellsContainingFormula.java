package com.aspose.cells.examples.data;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class FindingCellsContainingFormula {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(FindingCellsContainingFormula.class) + "data/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Finding the cell containing the specified formula
		Cells cells = worksheet.getCells();
		FindOptions findOptions = new FindOptions();
		findOptions.setLookInType(LookInType.FORMULAS);
		Cell cell = cells.find("=SUM(A5:A10)", null, findOptions);

		// Printing the name of the cell found after searching worksheet
		System.out.println("Name of the cell containing formula: " + cell.getName());

	}
}
