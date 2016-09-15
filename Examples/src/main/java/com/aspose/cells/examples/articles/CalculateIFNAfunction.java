package com.aspose.cells.examples.articles;

import com.aspose.cells.Cell;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class CalculateIFNAfunction {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CalculateIFNAfunction.class) + "articles/";
		// Create new workbook
		Workbook workbook = new Workbook();

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Add data for VLOOKUP
		worksheet.getCells().get("A1").putValue("Apple");
		worksheet.getCells().get("A2").putValue("Orange");
		worksheet.getCells().get("A3").putValue("Banana");

		// Access cell A5 and A6
		Cell cellA5 = worksheet.getCells().get("A5");
		Cell cellA6 = worksheet.getCells().get("A6");

		// Assign IFNA formula to A5 and A6
		cellA5.setFormula("=IFNA(VLOOKUP(\"Pear\",$A$1:$A$3,1,0),\"Not found\")");
		cellA6.setFormula("=IFNA(VLOOKUP(\"Orange\",$A$1:$A$3,1,0),\"Not found\")");

		// Caclulate the formula of workbook
		workbook.calculateFormula();

		// Print the values of A5 and A6
		System.out.println(cellA5.getStringValue());
		System.out.println(cellA6.getStringValue());

	}
}
