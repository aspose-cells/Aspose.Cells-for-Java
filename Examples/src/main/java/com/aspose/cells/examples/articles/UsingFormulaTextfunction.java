package com.aspose.cells.examples.articles;

import com.aspose.cells.Cell;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class UsingFormulaTextfunction {
	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(UsingFormulaTextfunction.class) + "articles/";
		// Create a workbook object
		Workbook workbook = new Workbook();

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Put some formula in cell A1
		Cell cellA1 = worksheet.getCells().get("A1");
		cellA1.setFormula("=Sum(B1:B10)");

		// Get the text of the formula in cell A2 using FORMULATEXT function
		Cell cellA2 = worksheet.getCells().get("A2");
		cellA2.setFormula("=FormulaText(A1)");

		// Calculate the workbook
		workbook.calculateFormula();

		// Print the results of A2. It will now print the text of the formula inside cell A1
		System.out.println(cellA2.getStringValue());

	}
}
