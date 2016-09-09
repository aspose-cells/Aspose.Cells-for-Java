package com.aspose.cells.examples.formulas;

import com.aspose.cells.Cell;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;
import com.aspose.cells.examples.formatting.TextAlignmentVertical;

public class DirectCalculationFormula {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(DirectCalculationFormula.class) + "formulas/";
		// Create a workbook
		Workbook workbook = new Workbook();

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Put 20 in cell A1
		Cell cellA1 = worksheet.getCells().get("A1");
		cellA1.putValue(20);

		// Put 30 in cell A2
		Cell cellA2 = worksheet.getCells().get("A2");
		cellA2.putValue(30);

		// Calculate the Sum of A1 and A2
		Object results = worksheet.calculateFormula("=Sum(A1:A2)");

		// Print the output
		System.out.println("Value of A1: " + cellA1.getStringValue());
		System.out.println("Value of A2: " + cellA2.getStringValue());
		System.out.println("Result of Sum(A1:A2): " + results.toString());
	}
}
