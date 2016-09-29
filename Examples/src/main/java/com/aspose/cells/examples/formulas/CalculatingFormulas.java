package com.aspose.cells.examples.formulas;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class CalculatingFormulas {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CalculatingFormulas.class) + "formulas/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Adding a new worksheet to the Excel object
		int sheetIndex = workbook.getWorksheets().add();

		// Obtaining the reference of the newly added worksheet by passing its sheet index
		Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);

		// Adding a value to "A1" cell
		worksheet.getCells().get("A1").putValue(1);

		// Adding a value to "A2" cell
		worksheet.getCells().get("A2").putValue(2);

		// Adding a value to "A3" cell
		worksheet.getCells().get("A3").putValue(3);

		// Adding a SUM formula to "A4" cell
		worksheet.getCells().get("A4").setFormula("=SUM(A1:A3)");

		// Calculating the results of formulas
		workbook.calculateFormula();

		// Get the calculated value of the cell
		String value = worksheet.getCells().get("A4").getStringValue();

		// Saving the Excel file
		workbook.save(dataDir + "CalculatingFormulas_out.xls");
	}
}
