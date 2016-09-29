package com.aspose.cells.examples.articles;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class CalculationOfArrayFormula {
	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(CalculationOfArrayFormula.class) + "articles/";
		// Create workbook from source excel file
		Workbook workbook = new Workbook(dataDir + "DataTable.xlsx");

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// When you will put 100 in B1, then all Data Table values formatted as Yellow will become 120
		worksheet.getCells().get("B1").putValue(100);

		// Calculate formula, now it also calculates Data Table array formula
		workbook.calculateFormula();

		// Save the workbook in pdf format
		workbook.save(dataDir + "COfAFormula_out.pdf");

	}
}
