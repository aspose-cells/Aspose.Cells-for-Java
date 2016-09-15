package com.aspose.cells.examples.articles;

import com.aspose.cells.Cell;
import com.aspose.cells.Validation;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class GetValidationAppliedonCell {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(GetValidationAppliedonCell.class) + "articles/";
		// Instantiate the workbook from sample Excel file
		Workbook workbook = new Workbook(dataDir + "book1.xlsx");

		// Access its first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Cell C1 has the Decimal Validation applied on it.
		// It can take only the values Between 10 and 20
		Cell cell = worksheet.getCells().get("C1");

		// Access the valditation applied on this cell
		Validation validation = cell.getValidation();

		// Read various properties of the validation
		System.out.println("Reading Properties of Validation");
		System.out.println("--------------------------------");
		System.out.println("Type: " + validation.getType());
		System.out.println("Operator: " + validation.getOperator());
		System.out.println("Formula1: " + validation.getFormula1());
		System.out.println("Formula2: " + validation.getFormula2());
		System.out.println("Ignore blank: " + validation.getIgnoreBlank());

	}
}
