package com.aspose.cells.examples.data;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class WholeNumberDataValidation {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(WholeNumberDataValidation.class) + "data/";

		// Instantiating an Workbook object
		Workbook workbook = new Workbook();
		WorksheetCollection worksheets = workbook.getWorksheets();

		// Accessing the Validations collection of the worksheet
		Worksheet worksheet = worksheets.get(0);

		// Applying the validation to a range of cells from A1 to B2 using the
		// CellArea structure
		CellArea area = new CellArea();
		area.StartRow = 0;
		area.StartColumn = 0;
		area.EndRow = 1;
		area.EndColumn = 1;

		ValidationCollection validations = worksheet.getValidations();

		// Creating a Validation object
		int index = validations.add(area);
		Validation validation = validations.get(index);

		// Setting the validation type to whole number
		validation.setType(ValidationType.WHOLE_NUMBER);

		// Setting the operator for validation to Between
		validation.setOperator(OperatorType.BETWEEN);

		// Setting the minimum value for the validation
		validation.setFormula1("10");

		// Setting the maximum value for the validation
		validation.setFormula2("1000");

		// Saving the Excel file
		workbook.save(dataDir + "WNDValidation_out.xls");

		// Print message
		System.out.println("Process completed successfully");

	}
}
