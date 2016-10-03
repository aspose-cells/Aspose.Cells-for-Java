package com.aspose.cells.examples.data;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class DateDataValidation {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(DateDataValidation.class) + "data/";

		// Create a workbook.
		Workbook workbook = new Workbook();

		// Obtain the cells of the first worksheet.
		Cells cells = workbook.getWorksheets().get(0).getCells();

		// Put a string value into the A1 cell.
		cells.get("A1").setValue("Please enter Date b/w 1/1/1970 and 12/31/1999");

		// Wrap the text.
		Style style = cells.get("A1").getStyle();
		style.setTextWrapped(true);
		cells.get("A1").setStyle(style);

		// Set row height and column width for the cells.
		cells.setRowHeight(0, 31);
		cells.setColumnWidth(0, 35);

		// Set a collection of CellArea which contains the data validation
		// settings.
		CellArea area = new CellArea();

		area.StartRow = 0;
		area.StartColumn = 1;
		area.EndRow = 0;
		area.EndColumn = 1;

		// Get the validations collection.
		ValidationCollection validations = workbook.getWorksheets().get(0).getValidations();

		// Add a new validation.
		int i = validations.add(area);
		Validation validation = validations.get(i);

		// Set the data validation type.
		validation.setType(ValidationType.DATE);

		// Set the operator for the data validation
		validation.setOperator(OperatorType.BETWEEN);

		// Set the value or expression associated with the data validation.
		validation.setFormula1("1/1/1970");

		// The value or expression associated with the second part of the data
		// validation.
		validation.setFormula2("12/31/1999");

		// Enable the error.
		validation.setShowError(true);

		// Set the validation alert style.
		validation.setAlertStyle(ValidationAlertType.STOP);

		// Set the title of the data-validation error dialog box
		validation.setErrorTitle("Date Error");

		// Set the data validation error message.
		validation.setErrorMessage("Enter a Valid Date");

		// Set and enable the data validation input message.
		validation.setInputMessage("Date Validation Type");
		validation.setIgnoreBlank(true);
		validation.setShowInput(true);

		// Save the excel file.
		workbook.save(dataDir + "DDValidation_out.xls", FileFormatType.EXCEL_97_TO_2003);

		// Print message
		System.out.println("Process completed successfully");

	}
}
