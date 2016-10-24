package com.aspose.cells.examples.data;

import com.aspose.cells.Name;
import com.aspose.cells.Workbook;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;
import com.aspose.cells.examples.charts.CreateChart;

public class SetSimpleFormulaforNamedRange {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(CreateChart.class) + "data/";
		// Create an instance of Workbook
		Workbook book = new Workbook();

		// Get the WorksheetCollection
		WorksheetCollection worksheets = book.getWorksheets();

		// Add a new Named Range with name "myName"
		int index = worksheets.getNames().add("myName");

		// Access the newly created Named Range
		Name name = worksheets.getNames().get(index);

		// Set RefersTo property of the Named Range to a formula
		// Formula references another cell in the same worksheet
		name.setRefersTo("=Sheet1!$A$3");

		// Set the formula in the cell A1 to the newly created Named Range
		worksheets.get(0).getCells().get("A1").setFormula("myName");

		// Insert the value in cell A3 which is being referenced in the Named
		// Range
		worksheets.get(0).getCells().get("A3").putValue("This is the value of A3");

		// Calculate formulas
		book.calculateFormula();

		// Save the result in XLSX format
		book.save(dataDir + "SetSimpleFormulaNamedRange_out.xlsx");
	}
}
