package com.aspose.cells.examples.data;

import com.aspose.cells.Name;
import com.aspose.cells.Workbook;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;
import com.aspose.cells.examples.charts.CreateChart;

public class NamedRangeToSumValues {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(CreateChart.class) + "data/";
		// Create an instance of Workbook
		Workbook book = new Workbook();

		// Get the WorksheetCollection
		WorksheetCollection worksheets = book.getWorksheets();

		// Insert some data in cell A1 of Sheet1
		worksheets.get("Sheet1").getCells().get("A1").putValue(10);

		// Add a new Worksheet and insert a value to cell A1
		worksheets.get(worksheets.add()).getCells().get("A1").putValue(10);

		// Add a new Named Range with name "range"
		int index = worksheets.getNames().add("range");

		// Access the newly created Named Range from the collection
		Name range = worksheets.getNames().get(index);

		// Set RefersTo property of the Named Range to a SUM function
		range.setRefersTo("=SUM(Sheet1!$A$1,Sheet2!$A$1)");

		// Insert the Named Range as formula to 3rd worksheet
		worksheets.get(worksheets.add()).getCells().get("A1").setFormula("range");

		// Calculate formulas
		book.calculateFormula();

		// Save the result in XLSX format
		book.save(dataDir + "NamedRangeToSumValues_out.xlsx");
	}
}
