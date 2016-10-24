package com.aspose.cells.examples.data;

import com.aspose.cells.Name;
import com.aspose.cells.Workbook;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;
import com.aspose.cells.examples.charts.CreateChart;

public class SetComplexFormulaforNamedRange {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(CreateChart.class) + "data/";
		// Create an instance of Workbook
		Workbook book = new Workbook();

		// Get the WorksheetCollection
		WorksheetCollection worksheets = book.getWorksheets();

		// Add a new Named Range with name "data"
		int index = worksheets.getNames().add("data");

		// Access the newly created Named Range from the collection
		Name data = worksheets.getNames().get(index);

		// Set RefersTo property of the Named Range to a cell range in same
		// worksheet
		data.setRefersTo("=Sheet1!$A$1:$A$10");

		// Add another Named Range with name "range"
		index = worksheets.getNames().add("range");

		// Access the newly created Named Range from the collection
		Name range = worksheets.getNames().get(index);

		// Set RefersTo property to a formula using the Named Range data
		range.setRefersTo("=INDEX(data,Sheet1!$A$1,1):INDEX(data,Sheet1!$A$1,9)");
	}
}
