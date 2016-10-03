package com.aspose.cells.examples.data;

import com.aspose.cells.Cells;
import com.aspose.cells.Range;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class IdentifyCellsinNamedRange {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(IdentifyCellsinNamedRange.class) + "data/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		WorksheetCollection worksheets = workbook.getWorksheets();

		// Accessing the first worksheet in the Excel file
		Worksheet sheet = worksheets.get(0);
		Cells cells = sheet.getCells();

		// Getting the specified named range
		Range range = worksheets.getRangeByName("TestRange");

		// Identify range cells.
		System.out.println("First Row : " + range.getFirstRow());
		System.out.println("First Column : " + range.getFirstColumn());
		System.out.println("Row Count : " + range.getRowCount());
		System.out.println("Column Count : " + range.getColumnCount());

	}
}
