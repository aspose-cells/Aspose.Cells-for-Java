package com.aspose.cells.examples.articles;

import com.aspose.cells.Cell;
import com.aspose.cells.ListObject;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AccessingTablefromCell {
	public static void main(String[] args) throws Exception {
		// ExStart:AccessingTablefromCell
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(AccessingTablefromCell.class);
		// Create workbook from source Excel file
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access cell D5 which lies inside the table
		Cell cell = worksheet.getCells().get("D5");

		// Put value inside the cell D5
		cell.putValue("D5 Data");

		// Access the Table from this cell
		ListObject table = cell.getTable();

		// Add some value using Row and Column Offset
		table.putCellValue(2, 2, "Offset [2,2]");

		// Save the workbook
		workbook.save(dataDir + "output.xlsx");
		// ExEnd:AccessingTablefromCell
	}
}
