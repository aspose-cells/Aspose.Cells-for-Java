package com.aspose.cells.examples.articles;

import com.aspose.cells.Cell;
import com.aspose.cells.Color;
import com.aspose.cells.PivotTable;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class GetCellObject {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(GetCellObject.class) + "articles/";
		// Create workbook object from source excel file
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access first pivot table inside the worksheet
		PivotTable pivotTable = worksheet.getPivotTables().get(0);

		// Access cell by display name of 2nd data field of the pivot table
		String displayName = pivotTable.getDataFields().get(1).getDisplayName();
		Cell cell = pivotTable.getCellByDisplayName(displayName);

		// Access cell style and set its fill color and font color
		Style style = cell.getStyle();
		style.setForegroundColor(Color.getLightBlue());
		style.getFont().setColor(Color.getBlack());

		// Set the style of the cell
		pivotTable.format(cell.getRow(), cell.getColumn(), style);

		// Save workbook
		workbook.save(dataDir + "GetCellObject_out.xlsx");

	}
}
