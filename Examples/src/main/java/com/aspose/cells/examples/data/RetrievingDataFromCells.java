package com.aspose.cells.examples.data;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class RetrievingDataFromCells {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(RetrievingDataFromCells.class) + "data/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Accessing the worksheet
		com.aspose.cells.Worksheet worksheet = workbook.getWorksheets().get(0);
		com.aspose.cells.Cells cells = worksheet.getCells();

		// get cell from cells collection
		com.aspose.cells.Cell cell = cells.get("A5");

		switch (cell.getType()) {
		case com.aspose.cells.CellValueType.IS_BOOL:
			System.out.println("Boolean Value: " + cell.getValue());
			break;
		case com.aspose.cells.CellValueType.IS_DATE_TIME:
			System.out.println("Date Value: " + cell.getValue());
			break;
		case com.aspose.cells.CellValueType.IS_NUMERIC:
			System.out.println("Numeric Value: " + cell.getValue());
			break;
		case com.aspose.cells.CellValueType.IS_STRING:
			System.out.println("String Value: " + cell.getValue());
			break;
		case com.aspose.cells.CellValueType.IS_NULL:
			System.out.println("Null Value");
			break;
		}

	}
}
