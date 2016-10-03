package com.aspose.cells.examples.data;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class AddingDataToCells {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddingDataToCells.class) + "data/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Accessing the added worksheet in the Excel file
		int sheetIndex = workbook.getWorksheets().add();
		com.aspose.cells.Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);
		com.aspose.cells.Cells cells = worksheet.getCells();

		// Adding a string value to the cell
		com.aspose.cells.Cell cell = cells.get("A1");
		cell.setValue("Hello World");

		// Adding a double value to the cell
		cell = cells.get("A2");
		cell.setValue(20.5);

		// Adding an integer value to the cell
		cell = cells.get("A3");
		cell.setValue(15);

		// Adding a boolean value to the cell
		cell = cells.get("A4");
		cell.setValue(true);

		// Adding a date/time value to the cell
		cell = cells.get("A5");
		cell.setValue(java.util.Calendar.getInstance());

		// Setting the display format of the date
		com.aspose.cells.Style style = cell.getStyle();
		style.setNumber(15);
		cell.setStyle(style);

		// Saving the Excel file
		workbook.save(dataDir + "AddingDataToCells_out.xls");

		// Print message
		System.out.println("Data Added Successfully");

	}
}
