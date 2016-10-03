package com.aspose.cells.examples.DrawingObjects;

import com.aspose.cells.CheckBox;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AddingCheckBoxControl {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddingCheckBoxControl.class) + "DrawingObjects/";
		// Instantiate a new Workbook.
		Workbook workbook = new Workbook();

		// Get the first worksheet in the book.
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Add a checkbox to the first worksheet in the workbook.
		int checkBoxIndex = worksheet.getCheckBoxes().add(5, 5, 100, 120);
		CheckBox checkBox = worksheet.getCheckBoxes().get(checkBoxIndex);

		// Set its text string.
		checkBox.setText("Check it!");

		// Put a value into B1 cell.
		worksheet.getCells().get("B1").setValue("LnkCell");

		// Set B1 cell as a linked cell for the checkbox.
		checkBox.setLinkedCell("=B1");

		// Save the excel file.
		workbook.save(dataDir + "AddingCheckBoxControl_out.xls");
	}
}
