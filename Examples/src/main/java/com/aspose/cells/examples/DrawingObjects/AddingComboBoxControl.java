package com.aspose.cells.examples.DrawingObjects;

import com.aspose.cells.Cells;
import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AddingComboBoxControl {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddingComboBoxControl.class) + "DrawingObjects/";
		// Create a new Workbook.
		Workbook workbook = new Workbook();

		// Get the first worksheet.
		Worksheet sheet = workbook.getWorksheets().get(0);

		// Get the worksheet cells collection.
		Cells cells = sheet.getCells();

		// Input a value.
		cells.get("B3").setValue("Employee:");

		Style style = cells.get("B3").getStyle();
		style.getFont().setBold(true);
		// Set it bold.
		cells.get("B3").setStyle(style);

		// Input some values that denote the input range for the combo box.
		cells.get("A2").setValue("Emp001");
		cells.get("A3").setValue("Emp002");
		cells.get("A4").setValue("Emp003");
		cells.get("A5").setValue("Emp004");
		cells.get("A6").setValue("Emp005");
		cells.get("A7").setValue("Emp006");

		// Add a new combo box.
		com.aspose.cells.ComboBox comboBox = (com.aspose.cells.ComboBox) sheet.getShapes()
				.addShape(MsoDrawingType.COMBO_BOX, 3, 0, 1, 0, 20, 100);

		// Set the linked cell;
		comboBox.setLinkedCell("A1");

		// Set the input range.
		comboBox.setInputRange("=A2:A7");

		// Set no. of list lines displayed in the combo box's list portion.
		comboBox.setDropDownLines(5);

		// Set the combo box with 3-D shading.
		comboBox.setShadow(true);

		// AutoFit Columns
		sheet.autoFitColumns();

		// Saves the file.
		workbook.save(dataDir + "AddingComboBoxControl_out.xls");
	}
}
