package com.aspose.cells.examples.DrawingObjects;

import com.aspose.cells.Cells;
import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.PlacementType;
import com.aspose.cells.SelectionType;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AddingListBoxControl {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddingListBoxControl.class) + "DrawingObjects/";
		// Create a new Workbook.
		Workbook workbook = new Workbook();

		// Get the first worksheet.
		Worksheet sheet = workbook.getWorksheets().get(0);

		// Get the worksheet cells collection.
		Cells cells = sheet.getCells();

		// Input a value.
		cells.get("B3").setValue("Choose Dept:");

		Style style = cells.get("B3").getStyle();
		style.getFont().setBold(true);
		// Set it bold.
		cells.get("B3").setStyle(style);

		// Input some values that denote the input range for the combo box.
		cells.get("A2").setValue("Sales");
		cells.get("A3").setValue("Finance");
		cells.get("A4").setValue("MIS");
		cells.get("A5").setValue("R&D");
		cells.get("A6").setValue("Marketing");
		cells.get("A7").setValue("HRA");

		// Add a new list box.
		com.aspose.cells.ListBox listBox = (com.aspose.cells.ListBox) sheet.getShapes()
				.addShape(MsoDrawingType.LIST_BOX, 3, 3, 1, 0, 100, 122);

		// Set the linked cell;
		listBox.setLinkedCell("A1");

		// Set the input range.
		listBox.setInputRange("=A2:A7");

		// Set the Placement Type, the way the list box is attached to the cells.
		listBox.setPlacement(PlacementType.FREE_FLOATING);

		// Set the list box with 3-D shading.
		listBox.setShadow(true);

		// Set the selection type.
		listBox.setSelectionType(SelectionType.SINGLE);

		// AutoFit Columns
		sheet.autoFitColumns();

		// Saves the file.
		workbook.save(dataDir + "AddingListBoxControl_out.xls");
	}
}
