package com.aspose.cells.examples.DrawingObjects.controls;

import com.aspose.cells.Color;
import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.PlacementType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AddingLabelControl {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddingLabelControl.class) + "DrawingObjects/";

		// Create a new Workbook.
		Workbook workbook = new Workbook();

		// Get the first worksheet.
		Worksheet sheet = workbook.getWorksheets().get(0);

		// Add a new label to the worksheet.
		com.aspose.cells.Label label = (com.aspose.cells.Label) sheet.getShapes().addShape(MsoDrawingType.LABEL, 2, 2,
				2, 0, 60, 120);

		// Set the caption of the label.
		label.setText("This is a Label");

		// Set the Placement Type, the way the label is attached to the cells.
		label.setPlacement(PlacementType.FREE_FLOATING);

		// Set the fill color of the label.
		label.getFillFormat().setForeColor(Color.getYellow());

		// Saves the file.
		workbook.save(dataDir + "ALControl-out.xls");
	}
}
