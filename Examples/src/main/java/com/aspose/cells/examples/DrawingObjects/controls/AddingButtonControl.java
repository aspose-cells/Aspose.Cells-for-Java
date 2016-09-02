package com.aspose.cells.examples.DrawingObjects.controls;

import com.aspose.cells.Color;
import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.PlacementType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AddingButtonControl {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(AddingButtonControl.class);
		// Create a new Workbook.
		Workbook workbook = new Workbook();

		// Get the first worksheet.
		Worksheet sheet = workbook.getWorksheets().get(0);

		// Add a new button to the worksheet.
		com.aspose.cells.Button button = (com.aspose.cells.Button) sheet.getShapes().addShape(MsoDrawingType.BUTTON, 2,
				2, 2, 0, 20, 80);

		// Set the caption of the button.
		button.setText("Aspose");

		// Set the Placement Type, the way the button is attached to the cells.
		button.setPlacement(PlacementType.FREE_FLOATING);

		// Set the font name.
		button.getFont().setName("Tahoma");

		// Set the caption string bold.
		button.getFont().setBold(true);

		// Set the color to blue.
		button.getFont().setColor(Color.getBlue());

		// Set the hyperlink for the button.
		button.addHyperlink("http://www.aspose.com/");

		// Saves the file.
		workbook.save(dataDir + "tstbutton.xls");
	}
}
