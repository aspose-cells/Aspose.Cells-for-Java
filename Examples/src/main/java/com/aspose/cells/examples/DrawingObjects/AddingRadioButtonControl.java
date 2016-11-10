package com.aspose.cells.examples.DrawingObjects;

import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.GradientStyleType;
import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.MsoLineDashStyle;
import com.aspose.cells.MsoLineStyle;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AddingRadioButtonControl {
	public static void main(String[] args) throws Exception {
		// ExStart:AddingRadioButtonControl
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddingRadioButtonControl.class) + "DrawingObjects/";

		// Create a new Workbook.
		Workbook workbook = new Workbook();

		// Get the first worksheet.
		Worksheet sheet = workbook.getWorksheets().get(0);

		// Get the worksheet cells collection.
		Cells cells = sheet.getCells();

		// Insert a value.
		cells.get("C2").setValue("Age Groups");

		Style style = cells.get("B3").getStyle();
		style.getFont().setBold(true);
		// Set it bold.
		cells.get("C2").setStyle(style);

		// Add a radio button to the first sheet.
		com.aspose.cells.RadioButton radio1 = (com.aspose.cells.RadioButton) sheet.getShapes()
				.addShape(MsoDrawingType.RADIO_BUTTON, 3, 0, 1, 0, 20, 100);

		// Set its text string.
		radio1.setText("20-29");

		// Set A1 cell as a linked cell for the radio button.
		radio1.setLinkedCell("A1");

		// Make the radio button 3-D.
		radio1.setShadow(true);

		// Set the foreground color of the radio button.
		radio1.getFill().setOneColorGradient(Color.getGreen(), 1, GradientStyleType.HORIZONTAL, 1);

		// set the line style of the radio button.
		radio1.getLine().setDashStyle(MsoLineStyle.THICK_THIN);

		// Set the weight of the radio button.
		radio1.getLine().setWeight(4);

		// Set the line color of the radio button.
		radio1.getLine().setOneColorGradient(Color.getBlue(), 1, GradientStyleType.HORIZONTAL, 1);

		// Set the dash style of the radio button.
		radio1.getLine().setDashStyle(MsoLineDashStyle.SOLID);

		// Add another radio button to the first sheet.
		com.aspose.cells.RadioButton radio2 = (com.aspose.cells.RadioButton) sheet.getShapes()
				.addShape(MsoDrawingType.RADIO_BUTTON, 6, 0, 1, 0, 20, 100);

		// Set its text string.
		radio2.setText("30-39");

		// Set A1 cell as a linked cell for the radio button.
		radio2.setLinkedCell("A1");

		// Make the radio button 3-D.
		radio2.setShadow(true);

		// Set the foreground color of the radio button.
		radio2.getFill().setOneColorGradient(Color.getGreen(), 1, GradientStyleType.HORIZONTAL, 1);

		// set the line style of the radio button.
		radio2.getLine().setDashStyle(MsoLineStyle.THICK_THIN);

		// Set the weight of the radio button.
		radio2.getLine().setWeight(4);

		// Set the line color of the radio button.
		radio2.getLine().setOneColorGradient(Color.getBlue(), 1, GradientStyleType.HORIZONTAL, 1);

		// Set the dash style of the radio button.
		radio2.getLine().setDashStyle(MsoLineDashStyle.SOLID);


		// Add another radio button to the first sheet.
		com.aspose.cells.RadioButton radio3 = (com.aspose.cells.RadioButton) sheet.getShapes()
				.addShape(MsoDrawingType.RADIO_BUTTON, 9, 0, 1, 0, 20, 100);

		// Set its text string.
		radio3.setText("40-49");

		// Set A1 cell as a linked cell for the radio button.
		radio3.setLinkedCell("A1");

		// Make the radio button 3-D.
		radio3.setShadow(true);

		// Set the foreground color of the radio button.
		radio3.getFill().setOneColorGradient(Color.getGreen(), 1, GradientStyleType.HORIZONTAL, 1);

		// set the line style of the radio button.
		radio3.getLine().setDashStyle(MsoLineStyle.THICK_THIN);

		// Set the weight of the radio button.
		radio3.getLine().setWeight(4);

		// Set the line color of the radio button.
		radio3.getLine().setOneColorGradient(Color.getBlue(), 1, GradientStyleType.HORIZONTAL, 1);

		// Set the dash style of the radio button.
		radio3.getLine().setDashStyle(MsoLineDashStyle.SOLID);

		// Save the excel file.
		workbook.save(dataDir + "ARBControl_out.xls");
		// ExEnd:AddingRadioButtonControl
	}
}
