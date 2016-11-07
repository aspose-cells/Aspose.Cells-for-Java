package com.aspose.cells.examples.DrawingObjects;

import com.aspose.cells.Color;
import com.aspose.cells.FillFormat;
import com.aspose.cells.GradientStyleType;
import com.aspose.cells.LineFormat;
import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.MsoLineDashStyle;
import com.aspose.cells.MsoLineStyle;
import com.aspose.cells.PlacementType;
import com.aspose.cells.TextBox;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AddingTextBoxControl {
	public static void main(String[] args) throws Exception {
		// ExStart:AddingTextBoxControl
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddingTextBoxControl.class) + "DrawingObjects/";
		// Instantiate a new Workbook.
		Workbook workbook = new Workbook();

		// Get the first worksheet in the book.
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Get the textbox object.
		int textboxIndex = worksheet.getTextBoxes().add(2, 1, 160, 200);
		TextBox textbox0 = worksheet.getTextBoxes().get(textboxIndex);

		// Fill the text.
		textbox0.setText("ASPOSE______The .NET & JAVA Component Publisher!");

		// Set the placement.
		textbox0.setPlacement(PlacementType.FREE_FLOATING);

		// Set the font color.
		textbox0.getFont().setColor(Color.getBlue());

		// Set the font to bold.
		textbox0.getFont().setBold(true);

		// Set the font size.
		textbox0.getFont().setSize(14);

		// Set font attribute to italic.
		textbox0.getFont().setItalic(true);

		// Add a hyperlink to the textbox.
		textbox0.addHyperlink("http://www.aspose.com/");

		// Get the filformat of the textbox.
		FillFormat fillformat = textbox0.getFill();

		// Set the fillcolor.
		fillformat.setOneColorGradient(Color.getSilver(), 1, GradientStyleType.HORIZONTAL, 1);

		// Get the lineformat type of the textbox.
		LineFormat lineformat = textbox0.getLine();

		// Set the line style.
		lineformat.setDashStyle(MsoLineStyle.THIN_THICK);

		// Set the line weight.
		lineformat.setWeight(6);

		// Set the dash style to squaredot.
		lineformat.setDashStyle(MsoLineDashStyle.SQUARE_DOT);

		// Get the second textbox.
		TextBox textbox1 = (com.aspose.cells.TextBox) worksheet.getShapes().addShape(MsoDrawingType.TEXT_BOX, 15, 0, 4,
				0, 85, 120);

		// Input some text to it.
		textbox1.setText("This is another simple text box");

		// Set the placement type as the textbox will move and resize with cells.
		textbox1.setPlacement(PlacementType.MOVE_AND_SIZE);

		// Save the excel file.
		workbook.save(dataDir + "AddingTextBoxControl_out.xls");
		// ExEnd:AddingTextBoxControl
	}
}
