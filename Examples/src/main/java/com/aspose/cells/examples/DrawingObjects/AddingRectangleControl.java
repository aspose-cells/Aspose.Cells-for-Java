package com.aspose.cells.examples.DrawingObjects;

import com.aspose.cells.Color;
import com.aspose.cells.FillFormat;
import com.aspose.cells.GradientStyleType;
import com.aspose.cells.LineFormat;
import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.MsoLineDashStyle;
import com.aspose.cells.MsoLineStyle;
import com.aspose.cells.PlacementType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class AddingRectangleControl {
	public static void main(String[] args) throws Exception {
		// ExStart:AddingRectangleControl
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddingRectangleControl.class) + "DrawingObjects/";
		// Instantiate a new Workbook.
		Workbook excelbook = new Workbook();

		// Add a rectangle control.
		com.aspose.cells.RectangleShape rectangle = (com.aspose.cells.RectangleShape) excelbook.getWorksheets().get(0)
				.getShapes().addShape(MsoDrawingType.RECTANGLE, 3, 2, 0, 0, 70, 130);

		// Set the placement of the rectangle.
		rectangle.setPlacement(PlacementType.FREE_FLOATING);

		// Set the fill format.
		FillFormat fillformat = rectangle.getFill();
		fillformat.setOneColorGradient(Color.getOlive(), 1, GradientStyleType.HORIZONTAL, 1);

		// Set the line style.
		LineFormat linestyle = rectangle.getLine();
		linestyle.setDashStyle(MsoLineStyle.THICK_THIN);

		// Set the line weight.
		linestyle.setWeight(4);

		// Set the color of the line.
		linestyle.setOneColorGradient(Color.getBlue(), 1, GradientStyleType.HORIZONTAL, 1);

		// Set the dash style of the rectangle.
		linestyle.setDashStyle(MsoLineDashStyle.SOLID);

		// Save the excel file.
		excelbook.save(dataDir + "AddingRectangleControl_out.xls");
		// ExEnd:AddingRectangleControl
	}
}
