package com.aspose.cells.examples.DrawingObjects.controls;

import com.aspose.cells.Color;
import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.MsoFillFormat;
import com.aspose.cells.MsoLineDashStyle;
import com.aspose.cells.MsoLineFormat;
import com.aspose.cells.MsoLineStyle;
import com.aspose.cells.PlacementType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class AddingRectangleControl {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(AddingRectangleControl.class);
		// Instantiate a new Workbook.
		Workbook excelbook = new Workbook();

		// Add a rectangle control.
		com.aspose.cells.RectangleShape rectangle = (com.aspose.cells.RectangleShape) excelbook.getWorksheets().get(0)
				.getShapes().addShape(MsoDrawingType.RECTANGLE, 3, 2, 0, 0, 70, 130);

		// Set the placement of the rectangle.
		rectangle.setPlacement(PlacementType.FREE_FLOATING);

		// Set the fill format.
		MsoFillFormat fillformat = rectangle.getFillFormat();
		fillformat.setForeColor(Color.getOlive());

		// Set the line style.
		MsoLineFormat linestyle = rectangle.getLineFormat();
		linestyle.setStyle(MsoLineStyle.THICK_THIN);

		// Set the line weight.
		linestyle.setWeight(4);

		// Set the color of the line.
		linestyle.setForeColor(Color.getBlue());

		// Set the dash style of the rectangle.
		linestyle.setDashStyle(MsoLineDashStyle.SOLID);

		// Save the excel file.
		excelbook.save(dataDir + "tstrectangle.xls");
	}
}
