package com.aspose.cells.examples.DrawingObjects.controls;

import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.MsoLineDashStyle;
import com.aspose.cells.MsoLineFormat;
import com.aspose.cells.PlacementType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AddingLineControl {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(AddingLineControl.class);

		// Instantiate a new Workbook.
		Workbook workbook = new Workbook();

		// Get the first worksheet in the book.
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Add a new line to the worksheet.
		com.aspose.cells.LineShape line1 = (com.aspose.cells.LineShape) worksheet.getShapes()
				.addShape(MsoDrawingType.LINE, 5, 1, 0, 0, 0, 250);

		// Set the line dash style
		MsoLineFormat shapeline = line1.getLineFormat();
		shapeline.setDashStyle(MsoLineDashStyle.SOLID);

		// Set the placement.
		line1.setPlacement(PlacementType.FREE_FLOATING);

		// Add another line to the worksheet.
		com.aspose.cells.LineShape line2 = (com.aspose.cells.LineShape) worksheet.getShapes()
				.addShape(MsoDrawingType.LINE, 7, 1, 0, 0, 85, 250);

		// Set the line dash style.
		shapeline = line2.getLineFormat();
		shapeline.setDashStyle(MsoLineDashStyle.DASH_LONG_DASH);

		// Set the weight of the line.
		MsoLineFormat lineformat = line2.getLineFormat();
		lineformat.setWeight(4);

		// Set the placement.
		line2.setPlacement(PlacementType.FREE_FLOATING);

		// Add the third line to the worksheet.
		com.aspose.cells.LineShape line3 = (com.aspose.cells.LineShape) worksheet.getShapes()
				.addShape(MsoDrawingType.LINE, 13, 1, 0, 0, 0, 250);

		// Set the line dash style
		shapeline = line1.getLineFormat();
		shapeline.setDashStyle(MsoLineDashStyle.SOLID);

		// Set the placement.
		line3.setPlacement(PlacementType.FREE_FLOATING);

		// Make the gridlines invisible in the first worksheet.
		workbook.getWorksheets().get(0).setGridlinesVisible(false);

		// Save the excel file.
		workbook.save(dataDir + "tstlines.xls");
	}
}
