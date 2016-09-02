package com.aspose.cells.examples.DrawingObjects.controls;

import com.aspose.cells.Color;
import com.aspose.cells.MsoArrowheadLength;
import com.aspose.cells.MsoArrowheadStyle;
import com.aspose.cells.MsoArrowheadWidth;
import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.MsoLineDashStyle;
import com.aspose.cells.MsoLineFormat;
import com.aspose.cells.PlacementType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AddinganArrowHead {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(AddinganArrowHead.class);
		// Instantiate a new Workbook.
		Workbook workbook = new Workbook();

		// Get the first worksheet in the book.
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Add a line to the worksheet.
		com.aspose.cells.LineShape line2 = (com.aspose.cells.LineShape) worksheet.getShapes()
				.addShape(MsoDrawingType.LINE, 7, 1, 1, 0, 85, 250);
		MsoLineFormat lineformat = line2.getLineFormat();

		// Set the line color
		lineformat.setForeColor(Color.getBlue());

		// Set the line style.
		lineformat.setDashStyle(MsoLineDashStyle.SOLID);

		// Set the weight of the line.
		lineformat.setWeight(3);

		// Set the placement.
		line2.setPlacement(PlacementType.FREE_FLOATING);

		// Set the line arrows.
		line2.setEndArrowheadWidth(MsoArrowheadWidth.MEDIUM);
		line2.setEndArrowheadStyle(MsoArrowheadStyle.ARROW);
		line2.setEndArrowheadLength(MsoArrowheadLength.MEDIUM);

		line2.setBeginArrowheadWidth(MsoArrowheadWidth.NARROW);
		line2.setBeginArrowheadStyle(MsoArrowheadStyle.ARROW_DIAMOND);
		line2.setBeginArrowheadLength(MsoArrowheadLength.MEDIUM);

		// Make the gridlines invisible in the first worksheet.
		workbook.getWorksheets().get(0).setGridlinesVisible(false);

		// Save the excel file.
		workbook.save(dataDir + "arrowlinetest.xls");
	}
}
