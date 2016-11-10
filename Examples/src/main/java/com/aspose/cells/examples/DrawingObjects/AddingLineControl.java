package com.aspose.cells.examples.DrawingObjects;

import com.aspose.cells.LineFormat;
import com.aspose.cells.LineShape;
import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.MsoLineDashStyle;
import com.aspose.cells.PlacementType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AddingLineControl {
	public static void main(String[] args) throws Exception {
		// ExStart:AddingLineControl
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddingLineControl.class) + "DrawingObjects/";

		//Instantiate a new Workbook.
		Workbook workbook = new Workbook();

		//Get the first worksheet in the book.
		Worksheet worksheet = workbook.getWorksheets().get(0);

		//Add a new line to the worksheet.
		LineShape line1  = (LineShape)worksheet.getShapes().addShape(MsoDrawingType.LINE,5, 1,0,0, 0, 250);
		line1.setHasLine(true);

		//Set the line dash style
		LineFormat shapeline = line1.getLine();
		shapeline.setDashStyle(MsoLineDashStyle.SOLID);

		//Set the placement.
		line1.setPlacement(PlacementType.FREE_FLOATING);

		//Add another line to the worksheet.
		LineShape line2  = (LineShape) worksheet.getShapes().addShape(MsoDrawingType.LINE, 7, 1,0,0, 85, 250);
		line2.setHasLine(true);

		//Set the line dash style.
		shapeline = line2.getLine();
		shapeline.setDashStyle(MsoLineDashStyle.DASH_LONG_DASH);
		shapeline.setWeight(4);

		//Set the placement.
		line2.setPlacement(PlacementType.FREE_FLOATING);

		//Add the third line to the worksheet.
		LineShape line3  = (LineShape)worksheet.getShapes().addShape(MsoDrawingType.LINE, 13, 1,0,0, 0, 250);
		line3.setHasLine(true);

		//Set the line dash style
		shapeline = line1.getLine();
		shapeline.setDashStyle(MsoLineDashStyle.SOLID);

		//Set the placement.
		line3.setPlacement(PlacementType.FREE_FLOATING);

		//Save the excel file.
		workbook.save(dataDir + "tstlines.xls");
		// ExEnd:AddingLineControl
	}
}
