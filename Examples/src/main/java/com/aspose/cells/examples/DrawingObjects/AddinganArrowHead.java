package com.aspose.cells.examples.DrawingObjects;

import com.aspose.cells.Color;
import com.aspose.cells.FillType;
import com.aspose.cells.LineShape;
import com.aspose.cells.MsoArrowheadLength;
import com.aspose.cells.MsoArrowheadStyle;
import com.aspose.cells.MsoArrowheadWidth;
import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.PlacementType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AddinganArrowHead {
	public static void main(String[] args) throws Exception {
		// ExStart:AddinganArrowHead
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddinganArrowHead.class) + "DrawingObjects/";
		//Instantiate a new Workbook.
		Workbook workbook = new Workbook();

		//Get the first worksheet in the book.
		Worksheet worksheet = workbook.getWorksheets().get(0);

		//Add a line to the worksheet
		LineShape line2 = (LineShape) worksheet.getShapes().addShape(MsoDrawingType.LINE, 7, 0, 1, 0, 85, 250);

		//Set the line color
		line2.getLine().setFillType(FillType.SOLID);
		line2.getLine().getSolidFill().setColor(Color.getRed());

		//Set the weight of the line.
		line2.getLine().setWeight(3);

		//Set the placement.
		line2.setPlacement(PlacementType.FREE_FLOATING);

		//Set the line arrows.
		line2.getLine().setEndArrowheadWidth(MsoArrowheadWidth.MEDIUM);
		line2.getLine().setEndArrowheadStyle(MsoArrowheadStyle.ARROW);
		line2.getLine().setEndArrowheadLength(MsoArrowheadLength.MEDIUM);
		line2.getLine().setBeginArrowheadStyle(MsoArrowheadStyle.ARROW_DIAMOND);
		line2.getLine().setBeginArrowheadLength(MsoArrowheadLength.MEDIUM);

		//Make the gridlines invisible in the first worksheet.
		workbook.getWorksheets().get(0).setGridlinesVisible(false);

		//Save the excel file.
		workbook.save(dataDir + "AddinganArrowHead_out.xlsx");
		// ExEnd:AddinganArrowHead
	}
}
