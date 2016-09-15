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

public class AddingArcControl {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddingArcControl.class) + "DrawingObjects/";
		// Instantiate a new Workbook.
		Workbook excelbook = new Workbook();
		// Add an arc shape.
		com.aspose.cells.ArcShape arc1 = (com.aspose.cells.ArcShape) excelbook.getWorksheets().get(0).getShapes()
				.addShape(MsoDrawingType.ARC, 2, 2, 0, 0, 130, 130);
		// Set the placement of the arc.
		arc1.setPlacement(PlacementType.FREE_FLOATING);
		// Set the fill format.
		MsoFillFormat fillformat = arc1.getFillFormat();
		fillformat.setForeColor(Color.getBlue());
		// Set the line style.
		MsoLineFormat lineformat = arc1.getLineFormat();
		lineformat.setStyle(MsoLineStyle.SINGLE);
		// Set the line weight.
		lineformat.setWeight(1);
		// Set the color of the arc line.
		lineformat.setForeColor(Color.getBlue());
		// Set the dash style of the arc.
		lineformat.setDashStyle(MsoLineDashStyle.SOLID);
		// Add another arc shape.
		com.aspose.cells.ArcShape arc2 = (com.aspose.cells.ArcShape) excelbook.getWorksheets().get(0).getShapes()
				.addShape(MsoDrawingType.ARC, 9, 2, 0, 0, 130, 130);
		// Set the placement of the arc.
		arc2.setPlacement(PlacementType.FREE_FLOATING);
		// Set the line style.
		MsoLineFormat lineformat1 = arc2.getLineFormat();
		lineformat1.setStyle(MsoLineStyle.SINGLE);
		// Set the line weight.
		lineformat1.setWeight(1);
		// Set the color of the arc line.
		lineformat1.setForeColor(Color.getBlue());
		// Set the dash style of the arc.
		lineformat1.setDashStyle(MsoLineDashStyle.SOLID);
		// Save the excel file.
		excelbook.save(dataDir + "AAControl-out.xls");
	}
}
