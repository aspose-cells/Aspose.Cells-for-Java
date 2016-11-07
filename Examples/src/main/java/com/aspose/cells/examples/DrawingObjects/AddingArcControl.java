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
		FillFormat fillformat = arc1.getFill();// getFillFormat();
		
		fillformat.setOneColorGradient(Color.getLime(), 1, GradientStyleType.HORIZONTAL, 1);  
		//setForeColor(Color.getBlue());
		
		// Set the line style.
		LineFormat lineformat = arc1.getLine();// getLineFormat();
		lineformat.setDashStyle(MsoLineStyle.SINGLE); //setStyle(MsoLineStyle.SINGLE);
		// Set the line weight.
		lineformat.setWeight(1);
		// Set the color of the arc line.
		lineformat.setOneColorGradient(Color.getLime(), 1, GradientStyleType.HORIZONTAL, 1); //setForeColor(Color.getBlue());
		// Set the dash style of the arc.
		lineformat.setDashStyle(MsoLineDashStyle.SOLID);
		
		// Add another arc shape.
		com.aspose.cells.ArcShape arc2 = (com.aspose.cells.ArcShape) excelbook.getWorksheets().get(0).getShapes()
				.addShape(MsoDrawingType.ARC, 9, 2, 0, 0, 130, 130);
		// Set the placement of the arc.
		arc2.setPlacement(PlacementType.FREE_FLOATING);
		// Set the line style.
		LineFormat lineformat1 = arc2.getLine(); //getLineFormat();
		lineformat1.setDashStyle(MsoLineStyle.SINGLE);
		// Set the line weight.
		lineformat1.setWeight(1);
		// Set the color of the arc line.
		lineformat1.setOneColorGradient(Color.getLime(), 1, GradientStyleType.HORIZONTAL, 1);//setForeColor(Color.getBlue());
		// Set the dash style of the arc.
		lineformat1.setDashStyle(MsoLineDashStyle.SOLID);
		// Save the excel file.
		excelbook.save(dataDir + "AddingArcControl_out.xls");
	}
}
