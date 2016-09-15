package com.aspose.cells.examples.DrawingObjects.controls;

import com.aspose.cells.Color;
import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.MsoFillFormat;
import com.aspose.cells.MsoLineDashStyle;
import com.aspose.cells.MsoLineFormat;
import com.aspose.cells.MsoLineStyle;
import com.aspose.cells.Oval;
import com.aspose.cells.PlacementType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class AddinganOvalControl {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddinganOvalControl.class) + "DrawingObjects/";
		// Instantiate a new Workbook.
		Workbook excelbook = new Workbook();

		// Add an oval shape.
		Oval oval1 = (com.aspose.cells.Oval) excelbook.getWorksheets().get(0).getShapes().addShape(MsoDrawingType.OVAL,
				2, 2, 0, 0, 130, 130);

		// Set the placement of the oval.
		oval1.setPlacement(PlacementType.FREE_FLOATING);

		// Set the fill format.
		MsoFillFormat fillformat = oval1.getFillFormat();
		fillformat.setForeColor(Color.getNavy());

		// Set the line style.
		MsoLineFormat lineformat = oval1.getLineFormat();
		lineformat.setStyle(MsoLineStyle.SINGLE);

		// Set the line weight.
		lineformat.setWeight(1);

		// Set the color of the oval line.
		lineformat.setForeColor(Color.getGreen());

		// Set the dash style of the oval.
		lineformat.setDashStyle(MsoLineDashStyle.SOLID);

		// Add another arc shape.
		Oval oval2 = (com.aspose.cells.Oval) excelbook.getWorksheets().get(0).getShapes().addShape(MsoDrawingType.OVAL,
				9, 2, 0, 0, 130, 130);

		// Set the placement of the oval.
		oval2.setPlacement(PlacementType.FREE_FLOATING);

		// Set the line style.
		MsoLineFormat lineformat1 = oval2.getLineFormat();
		lineformat1.setStyle(MsoLineStyle.SINGLE);

		// Set the line weight.
		lineformat1.setWeight(1);

		// Set the color of the oval line.
		lineformat1.setForeColor(Color.getBlue());

		// Set the dash style of the oval.
		lineformat1.setDashStyle(MsoLineDashStyle.SOLID);

		// Save the excel file.
		excelbook.save(dataDir + "AanOControl-out.xls");
	}
}
