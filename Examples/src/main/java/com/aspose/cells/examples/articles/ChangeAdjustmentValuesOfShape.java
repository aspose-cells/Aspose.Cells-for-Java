package com.aspose.cells.examples.articles;

import com.aspose.cells.Shape;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ChangeAdjustmentValuesOfShape {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ChangeAdjustmentValuesOfShape.class) + "articles/";

		// Create workbook object from source excel file
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access first three shapes of the worksheet
		Shape shape1 = worksheet.getShapes().get(0);
		Shape shape2 = worksheet.getShapes().get(1);
		Shape shape3 = worksheet.getShapes().get(2);

		// Change the adjustment values of the shapes
		shape1.getGeometry().getShapeAdjustValues().get(0).setValue(0.5d);
		shape2.getGeometry().getShapeAdjustValues().get(0).setValue(0.8d);
		shape3.getGeometry().getShapeAdjustValues().get(0).setValue(0.5d);

		// Save the workbook
		workbook.save(dataDir + "CAVOfShape_out.xlsx");

	}
}
