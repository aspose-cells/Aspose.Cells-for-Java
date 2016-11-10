package com.aspose.cells.examples.articles;

import com.aspose.cells.ShadowEffect;
import com.aspose.cells.Shape;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class WorkingWithShadowEffect {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(WorkingWithShadowEffect.class) + "articles/";

		// Loads the workbook which contains hidden external links
		Workbook wb = new Workbook(dataDir + "WorkingWithShadowEffect_in.xlsx");
				
		// Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		// Access first shape
		Shape sh = ws.getShapes().get(0);

		// Set the shadow effect of the shape
		// Set its Angle, Blur, Distance and Transparency properties
		ShadowEffect se = sh.getShadowEffect();
		se.setAngle(150);
		se.setBlur(4);
		se.setDistance(45);
		se.setTransparency(0.3);

		// Save the workbook in xlsx format
		wb.save(dataDir + "WorkingWithShadowEffect_out.xlsx");
	}
}
