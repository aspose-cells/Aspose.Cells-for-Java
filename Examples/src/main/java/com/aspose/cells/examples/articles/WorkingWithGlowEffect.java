package com.aspose.cells.examples.articles;

import com.aspose.cells.GlowEffect;
import com.aspose.cells.Shape;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class WorkingWithGlowEffect {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(WorkingWithGlowEffect.class) + "articles/";

		// Load your source excel file
		Workbook wb = new Workbook(dataDir + "WorkingWithGlowEffect_in.xlsx");

		// Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		// Access first shape
		Shape sh = ws.getShapes().get(0);

		// Set the glow effect of the shape
		// Set its Size and Transparency properties
		GlowEffect ge = sh.getGlow();
		ge.setSize(30);
		ge.setTransparency(0.4);

		// Save the workbook in xlsx format
		wb.save(dataDir + "WorkingWithGlowEffect_out.xlsx");
	}
}
