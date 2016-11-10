package com.aspose.cells.examples.articles;

import com.aspose.cells.ReflectionEffect;
import com.aspose.cells.Shape;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class WorkingWithReflectionEffect {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(WorkingWithReflectionEffect.class) + "articles/";

		// Load your source excel file
		Workbook wb = new Workbook(dataDir + "WorkingWithReflectionEffect_in.xlsx");

		// Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		// Access first shape
		Shape sh = ws.getShapes().get(0);

		// Set the reflection effect of the shape
		// Set its Blur, Size, Transparency and Distance properties
		ReflectionEffect re = sh.getReflection();
		re.setBlur(30);
		re.setSize(90);
		re.setTransparency(0);
		re.setDistance(80);

		// Save the workbook in xlsx format
		wb.save(dataDir + "WorkingWithReflectionEffect_out.xlsx");
	}
}
