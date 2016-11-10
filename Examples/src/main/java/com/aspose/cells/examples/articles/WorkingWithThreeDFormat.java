package com.aspose.cells.examples.articles;

import com.aspose.cells.BevelType;
import com.aspose.cells.Shape;
import com.aspose.cells.ThreeDFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class WorkingWithThreeDFormat {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(WorkingWithThreeDFormat.class) + "articles/";
		
		//Load excel file containing a shape
		Workbook wb = new Workbook(dataDir + "WorkingWithThreeDFormat_in.xlsx");

		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		//Access first shape
		Shape sh = ws.getShapes().get(0);

		//Apply different three dimensional settings
		ThreeDFormat n3df =  sh.getThreeDFormat();
		n3df.setContourWidth(17);
		n3df.setExtrusionHeight(32);	
		n3df.setTopBevelType(BevelType.HARD_EDGE);
		n3df.setTopBevelWidth (30);
		n3df.setTopBevelHeight(30);

		//Save the output excel file in xlsx format
		wb.save(dataDir + "WorkingWithThreeDFormat_out.xlsx");
	}
}
