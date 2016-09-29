package com.aspose.cells.examples.articles;

import com.aspose.cells.ShapeCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class CopyWorksheetBetweenWorkbooks {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CopyWorksheetBetweenWorkbooks.class) + "articles/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "Controls.xls");

		WorksheetCollection ws = workbook.getWorksheets();
		Worksheet sheet1 = ws.get("Control");
		Worksheet sheet2 = ws.get("Result");

		// Get the Shapes from the "Control" worksheet.
		ShapeCollection shapes = sheet1.getShapes();

		// Copy the Textbox to Second Worksheet
		sheet2.getShapes().addCopy(shapes.get(0), 5, 0, 2, 0);

		// Copy the oval shape to Second Worksheet
		sheet2.getShapes().addCopy(shapes.get(1), 10, 0, 2, 0);

		// Save the workbook
		workbook.save(dataDir + "CWBetweenWorkbooks_out.xls");

	}
}
